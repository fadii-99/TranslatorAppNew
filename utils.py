import os
import zipfile
import shutil
import tempfile
import re
import pandas as pd
from docx import Document
from langchain_openai import ChatOpenAI
from lxml import etree
from dotenv import load_dotenv
from langchain_core.prompts import PromptTemplate
import json

load_dotenv()

RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish",
    "Pashto", "Sindhi", "Dhivehi", "Kurdish", "ur", "ar"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language.lower()
        self.source_lang = 'English'
        self.extract_folder = tempfile.mkdtemp(prefix="docx_extract_")
        self.word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        self.translations_cache = {}
        self.used_glossary_words = set()
        self.used_glossary_pairs = set()
        self.nonglossary_pairs = set()
        self.word_report = []
        self.glossary_file_path = glossary_file_path
        self.OPENAI_API_KEY = OPENAI_API_KEY

    def normalize_text(self, text):
        """Normalize text: lowercase, remove extra spaces, keep spaces for phrases."""
        if not text:
            return text
        text = re.sub(r'\s+', ' ', text.strip())
        return text.lower()

    def read_document(self, file_path):
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            return '\n'.join([para.text for para in doc.paragraphs if para.text.strip()])
        else:
            raise ValueError("Only .docx supported.")

    def load_glossary(self):
        if not self.glossary_file_path or not os.path.exists(self.glossary_file_path):
            print("No glossary file provided or file does not exist.")
            return {}
        ext = os.path.splitext(self.glossary_file_path)[1].lower()
        try:
            if ext == ".csv":
                df = pd.read_csv(self.glossary_file_path)
            elif ext in [".xlsx", ".xls"]:
                df = pd.read_excel(self.glossary_file_path)
            else:
                print("Unsupported glossary file format. Expected .csv or .xlsx.")
                return {}
            # Look for English and target language columns
            word_col = None
            trans_col = None
            for col in df.columns:
                col_lower = col.strip().lower()
                if col_lower in ["word", "english"]:
                    word_col = col
                if col_lower == self.target_language or col_lower in ["translation", "arabic", "urdu"]:
                    trans_col = col
            if not word_col or not trans_col:
                print(f"Glossary columns not found. Expected 'English' and '{self.target_language}' or 'translation'.")
                return {}
            glossary = {
                str(word).strip().lower(): str(translation).strip()
                for word, translation in zip(df[word_col], df[trans_col])
                if str(word).strip() and str(translation).strip() and pd.notna(word) and pd.notna(translation)
            }
            print(f"Loaded glossary with {len(glossary)} entries: {list(glossary.items())[:5]}...")
            return glossary
        except Exception as e:
            print(f"Error loading glossary: {e}")
            return {}

    def get_used_glossary_words(self):
        return list(self.used_glossary_words)

    def get_used_glossary_pairs(self):
        return list(self.used_glossary_pairs)

    def get_nonglossary_pairs(self):
        return list(self.nonglossary_pairs)

    def get_word_report(self):
        return self.word_report

    def extract_docx(self):
        zip_input_file = self.input_file.replace('.docx', '.zip')
        if os.path.exists(zip_input_file):
            os.remove(zip_input_file)
        os.rename(self.input_file, zip_input_file)
        with zipfile.ZipFile(zip_input_file, 'r') as zip_ref:
            zip_ref.extractall(self.extract_folder)
        print("Extracted files and directories:")
        for root, dirs, files in os.walk(self.extract_folder):
            for name in files:
                print(os.path.join(root, name))
        document_xml = os.path.join(self.extract_folder, "word", "document.xml")
        if not os.path.exists(document_xml):
            print("ERROR: document.xml not found after extraction.")
            raise FileNotFoundError(f"document.xml not found in extracted folder: {document_xml}")
        return document_xml, zip_input_file



    def create_translated_docx(self, zip_input_file):
        # Create new zip from the extracted contents
        zip_output_file = self.output_file.replace('.docx', '.zip')
        if os.path.exists(zip_output_file):
            os.remove(zip_output_file)
        with zipfile.ZipFile(zip_output_file, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for foldername, subfolders, filenames in os.walk(self.extract_folder):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, self.extract_folder)
                    docx_zip.write(file_path, arcname)
        # Step: Rename .zip back to .docx
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(zip_output_file, self.output_file)
        # Cleanup: Optionally delete the old input zip if you want
        if os.path.exists(zip_input_file):
            os.remove(zip_input_file)

    def patch_styles_for_urdu(self):
        styles_path = os.path.join(self.extract_folder, "word", "styles.xml")
        if not os.path.exists(styles_path):
            print("styles.xml not found, skipping style patch.")
            return
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(styles_path, parser)
        root = tree.getroot()
        ns = {'w': self.word_ns}

        # Patch Normal and Heading1 styles
        for style_id in ["Normal", "Heading1"]:
            style = root.find(f'.//w:style[@w:styleId="{style_id}"]', namespaces=ns)
            if style is not None:
                # Paragraph props
                pPr = style.find('w:pPr', namespaces=ns)
                if pPr is None:
                    pPr = etree.SubElement(style, f'{{{self.word_ns}}}pPr')
                # Right align
                jc = pPr.find('w:jc', namespaces=ns)
                if jc is None:
                    jc = etree.SubElement(pPr, f'{{{self.word_ns}}}jc')
                jc.set(f'{{{self.word_ns}}}val', 'right')
                # Bidi
                bidi = pPr.find('w:bidi', namespaces=ns)
                if bidi is None:
                    bidi = etree.SubElement(pPr, f'{{{self.word_ns}}}bidi')
                bidi.set(f'{{{self.word_ns}}}val', "1")
                # Run props
                rPr = style.find('w:rPr', namespaces=ns)
                if rPr is None:
                    rPr = etree.SubElement(style, f'{{{self.word_ns}}}rPr')
                rFonts = rPr.find('w:rFonts', namespaces=ns)
                if rFonts is None:
                    rFonts = etree.SubElement(rPr, f'{{{self.word_ns}}}rFonts')
                rFonts.set(f'{{{self.word_ns}}}ascii', 'Jameel Noori Nastaleeq')
                rFonts.set(f'{{{self.word_ns}}}hAnsi', 'Jameel Noori Nastaleeq')
                rFonts.set(f'{{{self.word_ns}}}cs', 'Jameel Noori Nastaleeq')
                rFonts.set(f'{{{self.word_ns}}}bidi', 'Jameel Noori Nastaleeq')
                lang = rPr.find('w:lang', namespaces=ns)
                if lang is None:
                    lang = etree.SubElement(rPr, f'{{{self.word_ns}}}lang')
                lang.set(f'{{{self.word_ns}}}val', 'ur-PK')
                lang.set(f'{{{self.word_ns}}}bidi', 'ur-PK')
        tree.write(styles_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        print("styles.xml patched for Urdu RTL.")


    def patch_settings_for_urdu(self):
        settings_path = os.path.join(self.extract_folder, "word", "settings.xml")
        if not os.path.exists(settings_path):
            print("settings.xml not found, skipping settings patch.")
            return
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(settings_path, parser)
        root = tree.getroot()
        lang = root.find(f'.//{{{self.word_ns}}}lang')
        if lang is None:
            lang = etree.SubElement(root, f'{{{self.word_ns}}}lang')
        lang.set(f'{{{self.word_ns}}}val', 'ur-PK')
        lang.set(f'{{{self.word_ns}}}bidi', 'ur-PK')
        compat = root.find(f'{{{self.word_ns}}}compat')
        if compat is None:
            compat = etree.SubElement(root, f'{{{self.word_ns}}}compat')
        cs = compat.find(f'{{{self.word_ns}}}compatSetting[@w:name="useFELayout"]', namespaces={'w': self.word_ns})
        if cs is None:
            cs = etree.SubElement(compat, f'{{{self.word_ns}}}compatSetting')
            cs.set(f'{{{self.word_ns}}}name', 'useFELayout')
            cs.set(f'{{{self.word_ns}}}uri', 'http://schemas.microsoft.com/office/word')
        cs.set(f'{{{self.word_ns}}}val', '1')
        tree.write(settings_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        print("settings.xml patched for Urdu RTL.")


    def fix_bidi_lang_in_styles(self):
        styles_path = os.path.join(self.extract_folder, "word", "styles.xml")
        if not os.path.exists(styles_path):
            print("styles.xml not found.")
            return
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(styles_path, parser)
        root = tree.getroot()
        changed = False
        for lang_elem in root.iter(f'{{{self.word_ns}}}lang'):
            val = lang_elem.get(f'{{{self.word_ns}}}val')
            bidi = lang_elem.get(f'{{{self.word_ns}}}bidi')
            # If it's set to en-US/ar-SA, change it to ur-PK/ur-PK
            if val == "en-US" and bidi == "ar-SA":
                lang_elem.set(f'{{{self.word_ns}}}val', "ur-PK")
                lang_elem.set(f'{{{self.word_ns}}}bidi', "ur-PK")
                changed = True
        # As a fallback: if any <w:lang> has bidi="ar-SA", set to ur-PK
        for lang_elem in root.iter(f'{{{self.word_ns}}}lang'):
            if lang_elem.get(f'{{{self.word_ns}}}bidi') == "ar-SA":
                lang_elem.set(f'{{{self.word_ns}}}bidi', "ur-PK")
                changed = True
        if changed:
            tree.write(styles_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
            print("styles.xml bidi language fixed for Urdu.")
        else:
            print("No bidi language to fix in styles.xml.")


    def force_all_rtl(self, root):
    # For every paragraph
        for p in root.iter(f'{{{self.word_ns}}}p'):
            # Paragraph properties
            pPr = p.find(f'{{{self.word_ns}}}pPr')
            if pPr is None:
                pPr = etree.SubElement(p, f'{{{self.word_ns}}}pPr')

            # Set right alignment
            jc = pPr.find(f'{{{self.word_ns}}}jc')
            if jc is None:
                jc = etree.SubElement(pPr, f'{{{self.word_ns}}}jc')
            jc.set(f'{{{self.word_ns}}}val', 'right')

            # Set paragraph bidi (RTL)
            bidi = pPr.find(f'{{{self.word_ns}}}bidi')
            if bidi is None:
                bidi = etree.SubElement(pPr, f'{{{self.word_ns}}}bidi')
            bidi.set(f'{{{self.word_ns}}}val', "1")

            # Remove textDirection if present
            textDirection = pPr.find(f'{{{self.word_ns}}}textDirection')
            if textDirection is not None:
                pPr.remove(textDirection)

            # For every run in this paragraph
            for r in p.findall(f'.//{{{self.word_ns}}}r'):
                rPr = r.find(f'{{{self.word_ns}}}rPr')
                if rPr is None:
                    rPr = etree.SubElement(r, f'{{{self.word_ns}}}rPr')

                # Set run RTL
                rtl = rPr.find(f'{{{self.word_ns}}}rtl')
                if rtl is None:
                    rtl = etree.SubElement(rPr, f'{{{self.word_ns}}}rtl')
                rtl.set(f'{{{self.word_ns}}}val', "1")

                # Set bidi on run
                rbidi = rPr.find(f'{{{self.word_ns}}}bidi')
                if rbidi is None:
                    rbidi = etree.SubElement(rPr, f'{{{self.word_ns}}}bidi')
                rbidi.set(f'{{{self.word_ns}}}val', "1")

                # Set language and font for Urdu
                lang = rPr.find(f'{{{self.word_ns}}}lang')
                if lang is None:
                    lang = etree.SubElement(rPr, f'{{{self.word_ns}}}lang')
                lang.set(f'{{{self.word_ns}}}val', 'ur-PK')
                lang.set(f'{{{self.word_ns}}}bidi', 'ur-PK')
                font = rPr.find(f'{{{self.word_ns}}}rFonts')
                if font is None:
                    font = etree.SubElement(rPr, f'{{{self.word_ns}}}rFonts')
                font.set(f'{{{self.word_ns}}}ascii', 'Jameel Noori Nastaleeq')
                font.set(f'{{{self.word_ns}}}hAnsi', 'Jameel Noori Nastaleeq')
                font.set(f'{{{self.word_ns}}}cs', 'Jameel Noori Nastaleeq')
                font.set(f'{{{self.word_ns}}}bidi', 'Jameel Noori Nastaleeq')


    def _check_glossary_in_text(self, text, glossary):
        """Check for glossary words/phrases in text and track usage."""
        matches = []
        text_lower = self.normalize_text(text)
        for gloss_word in sorted(glossary, key=lambda x: len(x), reverse=True):
            gloss_word_lower = self.normalize_text(gloss_word)
            # Use word boundaries for English, relaxed boundaries for Arabic/Urdu
            if self.target_language in ['ar', 'urdu']:
                pattern = rf'(?<!\w){re.escape(gloss_word_lower)}(?!\w)'
            else:
                pattern = rf'\b{re.escape(gloss_word_lower)}\b'
            if re.search(pattern, text_lower, re.IGNORECASE):
                self.used_glossary_words.add(gloss_word)
                self.used_glossary_pairs.add((gloss_word, glossary[gloss_word]))
                matches.append((gloss_word, glossary[gloss_word]))
                self.word_report.append({
                    'Original Word': gloss_word,
                    'Translated Word': glossary[gloss_word],
                    'Source': 'Glossary',
                    'Used in Document': 'Yes'
                })
        print(f"Glossary matches for text '{text[:50]}...': {matches}")
        return matches

    def _tokenize_words(self, text, glossary):
        """Tokenize text into words, preserving multi-word glossary phrases."""
        words = []
        text_lower = self.normalize_text(text)
        original_text = text
        matched_positions = []
        for gloss_word in sorted(glossary, key=lambda x: len(x), reverse=True):
            gloss_word_lower = self.normalize_text(gloss_word)
            if self.target_language in ['ar', 'urdu']:
                pattern = rf'(?<!\w){re.escape(gloss_word_lower)}(?!\w)'
            else:
                pattern = rf'\b{re.escape(gloss_word_lower)}\b'
            for match in re.finditer(pattern, text_lower, re.IGNORECASE):
                start, end = match.span()
                matched_positions.append((start, end, gloss_word, glossary[gloss_word]))
        matched_positions.sort()
        last_end = 0
        for start, end, gloss_word, translation in matched_positions:
            before_text = original_text[last_end:start]
            if before_text:
                words.extend(re.findall(r'\S+', before_text))
            words.append(gloss_word)
            last_end = end
        if last_end < len(original_text):
            words.extend(re.findall(r'\S+', original_text[last_end:]))
        return words

    def translate_text(self, text):
        if not text.strip():
            return text, [], []
        glossary = self.load_glossary()

        # Step 1: Check for glossary matches and track them
        matches = self._check_glossary_in_text(text, glossary)

        # Step 2: Replace glossary terms in the text
        translated_text = text
        for gloss_word in sorted(glossary, key=lambda x: len(x), reverse=True):
            pattern = rf'(?<!\S){re.escape(gloss_word)}(?!\S)'
            translated_text = re.sub(pattern, glossary[gloss_word], translated_text, flags=re.IGNORECASE)

        # Step 3: Tokenize the original text to track non-glossary words
        words = self._tokenize_words(text, glossary)
        non_glossary_words = [w for w in words if w.lower() not in glossary]

        # Step 4: Translate non-glossary words using LLM and get lists
        local_glossary_pairs = []
        local_nonglossary_pairs = []
        if non_glossary_words:
            prompt = PromptTemplate.from_template(
                """
                You are a professional translator. Translate the given text from {source_language} to {output_language}.
                Your translations must be precise, accurate, and natural-sounding in the target language.

                Reference Glossary (use these translations if the word matches exactly, case-insensitive):
                {glossary_context}

                Text to translate ({source_language}):
                {input}

                Your response **must** be a valid JSON object with the following structure:
                {{
                    "translated_text": "translated text here",
                    "glossary_pairs": [["original word", "translated word"], ...],
                    "nonglossary_pairs": [["original word", "translated word"], ...]
                }}
                - "translated_text": The fully translated text. If no translation is needed, return the input text.
                - "glossary_pairs": List of [original, translated] pairs for words found in the glossary. Return an empty list if none apply.
                - "nonglossary_pairs": List of [original, translated] pairs for words translated by you. Return an empty list if none apply.
                If the input text is empty or untranslatable, return an empty JSON object with the above structure.
                Ensure the response is valid JSON with proper syntax.
                """
            )
            glossary_context = "\n".join([f"{word}: {translation}" for word, translation in glossary.items()])
            llm = ChatOpenAI(model_name="gpt-4o", api_key=self.OPENAI_API_KEY, temperature=0.3)
            chain = prompt | llm

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    target_lang = self.target_language
                    if target_lang == "ar":
                        target_lang = "Arabic"
                    elif target_lang == "ur":
                        target_lang = "Urdu"
                    response = chain.invoke({
                        "source_language": self.source_lang,
                        "output_language": target_lang,
                        "input": text,
                        "glossary_context": glossary_context
                    })
                    # Log raw response for debugging
                    print(f"LLM Response (Attempt {attempt + 1}): {response.content}")
                    
                    # Clean response to handle BOM or encoding issues
                    response_content = response.content.strip()
                    if not response_content:
                        print("Translation error: Empty LLM response")
                        break
                    
                    # Remove BOM if present
                    response_content = response_content.encode('utf-8').decode('utf-8-sig')
                    
                    # Extract JSON from Markdown code block if present
                    json_match = re.search(r'```json\n([\s\S]*?)\n```', response_content)
                    if json_match:
                        response_content = json_match.group(1).strip()
                    else:
                        # If no Markdown wrapper, assume response is raw JSON
                        response_content = response_content.strip()
                    
                    # Log the exact string being parsed
                    print(f"Content to parse (Attempt {attempt + 1}): {repr(response_content)}")
                    
                    # Parse JSON response
                    result = json.loads(response_content)
                    
                    # Validate JSON structure
                    if not isinstance(result, dict) or "translated_text" not in result:
                        print("Translation error: Invalid JSON structure")
                        break
                    
                    translated_text = result.get("translated_text", translated_text)
                    local_glossary_pairs = result.get("glossary_pairs", [])
                    local_nonglossary_pairs = result.get("nonglossary_pairs", [])

                    # Validate glossary pairs against loaded glossary
                    valid_glossary_pairs = [(orig, trans) for orig, trans in local_glossary_pairs if orig.lower() in glossary and glossary[orig.lower()] == trans]
                    for orig, trans in valid_glossary_pairs:
                        if trans.strip() and trans != orig:
                            self.used_glossary_pairs.add((orig, trans))
                            self.used_glossary_words.add(orig)
                            if not any(r['Original Word'] == orig and r['Source'] == 'Glossary' for r in self.word_report):
                                self.word_report.append({
                                    'Original Word': orig,
                                    'Translated Word': trans,
                                    'Source': 'Glossary',
                                    'Used in Document': 'Yes'
                                })
                    for orig, trans in local_nonglossary_pairs:
                        if trans.strip() and trans != orig:
                            self.nonglossary_pairs.add((orig, trans))
                            self.word_report.append({
                                'Original Word': orig,
                                'Translated Word': trans,
                                'Source': 'LLM',
                                'Used in Document': 'Yes'
                            })
                    break  # Success, exit retry loop
                except json.JSONDecodeError as e:
                    print(f"Translation error: JSON parsing failed: {e}")
                    print(f"Failed content: {repr(response_content)}")
                    if attempt == max_retries - 1:
                        print("Max retries reached. Using glossary-based translation as fallback.")
                        break
                except Exception as e:
                    print(f"Translation error: {e}")
                    print(f"Failed content: {repr(response_content)}")
                    if attempt == max_retries - 1:
                        print("Max retries reached. Using glossary-based translation as fallback.")
                        break

        return translated_text, local_glossary_pairs, local_nonglossary_pairs

    def translate_xml_to_language(self, xml_path, source_lang="en", target_lang="ur", output_path=None):
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_path, parser)
        root = tree.getroot()

        # 1. Translate each paragraph's text
        for element in root.iter(f'{{{self.word_ns}}}p'):
            paragraph_text = ''
            text_elements = []
            for t in element.iter(f'{{{self.word_ns}}}t'):
                if t.text and t.text.strip():
                    paragraph_text += t.text
                    text_elements.append((t, 'text'))
                if t.tail and t.tail.strip():
                    paragraph_text += t.tail
                    text_elements.append((t, 'tail'))
            if paragraph_text.strip():
                translated_text, gloss_pairs, nongloss_pairs = self.translate_text(paragraph_text)
                for t, t_type in text_elements:
                    if t_type == 'text' and t.text.strip():
                        t.text = translated_text
                        translated_text = ''
                    elif t_type == 'tail' and t.tail.strip():
                        t.tail = translated_text
                        translated_text = ''

        # 2. Force RTL and right-aligned formatting if needed
        if target_lang.lower() in [lang.lower() for lang in RTL_LANGUAGES]:
            self.force_all_rtl(root)
            self.patch_styles_for_urdu()
            self.patch_settings_for_urdu()
            self.fix_bidi_lang_in_styles()


        # 4. Save output
        if output_path:
            tree.write(output_path, encoding='utf-8', xml_declaration=True, pretty_print=True)


    def generate_usage_report(self):
        glossary = self.load_glossary()
        used_words = set(self.used_glossary_words)
        for word, translation in glossary.items():
            if word not in used_words and not any(
                r['Original Word'] == word and r['Source'] == 'Glossary' for r in self.word_report
            ):
                self.word_report.append({
                    'Original Word': word,
                    'Translated Word': translation,
                    'Source': 'Glossary',
                    'Used in Document': 'No'
                })
        return self.word_report

    def run(self):
        if os.path.exists(self.extract_folder):
            try:
                shutil.rmtree(self.extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing folder: {e}")
        os.makedirs(self.extract_folder, exist_ok=True)

        try:
            xml_path, zip_input_file = self.extract_docx()
            self.translate_xml_to_language(xml_path, source_lang="en", target_lang=self.target_language, output_path=xml_path)
            self.create_translated_docx(zip_input_file)
            report = self.generate_usage_report()
            report_df = pd.DataFrame(report)
            report_df.to_csv("word_by_word_report.csv", index=False, encoding='utf-8')
            print(f"Translation complete! Saved as: {self.output_file}")
            print(f"Word-by-word report saved as: word_by_word_report.csv")
            return report, self.get_used_glossary_pairs(), self.get_nonglossary_pairs()
        except Exception as e:
            print(f"Error during DOCX translation: {e}")
            return [], [], []
        finally:
            if os.path.exists(self.extract_folder):
                try:
                    shutil.rmtree(self.extract_folder)
                except Exception as e:
                    print(f"Warning: Could not clean up temporary folder: {e}")

def translate_file(input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY):
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension == '.docx':
        translator = DocxTranslator(input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx")
    report, used_glossary_pairs, nonglossary_pairs = translator.run()
    return report, used_glossary_pairs, nonglossary_pairs