import os
import zipfile
import shutil
import tempfile
import re
from collections import Counter
import pandas as pd
from docx import Document
from langchain_openai import ChatOpenAI
from lxml import etree
from dotenv import load_dotenv
from modernmt import ModernMT
from langchain_core.prompts import PromptTemplate

load_dotenv()

RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish",
    "Pashto", "Sindhi", "Dhivehi", "Kurdish", "ur", "ar"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.source_lang = 'English'
        self.extract_folder = tempfile.mkdtemp(prefix="docx_extract_")
        self.word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        self.translations_cache = {}  # Cache to store translations
        self.used_glossary_words = set()  # NEW: Track which glossary words were used
        self.used_glossary_pairs = set() 
        # Set glossary file based on target language
        if target_language.lower() == 'ar':
            self.glossary_file = "glossary_ar.xlsx"
        elif target_language.lower() == 'ur':
            self.glossary_file = "glossary_ur.xlsx"
        else:
            raise ValueError(f"Unsupported target language: {target_language}. Only 'ar' and 'ur' are supported.")
        if ModernMT_key:
            self.mmt = ModernMT(ModernMT_key)
            self.OPENAI_API_KEY = None
        elif OPENAI_API_KEY:
            self.mmt = None
            self.OPENAI_API_KEY = OPENAI_API_KEY
        else:
            raise ValueError("Either ModernMT_key or OPENAI_API_KEY must be provided")

    def read_document(self, file_path):
        """Read docx document and return text"""
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            return '\n'.join([para.text for para in doc.paragraphs])
        else:
            raise ValueError("Only .docx supported.")

    def process_text(self, text):
        """Clean text and extract words with frequencies"""
        words = re.findall(r'\b\w+\b', text.lower())
        return Counter(words)

    def load_glossary(self):
        """Load glossary from Excel file if it exists"""
        if os.path.exists(self.glossary_file):
            try:
                df = pd.read_excel(self.glossary_file)
                return dict(zip(df['Word'], df[f'{self.target_language.upper()} Translation']))
            except Exception as e:
                print(f"Error loading glossary {self.glossary_file}: {e}")
                return {}
        return {}
    def get_used_glossary_words(self):
        """Return just the list of English glossary words that were used."""
        return list(self.used_glossary_words)

    def get_used_glossary_pairs(self):
        """Return (word, translation) tuples that were used in the translation."""
        return list(self.used_glossary_pairs)


    def generate_glossary(self):
        """Generate or update the language-specific glossary from the document"""
        # Load existing glossary if it exists
        existing_glossary = {}
        existing_df = None
        if os.path.exists(self.glossary_file):
            try:
                existing_df = pd.read_excel(self.glossary_file)
                existing_glossary = dict(zip(existing_df['Word'], existing_df[f'{self.target_language.upper()} Translation']))
                print(f"Loaded existing glossary with {len(existing_glossary)} entries: {self.glossary_file}")
            except Exception as e:
                print(f"Error loading glossary {self.glossary_file}: {e}")

        # Read and process the document
        text = self.read_document(self.input_file)
        word_counts = self.process_text(text)

        # Process new words
        new_data = []
        for word, count in word_counts.items():
            # Skip if word is in glossary
            if word in existing_glossary:
                continue

            # Translate new words
            translated = self.translate_word(word)
            new_data.append({
                'Word': word,
                'Frequency': count,
                f'{self.target_language.upper()} Translation': translated
            })

        if new_data:
            # Create DataFrame for new entries
            new_df = pd.DataFrame(new_data)
            if existing_df is not None:
                # Append new entries to existing glossary
                updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                # Remove duplicates, keeping the latest entry
                updated_df = updated_df.drop_duplicates(subset=['Word'], keep='last')
                updated_df.sort_values(by='Frequency', ascending=False, inplace=True)
                updated_df.to_excel(self.glossary_file, index=False)
                print(f"Appended {len(new_data)} new words to glossary: {self.glossary_file}")
            else:
                # Create new glossary if none exists
                new_df.sort_values(by='Frequency', ascending=False, inplace=True)
                new_df.to_excel(self.glossary_file, index=False)
                print(f"Created new glossary with {len(new_data)} words: {self.glossary_file}")
        else:
            print(f"No new words to add to glossary: {self.glossary_file}")

    def translate_word(self, word):
        """Translate a single word to target language using OpenAI or ModernMT"""
        if word in self.translations_cache:
            return self.translations_cache[word]
        
        try:
            if self.OPENAI_API_KEY:
                prompt = PromptTemplate.from_template(
                    """
                    You are a professional translator. Translate the given word from English to {target_language}.
                    Provide only the translation without any explanations or quotation marks.
                    
                    English word: {input}
                    {target_language} translation:
                    """
                )
                llm = ChatOpenAI(
                    model_name="gpt-4o",
                    api_key=self.OPENAI_API_KEY,
                    temperature=0.1
                )
                chain = prompt | llm
                response = chain.invoke({"input": word, "target_language": self.target_language})
                translated = response.content.strip()
            elif self.mmt:
                translation = self.mmt.translate("en", self.target_language, word)
                translated = translation.translation
            
            self.translations_cache[word] = translated
            return translated
        except Exception as e:
            print(f"Error translating {word}: {str(e)}")
            return "Translation Error"

    def extract_docx(self):
        with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
            zip_ref.extractall(self.extract_folder)
        return os.path.join(self.extract_folder, "word", "document.xml")

    def create_translated_docx(self):
        base_name = self.output_file.replace('.docx', '')
        shutil.make_archive(base_name, 'zip', self.extract_folder)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(base_name + '.zip', self.output_file)

    def translate_text(self, text):
        glossary = self.load_glossary()
        glossary_context = "\n".join([f"{word}: {translation}" for word, translation in glossary.items()])

        if self.OPENAI_API_KEY:
            prompt = PromptTemplate.from_template(
                """
                You are a professional translator. Translate the given text from {source_language} to {output_language}.
                Your translation must be precise, accurate, fluent, and natural-sounding in the target language, preserving the original meaning.

                When translating, use the glossary as a reference for key terms, but ensure you:
                1. Maintain proper grammatical structure in the target language
                2. Adapt the sentence flow naturally, not word-for-word
                3. Keep the context and meaning of the sentence intact

                Reference Glossary:
                {glossary_context}

                Original text ({source_language}):
                {input}

                Your response **must** follow this exact format:

                Translated text ({output_language}): <translation_here>

                Glossary words used (English, comma-separated): <word1, word2, ...>
                """
            )
            llm = ChatOpenAI(model_name="gpt-4o", api_key=self.OPENAI_API_KEY, temperature=0.3)
            chain = prompt | llm

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
                resp = response.content.strip()

                # PARSE THE RESPONSE
                m1 = re.search(r"Translated text.*?:\s*(.*?)\n+Glossary words used.*?:\s*(.*)", resp, re.DOTALL)
                if m1:
                    translated_text = m1.group(1).strip()
                    glossary_used = m1.group(2).strip()
                    glossary_used_list = [w.strip() for w in glossary_used.split(",") if w.strip()]
                    for word in glossary_used_list:
                        self.used_glossary_words.add(word)
                        if word in glossary:
                            self.used_glossary_pairs.add((word, glossary[word]))
                else:
                    # fallback: return everything, don't break
                    translated_text = resp
            except Exception as e:
                print(f"Translation error: {e}")
                return text

        elif self.mmt:
            # ... (ModernMT logic stays as is)
            try:
                translation = self.mmt.translate("en", self.target_language, text)
                translated_text = translation.translation
            except Exception as e:
                print(f"Error translating text with ModernMT: {e}")
                return text

        return translated_text


    def translate_xml_to_language(self, xml_path, source_lang="en", target_lang="ur", output_path=None):
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_path, parser)
        root = tree.getroot()
        
        for element in root.iter():
            if element.text and element.text.strip():
                original_text = element.text.strip()
                try:
                    translated_text = self.translate_text(original_text)
                    element.text = translated_text
                except Exception as e:
                    print(f"Error translating text: {e}")
            
            if element.tail and element.tail.strip():
                original_tail = element.tail.strip()
                try:
                    translated_tail = self.translate_text(original_tail)
                    element.tail = translated_tail
                except Exception as e:
                    print(f"Error translating tail text: {e}")

        # Apply RTL formatting for RTL languages
        if target_lang.lower() in [lang.lower() for lang in RTL_LANGUAGES]:
            # Update paragraph properties
            for p in root.findall('.//{%s}p' % self.word_ns):
                pPr = p.find('{%s}pPr' % self.word_ns)
                if pPr is None:
                    pPr = etree.SubElement(p, '{%s}pPr' % self.word_ns)
                
                # Add bidi property for RTL text direction
                if pPr.find('{%s}bidi' % self.word_ns) is None:
                    bidi = etree.SubElement(pPr, '{%s}bidi' % self.word_ns)
                
                # Add right alignment
                jc = pPr.find('{%s}jc' % self.word_ns)
                if jc is None:
                    jc = etree.SubElement(pPr, '{%s}jc' % self.word_ns)
                jc.set('{%s}val' % self.word_ns, 'right')

                # Update run properties for all runs in the paragraph
                for r in p.findall('.//{%s}r' % self.word_ns):
                    rPr = r.find('{%s}rPr' % self.word_ns)
                    if rPr is None:
                        rPr = etree.SubElement(r, '{%s}rPr' % self.word_ns)
                    
                    # Add bidi to run properties
                    if rPr.find('{%s}bidi' % self.word_ns) is None:
                        bidi = etree.SubElement(rPr, '{%s}bidi' % self.word_ns)
                    
                    # Set language
                    lang = rPr.find('{%s}lang' % self.word_ns)
                    if lang is None:
                        lang = etree.SubElement(rPr, '{%s}lang' % self.word_ns)
                    lang.set('{%s}val' % self.word_ns, 'ur-PK' if target_lang.lower() == 'ur' else 'ar-SA')
                    
                    # Set language-specific font
                    font = rPr.find('{%s}rFonts' % self.word_ns)
                    if font is None:
                        font = etree.SubElement(rPr, '{%s}rFonts' % self.word_ns)
                    if target_lang.lower() == 'ur':
                        font.set('{%s}ascii' % self.word_ns, 'Jameel Noori Nastaleeq')
                        font.set('{%s}hAnsi' % self.word_ns, 'Jameel Noori Nastaleeq')
                    else:  # ar
                        font.set('{%s}ascii' % self.word_ns, 'Amiri')
                        font.set('{%s}hAnsi' % self.word_ns, 'Amiri')

        # Update document settings for RTL
        if target_lang.lower() in [lang.lower() for lang in RTL_LANGUAGES]:
            settings_path = os.path.join(self.extract_folder, "word", "settings.xml")
            if os.path.exists(settings_path):
                settings_tree = etree.parse(settings_path, parser)
                settings_root = settings_tree.getroot()
                
                # Add default language setting
                lang = settings_root.find('.//{%s}lang' % self.word_ns)
                if lang is None:
                    lang = etree.SubElement(settings_root, '{%s}lang' % self.word_ns)
                lang.set('{%s}val' % self.word_ns, 'ur-PK' if target_lang.lower() == 'ur' else 'ar-SA')
                
                # Add RTL compatibility (optional, for complex scripts)
                cs = settings_root.find('.//{%s}compatSetting[@name="useFELayout"]' % self.word_ns)
                if cs is None:
                    compat = settings_root.find('{%s}compat' % self.word_ns)
                    if compat is None:
                        compat = etree.SubElement(settings_root, '{%s}compat' % self.word_ns)
                    cs = etree.SubElement(compat, '{%s}compatSetting' % self.word_ns)
                    cs.set('{%s}name' % self.word_ns, 'useFELayout')
                    cs.set('{%s}uri' % self.word_ns, 'http://schemas.microsoft.com/office/word')
                    cs.set('{%s}val' % self.word_ns, '1')
                
                settings_tree.write(settings_path, encoding='utf-8', xml_declaration=True, pretty_print=True)

        if output_path:
            tree.write(output_path, encoding='utf-8', xml_declaration=True, pretty_print=True)

    def run(self):
        if os.path.exists(self.extract_folder):
            try:
                shutil.rmtree(self.extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing folder: {e}")
        os.makedirs(self.extract_folder, exist_ok=True)

        try:
            self.generate_glossary()
            xml_path = self.extract_docx()
            self.translate_xml_to_language(xml_path, source_lang="en", target_lang=self.target_language, output_path=xml_path)
            self.create_translated_docx()
            print(f"Translation complete! Saved as: {self.output_file}")
        except Exception as e:
            print(f"Error during DOCX translation: {e}")
        finally:
            if os.path.exists(self.extract_folder):
                try:
                    shutil.rmtree(self.extract_folder)
                except Exception as e:
                    print(f"Warning: Could not clean up temporary folder: {e}")

def translate_file(input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY):
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension == '.docx':
        translator = DocxTranslator(input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx")
    translator.run()
    # NEW: Get the used glossary list
    used_glossary = translator.get_used_glossary_words()
    used_glossary_pairs = translator.get_used_glossary_pairs()
    # You can log it, return it, or save it as needed:
    print("Glossary words used in translation:", used_glossary)
    print("Word/Translation pairs used:", used_glossary_pairs)
    return used_glossary, used_glossary_pairs
