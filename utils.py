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

load_dotenv()

RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish",
    "Pashto", "Sindhi", "Dhivehi", "Kurdish", "ur", "ar"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.source_lang = 'English'
        self.extract_folder = tempfile.mkdtemp(prefix="docx_extract_")
        self.word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        self.translations_cache = {}
        self.used_glossary_words = set()
        self.used_glossary_pairs = set()
        self.nonglossary_pairs = set()
        self.glossary_file_path = glossary_file_path
        self.OPENAI_API_KEY = OPENAI_API_KEY

    def read_document(self, file_path):
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            return '\n'.join([para.text for para in doc.paragraphs])
        else:
            raise ValueError("Only .docx supported.")

    def load_glossary(self):
        if not self.glossary_file_path or not os.path.exists(self.glossary_file_path):
            return {}
        ext = os.path.splitext(self.glossary_file_path)[1].lower()
        try:
            if ext == ".csv":
                df = pd.read_csv(self.glossary_file_path)
            elif ext in [".xlsx", ".xls"]:
                df = pd.read_excel(self.glossary_file_path)
            else:
                return {}
            # Must have columns: "Word" and "[LANG Translation]" (like "ARABIC Translation")
            # Try to detect the translation column
            word_col = [col for col in df.columns if col.strip().lower() == "word"]
            trans_col = [col for col in df.columns if self.target_language.upper() in col.upper()]
            if not word_col or not trans_col:
                return {}
            # Ensure strings for matching
            return dict(zip(df[word_col[0]].astype(str).str.strip(), df[trans_col[0]].astype(str).str.strip()))
        except Exception as e:
            print(f"Error loading uploaded glossary: {e}")
            return {}

    def get_used_glossary_words(self):
        return list(self.used_glossary_words)

    def get_used_glossary_pairs(self):
        return list(self.used_glossary_pairs)

    def get_nonglossary_pairs(self):
        return list(self.nonglossary_pairs)

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

                Use the glossary for any word if present. If not, always translate yourself. **Never reply in English, never ask for original text, and never apologize.** Every input must be translated. Do not skip or refuse.

                Reference Glossary:
                {glossary_context}

                Original text ({source_language}):
                {input}

                Your response **must** follow this exact format:

                Translated text ({output_language}): <translation_here>

                List 1 (words used from glossary):
                word 1, translation
                word 2, translation

                List 2 (words translated but not in glossary):
                word 1, translation
                word 2, translation
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

                # Robust parsing of the output format
                m = re.search(
                    r"Translated text.*?:\s*(.*?)\n+List 1.*?:\s*(.*?)(?:\n+List 2.*?:\s*(.*))?$",
                    resp, re.DOTALL
                )
                if m:
                    translated_text = m.group(1).strip()
                    glossary_list_raw = m.group(2).strip()
                    nonglossary_list_raw = m.group(3).strip() if m.group(3) else ""

                    glossary_pairs = []
                    for line in glossary_list_raw.split('\n'):
                        if ',' in line:
                            eng, trans = [x.strip() for x in line.split(',', 1)]
                            if eng and trans:
                                glossary_pairs.append((eng, trans))
                    nonglossary_pairs = []
                    for line in nonglossary_list_raw.split('\n'):
                        if ',' in line:
                            eng, trans = [x.strip() for x in line.split(',', 1)]
                            if eng and trans:
                                nonglossary_pairs.append((eng, trans))
                    for pair in glossary_pairs:
                        self.used_glossary_words.add(pair[0])
                        self.used_glossary_pairs.add(pair)
                    for pair in nonglossary_pairs:
                        self.nonglossary_pairs.add(pair)
                else:
                    # If parsing fails, try to extract Urdu/Arabic text chunks
                    translated_text = resp
                    if self.target_language.lower() == 'ur':
                        urdu_text_match = re.findall(r'[\u0600-\u06FF\s,.\-؛؟]+', resp)
                        if urdu_text_match:
                            translated_text = ' '.join(urdu_text_match)
                    # Add Arabic support if needed
                    if self.target_language.lower() == 'ar':
                        arabic_text_match = re.findall(r'[\u0600-\u06FF\s,.\-؛؟]+', resp)
                        if arabic_text_match:
                            translated_text = ' '.join(arabic_text_match)

                # If translation is an unwanted English prompt, fallback
                unwanted_phrases = [
                    "Certainly! Please provide the original text",
                    "please provide the original text",
                    "I'm sorry",
                    "cannot translate",
                    "unable to translate",
                    "provide more information",
                    "not enough information",
                    "please provide"
                ]
                if any(phrase.lower() in translated_text.lower() for phrase in unwanted_phrases):
                    translated_text = text  # Fallback: just use the original text (or try translating with a different engine, if you want)

            except Exception as e:
                print(f"Translation error: {e}")
                return text

        else:
            print("No OpenAI API key provided.")
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

        # RTL formatting
        if target_lang.lower() in [lang.lower() for lang in RTL_LANGUAGES]:
            for p in root.findall('.//{%s}p' % self.word_ns):
                pPr = p.find('{%s}pPr' % self.word_ns)
                if pPr is None:
                    pPr = etree.SubElement(p, '{%s}pPr' % self.word_ns)
                if pPr.find('{%s}bidi' % self.word_ns) is None:
                    bidi = etree.SubElement(pPr, '{%s}bidi' % self.word_ns)
                jc = pPr.find('{%s}jc' % self.word_ns)
                if jc is None:
                    jc = etree.SubElement(pPr, '{%s}jc' % self.word_ns)
                jc.set('{%s}val' % self.word_ns, 'right')
                for r in p.findall('.//{%s}r' % self.word_ns):
                    rPr = r.find('{%s}rPr' % self.word_ns)
                    if rPr is None:
                        rPr = etree.SubElement(r, '{%s}rPr' % self.word_ns)
                    if rPr.find('{%s}bidi' % self.word_ns) is None:
                        bidi = etree.SubElement(rPr, '{%s}bidi' % self.word_ns)
                    lang = rPr.find('{%s}lang' % self.word_ns)
                    if lang is None:
                        lang = etree.SubElement(rPr, '{%s}lang' % self.word_ns)
                    lang.set('{%s}val' % self.word_ns, 'ur-PK' if target_lang.lower() == 'ur' else 'ar-SA')
                    font = rPr.find('{%s}rFonts' % self.word_ns)
                    if font is None:
                        font = etree.SubElement(rPr, '{%s}rFonts' % self.word_ns)
                    if target_lang.lower() == 'ur':
                        font.set('{%s}ascii' % self.word_ns, 'Jameel Noori Nastaleeq')
                        font.set('{%s}hAnsi' % self.word_ns, 'Jameel Noori Nastaleeq')
                    else:
                        font.set('{%s}ascii' % self.word_ns, 'Amiri')
                        font.set('{%s}hAnsi' % self.word_ns, 'Amiri')

        if target_lang.lower() in [lang.lower() for lang in RTL_LANGUAGES]:
            settings_path = os.path.join(self.extract_folder, "word", "settings.xml")
            if os.path.exists(settings_path):
                settings_tree = etree.parse(settings_path, parser)
                settings_root = settings_tree.getroot()
                lang = settings_root.find('.//{%s}lang' % self.word_ns)
                if lang is None:
                    lang = etree.SubElement(settings_root, '{%s}lang' % self.word_ns)
                lang.set('{%s}val' % self.word_ns, 'ur-PK' if target_lang.lower() == 'ur' else 'ar-SA')
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

def translate_file(input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY):
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension == '.docx':
        translator = DocxTranslator(input_file, output_file, target_language, glossary_file_path, OPENAI_API_KEY)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx")
    translator.run()
    used_glossary = translator.get_used_glossary_words()
    used_glossary_pairs = translator.get_used_glossary_pairs()
    nonglossary_pairs = translator.get_nonglossary_pairs()
    return used_glossary, used_glossary_pairs, nonglossary_pairs
