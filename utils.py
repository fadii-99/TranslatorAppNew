# utils.py  –  full module with numbered-list Arabic digit support
import os, zipfile, shutil, tempfile, re, json
import pandas as pd
from lxml import etree
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.prompts import PromptTemplate

load_dotenv()

# ─────────────────────────  RTL helpers ─────────────────────────
RTL_LANGUAGES = {
    "arabic","hebrew","persian","urdu","yiddish",
    "pashto","sindhi","dhivehi","kurdish","ur","ar"
}
def normalize_rtl_punctuation(txt:str)->str:
    mapping={'.':'۔',',':'،',';':'؛','?':'؟','!':'！',
             '(':'﴾',')':'﴿','[':'〚',']':'〛',':':'：',
             '%':'٪','-':'–'}
    for a,b in mapping.items():
        txt=txt.replace(a,b)
    return txt
# ----------------------------------------------------------------

# ───── prefix & digit helpers ─────
RLM='\u200F'
_PREFIX_RE=re.compile(r"""^(\s*(?:[\u2022•\-\u25CF·]|\d+[\.\)\-:])\s*)""",re.VERBOSE)
_ARABIC_DIGITS=str.maketrans("0123456789","٠١٢٣٤٥٦٧٨٩")

def split_prefix(line:str):
    m=_PREFIX_RE.match(line)
    return (m.group(1),line[m.end():].lstrip()) if m else ('',line)

def to_arabic_prefix(p:str)->str:
    return p.translate(_ARABIC_DIGITS).replace('.', '۔')
# ----------------------------------

class DocxTranslator:
    def __init__(self,in_file,out_file,target_lang,glossary_path,key):
        self.in_file,self.out_file=in_file,out_file
        self.target=target_lang.lower()
        self.glossary_path=glossary_path
        self.key=key

        self.extract=tempfile.mkdtemp(prefix="docx_")
        self.word_ns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        self.used_glossary_words=set()
        self.used_glossary_pairs=set()
        self.nonglossary_pairs=set()
        self.word_report=[]

    # ------------- helpers -------------
    @staticmethod
    def _norm(s): return re.sub(r"\s+"," ",s.strip()).lower() if s else s
    def _sanitize(self,pairs):
        return [(str(p[0]),str(p[1])) for p in pairs if isinstance(p,(list,tuple)) and len(p)==2]
    # -------- glossary --------
    def _load_glossary(self):
        if not(self.glossary_path and os.path.exists(self.glossary_path)): return {}
        ext=os.path.splitext(self.glossary_path)[1].lower()
        df=pd.read_csv(self.glossary_path) if ext==".csv" else pd.read_excel(self.glossary_path)
        wcol=tcol=None
        for c in df.columns:
            cl=c.strip().lower()
            if cl in("word","english"): wcol=c
            if cl in(self.target,"translation","arabic","urdu"): tcol=c
        if not(wcol and tcol): return {}
        return {str(w).strip().lower():str(t).strip()
                for w,t in zip(df[wcol],df[tcol]) if pd.notna(w) and pd.notna(t)}
    # -------- file ops --------
    def _extract(self):
        with zipfile.ZipFile(self.in_file) as z: z.extractall(self.extract)
        return os.path.join(self.extract,"word","document.xml")
    def _repack(self):
        base=self.out_file[:-5]
        shutil.make_archive(base,"zip",self.extract)
        if os.path.exists(self.out_file): os.remove(self.out_file)
        os.rename(base+".zip",self.out_file)
    # ----- numbering tweak -----
    def _arabize_numbering(self):
        num_path = os.path.join(self.extract, "word", "numbering.xml")
        if not (self.target in RTL_LANGUAGES and os.path.exists(num_path)):
            return

        tree = etree.parse(num_path)
        for fmt in tree.xpath('//w:numFmt', namespaces={'w': self.word_ns}):
            val = fmt.get(f'{{{self.word_ns}}}val')
            if val in ("decimal", "decimalZero", "decimalFullWidth", "decimalHalfWidth"):
                fmt.set(f'{{{self.word_ns}}}val', "hindiNumbers")   # ← only change

        tree.write(num_path, encoding="utf-8", xml_declaration=True)
    # -------- text helpers -----
    def _apply_glossary(self,txt,g):
        out=txt
        for gw,gt in sorted(g.items(),key=lambda kv:-len(kv[0])):
            out=re.sub(rf'(?<!\S){re.escape(gw)}(?!\S)',gt,out,flags=re.IGNORECASE)
        return out
    def _tokenize(self,txt,g):
        matches=[]; tl=self._norm(txt)
        for gw in sorted(g,key=len,reverse=True):
            pat=rf'(?<!\w){re.escape(gw)}(?!\w)' if self.target in("ar","ur") else rf'\b{re.escape(gw)}\b'
            for m in re.finditer(pat,tl,re.IGNORECASE): matches.append((m.start(),m.end(),gw))
        matches.sort(); words,last=[],0
        for s,e,gw in matches:
            words+=re.findall(r'\S+',txt[last:s]); words.append(gw); last=e
        words+=re.findall(r'\S+',txt[last:]); return words
    # -------- translate text -------
    def translate_text(self,text):
        if not text.strip(): return text,[],[]
        prefix,core=split_prefix(text)

        gloss=self._load_glossary()
        for gw in gloss:
            if re.search(rf'\b{re.escape(gw)}\b',self._norm(core)):
                self.used_glossary_words.add(gw)
                self.used_glossary_pairs.add((gw,gloss[gw]))

        trans_core=self._apply_glossary(core,gloss)
        tokens=self._tokenize(core,gloss)
        need_llm=[t for t in tokens if t.lower() not in gloss]

        gloss_pairs=nongloss_pairs=[]
        if need_llm:
            prompt=PromptTemplate.from_template(
                """Translate from {source_language} to {output_language}.
                Glossary:
                {glossary_context}
                Text:
                {input}
                Return JSON:
                {{"translated_text":"...","glossary_pairs":[...],"nonglossary_pairs":[...]}}"""
            )
            ctx="\n".join(f"{w}: {t}" for w,t in gloss.items())
            tgt="Arabic" if self.target=="ar" else "Urdu" if self.target=="ur" else self.target
            llm=ChatOpenAI(model_name="gpt-4o",api_key=self.key,temperature=0.3)
            chain=prompt|llm
            for _ in range(3):
                try:
                    raw=chain.invoke({
                        "source_language":"English",
                        "output_language":tgt,
                        "input":core,
                        "glossary_context":ctx
                    }).content.strip().encode("utf-8").decode("utf-8-sig")
                    m=re.search(r"```json\n([\s\S]*?)\n```",raw)
                    data=json.loads(m.group(1) if m else raw)
                    trans_core=data.get("translated_text",trans_core)
                    gloss_pairs=self._sanitize(data.get("glossary_pairs",[]))
                    nongloss_pairs=self._sanitize(data.get("nonglossary_pairs",[]))
                    break
                except Exception: continue

        if self.target in RTL_LANGUAGES:
            trans_core=normalize_rtl_punctuation(trans_core)

        if prefix and self.target in RTL_LANGUAGES:
            pr=to_arabic_prefix(prefix)
            translated=f"{pr}{RLM}\u00A0{trans_core}"
        else:
            translated=f"{prefix}{RLM} {trans_core}" if prefix else trans_core

        for o,t in gloss_pairs:   self.used_glossary_pairs.add((o,t))
        for o,t in nongloss_pairs:self.nonglossary_pairs.add((o,t))

        return translated,gloss_pairs,nongloss_pairs
    # ---------- XML walk ----------
    def _translate_xml(self,xml_path):
        tree=etree.parse(xml_path,etree.XMLParser(remove_blank_text=False))
        root=tree.getroot()
        for p in root.iter(f'{{{self.word_ns}}}p'):
            combined,segs="",[]
            for t in p.iter(f'{{{self.word_ns}}}t'):
                if t.text and t.text.strip():
                    combined+=t.text; segs.append((t,"text"))
                if t.tail and t.tail.strip():
                    combined+=t.tail; segs.append((t,"tail"))
            if combined.strip():
                new,_,_=self.translate_text(combined)
                for el,kind in segs:
                    if kind=="text" and el.text and el.text.strip():
                        el.text,new=new,""
                    elif kind=="tail" and el.tail and el.tail.strip():
                        el.tail,new=new,""
        if self.target in RTL_LANGUAGES:
            for p in root.findall(".//w:p",namespaces={"w":self.word_ns}):
                pPr=p.find("w:pPr",namespaces={"w":self.word_ns})
                if pPr is None:
                    pPr=etree.SubElement(p,f"{{{self.word_ns}}}pPr")
                if pPr.find(f"{{{self.word_ns}}}bidi") is None:
                    etree.SubElement(pPr,f"{{{self.word_ns}}}bidi")
        with open(xml_path,"w",encoding="utf-8") as f:
            f.write(etree.tostring(root,encoding="unicode",pretty_print=True))
    # ---------- driver ----------
    def run(self):
        if os.path.exists(self.extract): shutil.rmtree(self.extract,ignore_errors=True)
        os.makedirs(self.extract,exist_ok=True)
        try:
            xml=self._extract()
            self._translate_xml(xml)
            self._arabize_numbering()   # <-- new numbering tweak
            self._repack()

            gloss=self._load_glossary()
            for w,t in gloss.items():
                if (w,t) not in self.used_glossary_pairs:
                    self.word_report.append({"Original Word":w,"Translated Word":t,
                                             "Source":"Glossary","Used in Document":"No"})
            return (self.word_report,
                    list(self.used_glossary_pairs),
                    list(self.nonglossary_pairs))
        finally:
            shutil.rmtree(self.extract,ignore_errors=True)

# -------- wrapper for Streamlit --------
def translate_file(in_file,out_file,target_language,
                   glossary_path,OPENAI_API_KEY):
    if not in_file.lower().endswith(".docx"):
        raise ValueError("Only .docx files are supported.")
    return DocxTranslator(in_file,out_file,target_language,
                          glossary_path,OPENAI_API_KEY).run()
