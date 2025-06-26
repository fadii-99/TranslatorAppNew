import os
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from utils import translate_file
import uuid

# Load environment variables
load_dotenv()
ModernMT_key = os.environ.get("ModernMT_key") or st.secrets.get('ModernMT_key')
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY") or st.secrets.get('OPENAI_API_KEY')

st.set_page_config(page_title="Document Translator", layout="centered")
st.title("Document Translator")

# Upload DOCX file
docx_file = st.file_uploader("Upload your Word document to translate", type=["docx"])

# Upload glossary file
glossary_file = st.file_uploader("Upload your Glossary (CSV or Excel)", type=["csv", "xlsx"])

file_extension = os.path.splitext(docx_file.name)[1].lower()[1:] if docx_file else None

selected_model = st.selectbox("Select Translation Model:", ["OpenAI"])
st.info("Source language: English (Fixed)")

language_map = {"Arabic": "ar", "Urdu": "ur"}
languages = list(language_map.keys())
target_langs = []
if docx_file and file_extension == "docx":
    target_lang = st.selectbox("Select Target Language (for Word Documents):", languages, index=0)
    target_langs = [language_map.get(target_lang)] if target_lang else []
else:
    st.write(f"Please upload a DOCX file to select target languages.")

if docx_file and glossary_file and target_langs:
    if st.button("Translate"):
        st.success(f"Files uploaded successfully! Processing DOCX file...")
        # Save uploaded files temporarily
        input_file_path = f"uploaded_file_{str(uuid.uuid4())}.docx"
        with open(input_file_path, "wb") as f:
            f.write(docx_file.read())

        # Save glossary file
        glossary_ext = os.path.splitext(glossary_file.name)[1].lower()
        glossary_temp_path = f"uploaded_glossary_{str(uuid.uuid4())}{glossary_ext}"
        with open(glossary_temp_path, "wb") as f:
            f.write(glossary_file.read())

        with st.spinner(f"Translating DOCX to {', '.join(target_langs)} using {selected_model}..."):
            output_file_path = f"translated_document_{target_langs[0].lower()}_{selected_model}_{str(uuid.uuid4())}.docx"

            # Translate the file (pass glossary path!)
            used_glossary, used_glossary_pairs, used_non_glossary_pairs =  translate_file(
                input_file_path, output_file_path, target_langs[0], glossary_temp_path, OPENAI_API_KEY
            )

            with open(output_file_path, "rb") as f:
                st.session_state["translated_docx_bytes"] = f.read()
            df_glossary = pd.DataFrame(used_glossary_pairs, columns=['English', f'{target_lang.upper()} Translation'])
            st.session_state["glossary_csv_bytes"] = df_glossary.to_csv(index=False).encode('utf-8')
            df_non_glossary = pd.DataFrame(used_non_glossary_pairs, columns=['English', f'{target_lang.upper()} Translation'])
            st.session_state["non_glossary_csv_bytes"] = df_non_glossary.to_csv(index=False).encode('utf-8')

            st.session_state["target_lang"] = target_lang
            st.session_state["target_lang_code"] = target_langs[0]
            st.session_state["file_extension"] = file_extension
            st.success("Translation complete! You can now download your files below.")

if "translated_docx_bytes" in st.session_state:
    target_lang = st.session_state.get("target_lang", "Arabic")
    target_lang_code = st.session_state.get("target_lang_code", "ar")
    file_extension = st.session_state.get("file_extension", "docx")
    mime_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "odt": "application/vnd.oasis.opendocument.text",
        "doc": "application/doc"
    }

    st.markdown("---")
    st.subheader("Download your translated files:")

    st.download_button(
        label=f"Download Translated Document ({target_lang})",
        data=st.session_state["translated_docx_bytes"],
        file_name=f"translated_{target_lang_code}.{file_extension}",
        mime=mime_types[file_extension],
        use_container_width=True
    )

    st.write("### Glossary Words Used in Translation:")
    df_glossary = pd.read_csv(pd.io.common.BytesIO(st.session_state["glossary_csv_bytes"]))
    st.dataframe(df_glossary, use_container_width=True)
    st.download_button(
        "Download Used Glossary List (CSV)",
        st.session_state["glossary_csv_bytes"],
        f"used_glossary_{target_lang_code}.csv",
        "text/csv",
        use_container_width=True
    )

    st.write("### Words Translated by LLM (Not in Glossary):")
    df_non_glossary = pd.read_csv(pd.io.common.BytesIO(st.session_state["non_glossary_csv_bytes"]))
    st.dataframe(df_non_glossary, use_container_width=True)
    st.download_button(
        "Download Non-Glossary Translated List (CSV)",
        st.session_state["non_glossary_csv_bytes"],
        f"non_glossary_translations_{target_lang_code}.csv",
        "text/csv",
        use_container_width=True
    )

else:
    if docx_file and not target_langs:
        st.warning("Please select a target language and click 'Translate' to proceed.")
    elif docx_file:
        st.info("Select a target language and click 'Translate' to start.")
