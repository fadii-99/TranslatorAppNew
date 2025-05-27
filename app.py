import os
import streamlit as st
import pandas as pd
import io
from langchain_community.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from dotenv import load_dotenv
from utils import translate_file
import uuid
# Load environment variables
load_dotenv()
# Try to get keys from environment variables first, then fall back to Streamlit secrets
ModernMT_key = os.environ.get("ModernMT_key") or st.secrets.get('ModernMT_key')
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY") or st.secrets.get('OPENAI_API_KEY')

# print(f"OPENAI_API_KEY: {OPENAI_API_KEY}")
# print(f"ModernMT_key: {ModernMT_key}")


# Streamlit UI
st.set_page_config(page_title="Document Translator", layout="centered")
st.title("Document Translator")

# Step 1: User selects file format
# file_format = st.selectbox("Choose File Format to Upload:", [ "DOCX", "TXT", "ODT"])
file_format = st.selectbox("Choose File Format to Upload:", [ "DOCX"])

# Step 2: File upload based on selected format
if file_format == "DOCX":
    uploaded_file = st.file_uploader("Upload your Word document", type=["docx"])
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()[1:] if uploaded_file else None


# Model selection
selected_model = st.selectbox("Select Translation Model:", [ "OpenAI"])

# Fixed source language
st.info("Source language: English (Fixed)")

# Target language selection
language_map = {
    "Arabic": "ar",
    "Urdu": "ur",
}
languages = list(language_map.keys())  # Get the list of languages from the dictionary
target_langs = []

if uploaded_file:
    if file_extension in "docx":
        target_lang = st.selectbox("Select Target Language (for Word Documents):", languages, index=0)  # Default to Arabic
        target_langs = [language_map.get(target_lang)] if target_lang else []
else:
    st.write(f"Please upload a {file_format} file to select target languages.")

# Add a Translate button to initiate translation
if uploaded_file and target_langs:
    if st.button("Translate"):
        st.success(f"File uploaded successfully! Processing {file_format} file...")
        
        # Save uploaded file temporarily
        input_file_path = f"uploaded_file_{str(uuid.uuid4())}.{file_extension}"
        with open(input_file_path, "wb") as f:
            f.write(uploaded_file.read())

        # Start translation process with loading spinner
        with st.spinner(f"Translating {file_format} to {', '.join(target_langs)} using {selected_model}..."):

            output_file_path = f"translated_document_{target_langs[0].lower()}_{selected_model}_{str(uuid.uuid4())}.{file_extension}"
            
            # Translate the file and get the glossary usage info
            if selected_model == 'OpenAI':
                used_glossary, used_glossary_pairs = translate_file(
                    input_file_path, output_file_path, target_langs[0], None, OPENAI_API_KEY
                )
            else:
                used_glossary, used_glossary_pairs = translate_file(
                    input_file_path, output_file_path, target_langs[0], ModernMT_key, None
                )

            # Define appropriate MIME types for DOCX files
            mime_types = {
                "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "odt": "application/vnd.oasis.opendocument.text",
                "doc": "application/doc"
            }

            with open(output_file_path, "rb") as f:
                st.download_button(
                    label=f"Download Translated Document ({target_langs[0]})",
                    data=f,
                    file_name=f"{output_file_path}",
                    mime=mime_types[file_extension],
                )
            
            # Show used glossary words in the Streamlit app
            st.write("### Glossary Words Used in Translation:")
            if used_glossary:
                st.markdown("---")
                st.write("📚 **Matched Glossary Terms:**")
                
                # Create columns for better space utilization
                TERMS_PER_COLUMN = 10
                terms = sorted(used_glossary)
                
                # Display terms in a grid layout
                cols = st.columns(3)  # Create 3 columns
                for i, term in enumerate(terms):
                    col_index = i % 3
                    cols[col_index].write(f"• {term}")
                
                # Add pagination if there are many terms
                # if len(terms) > 30:  # Show pagination for large lists
                #     items_per_page = 30
                #     pages = len(terms) // items_per_page + (1 if len(terms) % items_per_page > 0 else 0)
                #     current_page = st.selectbox("Page", range(1, pages + 1), label_visibility="collapsed")
                #     start_idx = (current_page - 1) * items_per_page
                #     end_idx = min(start_idx + items_per_page, len(terms))
                    
                st.markdown("---")
                import pandas as pd
                df = pd.DataFrame(used_glossary_pairs, columns=['English', f'{target_lang.upper()} Translation'])
                csv = df.to_csv(index=False)
                st.download_button(
                    "Download Used Glossary List (CSV)",
                    csv,
                    f"used_glossary_{target_langs[0]}.csv",
                    "text/csv"
                )
            else:
                st.write("No glossary words were matched in this translation.")



else:
    if uploaded_file and not target_langs:
        st.warning("Please select a target language and click 'Translate' to proceed.")
    elif uploaded_file:
        st.info("Select a target language and click 'Translate' to start.")



















        