import os
import gc
import tempfile
import uuid
import pandas as pd
import streamlit as st

from llama_index.core import Settings, PromptTemplate, VectorStoreIndex
from llama_index.llms.ollama import Ollama
from llama_index.embeddings.huggingface import HuggingFaceEmbedding
from llama_index.core import SimpleDirectoryReader
from llama_index.readers.docling import DoclingReader
from llama_index.core.node_parser import MarkdownNodeParser

# ------------------------------------------------------------
# CONFIGURATION APP
# ------------------------------------------------------------
st.set_page_config(
    page_title="Assistant Excel IA",
    page_icon="📊",
    layout="wide"
)

# ------------------------------------------------------------
# STYLE SIMPLE
# ------------------------------------------------------------
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}

.block-container {
    padding-top: 2rem;
}

.stButton>button {
    background-color: #00c781;
    color: white;
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# SESSION
# ------------------------------------------------------------
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

if "engine" not in st.session_state:
    st.session_state.engine = None

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# ------------------------------------------------------------
# MODELES (CACHE)
# ------------------------------------------------------------
@st.cache_resource
def init_llm():
    return Ollama(model="llama3.2", request_timeout=120.0)

@st.cache_resource
def init_embedding():
    return HuggingFaceEmbedding(
        model_name="BAAI/bge-large-en-v1.5",
        trust_remote_code=True
    )

# ------------------------------------------------------------
# RESET
# ------------------------------------------------------------
def reset():
    st.session_state.chat_history = []
    st.session_state.engine = None
    gc.collect()

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------
col1, col2 = st.columns([6,1])

with col1:
    st.markdown("""
    <h2 style='text-align:center;'>📊 Assistant Excel IA</h2>
    <p style='text-align:center; color:gray;'>
    Uploadez votre fichier → Posez vos questions → Obtenez des insights
    </p>
    """, unsafe_allow_html=True)

with col2:
    st.button("Réinitialiser ↺", on_click=reset)

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
with st.sidebar:

    st.markdown("## 📁 Importer un fichier Excel")

    uploaded_file = st.file_uploader(
        "Glissez votre fichier ici",
        type=["xlsx", "xls"]
    )

    st.markdown("---")

    st.markdown("### 💡 Exemples de questions")
    st.markdown("""
    - Résume les données  
    - Quelles sont les valeurs manquantes ?  
    - Quelles sont les valeurs les plus élevées ?  
    - Donne-moi des insights  
    """)

# ------------------------------------------------------------
# INDEXATION
# ------------------------------------------------------------
def traiter_fichier(file):

    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, file.name)

        with open(path, "wb") as f:
            f.write(file.getvalue())

        reader = DoclingReader()

        loader = SimpleDirectoryReader(
            input_dir=temp_dir,
            file_extractor={".xlsx": reader},
        )

        documents = loader.load_data()

        # Initialisation modèles
        Settings.llm = init_llm()
        Settings.embed_model = init_embedding()

        parser = MarkdownNodeParser()

        index = VectorStoreIndex.from_documents(
            documents,
            transformations=[parser],
            show_progress=True
        )

        query_engine = index.as_query_engine(streaming=True)

        # Prompt optimisé métier
        prompt = PromptTemplate(
            """Tu es un data analyst expert.

Analyse les données Excel et réponds :

- clairement
- avec des chiffres
- avec des insights utiles

Si l'information n'existe pas, dis :
"Je ne trouve pas cette information dans les données."

Contexte :
{context_str}

Question :
{query_str}

Réponse :
"""
        )

        query_engine.update_prompts({
            "response_synthesizer:text_qa_template": prompt
        })

        return query_engine

# ------------------------------------------------------------
# PREVIEW
# ------------------------------------------------------------
def afficher_apercu(file):
    df = pd.read_excel(file)
    st.markdown("### 👀 Aperçu des données")
    st.dataframe(df.head(50))

# ------------------------------------------------------------
# TRAITEMENT FICHIER
# ------------------------------------------------------------
if uploaded_file:

    if st.session_state.engine is None:

        with st.spinner("⚡ Analyse du fichier en cours..."):
            try:
                st.session_state.engine = traiter_fichier(uploaded_file)
                st.success("✅ Fichier prêt ! Posez vos questions.")
            except Exception as e:
                st.error(f"Erreur : {e}")
                st.stop()

    afficher_apercu(uploaded_file)

# ------------------------------------------------------------
# CHAT
# ------------------------------------------------------------
for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

question = st.chat_input("Posez votre question sur les données...")

if question:

    st.session_state.chat_history.append({
        "role": "user",
        "content": question
    })

    with st.chat_message("user"):
        st.markdown(question)

    with st.chat_message("assistant"):

        if st.session_state.engine is None:
            st.error("Veuillez importer un fichier Excel.")
        else:
            zone_reponse = st.empty()
            reponse_complete = ""

            stream = st.session_state.engine.query(question)

            for chunk in stream.response_gen:
                reponse_complete += chunk
                zone_reponse.markdown(reponse_complete + "▌")

            zone_reponse.markdown(reponse_complete)

            st.session_state.chat_history.append({
                "role": "assistant",
                "content": reponse_complete
            })
