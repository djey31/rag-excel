import streamlit as st
import os
import tempfile

st.set_page_config(
    page_title="Excel RAG AI",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------- CSS (inchangé, design premium) ----------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap');

:root {
    --bg: #07090f;
    --surface: #0f1520;
    --surface2: #161f30;
    --border: #1e2d45;
    --accent: #00e5a0;
    --accent2: #3b82f6;
    --text: #e2e8f0;
    --muted: #4e6180;
    --radius: 10px;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    background: var(--bg) !important;
    color: var(--text) !important;
}

#MainMenu, footer, header { visibility: hidden; }

.block-container {
    padding: 1.5rem 2rem 4rem !important;
    max-width: 1200px;
}

[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}

.brand {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 1.2rem;
    display: flex;
    align-items: center;
    gap: 10px;
}

.hex {
    width: 26px;
    height: 26px;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    clip-path: polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%);
}

.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--accent), #00b87a);
    color: #07090f;
}

.stChatInput > div {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------- SESSION ----------------
for k, v in {
    "messages": [],
    "vectorstore": None,
    "qa_chain": None,
    "indexed": False
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------- API KEY ----------------
try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except:
    st.error("GROQ_API_KEY manquante")
    st.stop()

# ---------------- PARAMETRES SIMPLIFIES ----------------
MODEL = "llama-3.3-70b-versatile"
CHUNK_SIZE = 600
CHUNK_OVERLAP = 80
TOP_K = 5
TEMPERATURE = 0.0

# ---------------- EMBEDDINGS ----------------
@st.cache_resource
def get_embeddings():
    from langchain_community.embeddings import HuggingFaceEmbeddings
    return HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2"
    )

# ---------------- INDEXATION ----------------
def index_file(file_bytes, filename):

    from docling.document_converter import DocumentConverter
    from langchain_community.vectorstores import Chroma
    from langchain_text_splitters import RecursiveCharacterTextSplitter

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    converter = DocumentConverter()
    result = converter.convert(tmp_path)
    markdown = result.document.export_to_markdown()

    os.unlink(tmp_path)

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n", "\n", "|", " "]
    )

    chunks = splitter.split_text(markdown)

    embeddings = get_embeddings()

    vectorstore = Chroma.from_texts(
        texts=chunks,
        embedding=embeddings
    )

    return vectorstore

# ---------------- QA ----------------
def build_chain(vectorstore):

    from langchain_groq import ChatGroq
    from langchain.chains import RetrievalQA
    from langchain.prompts import PromptTemplate

    llm = ChatGroq(
        model=MODEL,
        temperature=TEMPERATURE,
        groq_api_key=GROQ_API_KEY
    )

    retriever = vectorstore.as_retriever(search_kwargs={"k": TOP_K})

    prompt = PromptTemplate(
        template="""Tu es un analyste de données.

Réponds clairement avec des informations utiles.
Si l'information n'existe pas, dis que tu ne la trouves pas.

Contexte :
{context}

Question :
{question}

Réponse :
""",
        input_variables=["context", "question"]
    )

    return RetrievalQA.from_chain_type(
        llm=llm,
        retriever=retriever,
        chain_type="stuff",
        chain_type_kwargs={"prompt": prompt}
    )

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown('<div class="brand"><div class="hex"></div>ExcelRAG</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Fichier Excel", type=["xlsx", "xls"])

    st.markdown("""
    <div style="font-size:0.75rem; color:var(--muted);">
    Paramètres optimisés automatiquement
    </div>
    """, unsafe_allow_html=True)

    index_btn = st.button("Analyser", type="primary")

# ---------------- MAIN ----------------
st.markdown("""
<div style="text-align:center; margin-top:40px;">
<h2>Analyse intelligente de fichiers Excel</h2>
<p style="color:#4e6180;">Chargez un fichier puis posez vos questions</p>
</div>
""", unsafe_allow_html=True)

# ---------------- INDEX ----------------
if index_btn:

    if not uploaded:
        st.error("Veuillez charger un fichier")
    else:
        with st.spinner("Analyse en cours"):
            vs = index_file(uploaded.getvalue(), uploaded.name)

            st.session_state.vectorstore = vs
            st.session_state.qa_chain = build_chain(vs)
            st.session_state.indexed = True

        st.success("Fichier prêt")

# ---------------- CHAT ----------------
if st.session_state.indexed:

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    question = st.chat_input("Posez votre question")

    if question:

        st.session_state.messages.append({"role": "user", "content": question})

        with st.chat_message("user"):
            st.markdown(question)

        with st.chat_message("assistant"):

            res = st.session_state.qa_chain.invoke({"query": question})
            answer = res["result"]

            st.markdown(answer)

            st.session_state.messages.append({
                "role": "assistant",
                "content": answer
            })
