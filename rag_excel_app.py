import streamlit as st
import os
import tempfile

# ------------------------------------------------------------
# CONFIG APP
# ------------------------------------------------------------
st.set_page_config(
    page_title="Excel AI",
    page_icon="📊",
    layout="wide"
)

# ------------------------------------------------------------
# SESSION STATE
# ------------------------------------------------------------
for k, v in {
    "messages": [],
    "vectorstore": None,
    "qa_chain": None,
    "indexed": False,
    "file_info": {}
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ------------------------------------------------------------
# GROQ API KEY
# ------------------------------------------------------------
try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except:
    st.error("❌ Ajoute ta clé GROQ dans secrets")
    st.stop()

# ------------------------------------------------------------
# PARAMÈTRES CACHÉS (IMPORTANT 🔥)
# ------------------------------------------------------------
MODEL = "llama-3.3-70b-versatile"
CHUNK_SIZE = 600
CHUNK_OVERLAP = 80
TOP_K = 5
TEMPERATURE = 0.0

# ------------------------------------------------------------
# EMBEDDINGS
# ------------------------------------------------------------
@st.cache_resource
def get_embeddings():
    from langchain_community.embeddings import HuggingFaceEmbeddings
    return HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2"
    )

# ------------------------------------------------------------
# INDEXATION (DOC + CHUNKS + VECTOR DB)
# ------------------------------------------------------------
def index_file(file_bytes, filename):

    from docling.document_converter import DocumentConverter
    from langchain_community.vectorstores import Chroma
    from langchain.text_splitter import RecursiveCharacterTextSplitter

    # 1. Sauvegarde temporaire
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    # 2. DOC LING (🔥 clé)
    converter = DocumentConverter()
    result = converter.convert(tmp_path)

    # Excel → texte markdown
    markdown = result.document.export_to_markdown()

    os.unlink(tmp_path)

    # 3. SPLIT TEXTE
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n", "\n", "|", " "]
    )

    chunks = splitter.split_text(markdown)

    # 4. EMBEDDINGS
    embeddings = get_embeddings()

    # 5. VECTOR STORE
    vectorstore = Chroma.from_texts(
        texts=chunks,
        embedding=embeddings,
        metadatas=[{"source": filename, "chunk": i} for i in range(len(chunks))]
    )

    return vectorstore, len(chunks), markdown

# ------------------------------------------------------------
# QA CHAIN (LLM)
# ------------------------------------------------------------
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
        template="""Tu es un data analyst expert.

Analyse les données Excel et réponds :

- clairement
- avec des chiffres
- avec des insights utiles

Si tu ne trouves pas, dis :
"Je ne trouve pas cette information dans les données."

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
        chain_type_kwargs={"prompt": prompt},
        return_source_documents=True,
    )

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
with st.sidebar:

    st.markdown("## 📁 Upload Excel")

    uploaded = st.file_uploader("Fichier", type=["xlsx", "xls"])

    st.markdown("⚙️ Paramètres optimisés automatiquement")

    index_btn = st.button("Analyser")

# ------------------------------------------------------------
# MAIN UI
# ------------------------------------------------------------
st.title("📊 Excel AI Assistant")

if index_btn:

    if not uploaded:
        st.error("Upload un fichier")
    else:
        with st.spinner("Analyse en cours..."):

            vs, n, md = index_file(uploaded.getvalue(), uploaded.name)

            st.session_state.vectorstore = vs
            st.session_state.qa_chain = build_chain(vs)
            st.session_state.indexed = True

            st.success(f"✅ {n} morceaux analysés")

# ------------------------------------------------------------
# CHAT
# ------------------------------------------------------------
if st.session_state.indexed:

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    question = st.chat_input("Pose ta question...")

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
