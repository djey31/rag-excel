import streamlit as st
import os
import tempfile

st.set_page_config(
    page_title="Excel RAG · AI",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap');

:root {
    --bg:       #07090f;
    --surface:  #0f1520;
    --surface2: #161f30;
    --border:   #1e2d45;
    --accent:   #00e5a0;
    --accent2:  #3b82f6;
    --text:     #e2e8f0;
    --muted:    #4e6180;
    --radius:   10px;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    background: var(--bg) !important;
    color: var(--text) !important;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 1.5rem 2rem 4rem !important; max-width: 1300px; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 1.2rem; }

.brand {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 1.25rem;
    letter-spacing: -0.3px;
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 1.8rem;
    padding: 0 0.4rem;
}
.hex {
    width: 28px; height: 28px; flex-shrink: 0;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    clip-path: polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%);
}

.slabel {
    font-family: 'DM Mono', monospace;
    font-size: 0.62rem; font-weight: 500;
    letter-spacing: 0.14em; text-transform: uppercase;
    color: var(--muted); margin: 1.4rem 0 0.45rem;
    padding: 0 0.3rem;
}

.badge {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 3px 9px; border-radius: 20px;
    font-family: 'DM Mono', monospace;
    font-size: 0.68rem;
}
.badge-ok  { background: rgba(0,229,160,.1); color: var(--accent); border: 1px solid rgba(0,229,160,.2); }
.badge-off { background: rgba(78,97,128,.1); color: var(--muted);  border: 1px solid rgba(78,97,128,.2); }
.badge-run { background: rgba(59,130,246,.1); color: var(--accent2); border: 1px solid rgba(59,130,246,.2); }
.dot { width: 5px; height: 5px; border-radius: 50%; background: currentColor; }
.pulse { animation: blink 1.4s ease-in-out infinite; }
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:.25} }

/* Uploader */
[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1.5px dashed var(--border) !important;
    border-radius: var(--radius) !important;
}
[data-testid="stFileUploader"]:hover { border-color: var(--accent) !important; }

/* Buttons */
.stButton > button {
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; font-size: 0.82rem !important;
    border-radius: 8px !important; border: none !important;
    padding: 0.5rem 1rem !important;
    transition: all .18s !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--accent), #00b87a) !important;
    color: #07090f !important;
}
.stButton > button[kind="primary"]:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(0,229,160,.25) !important; }
.stButton > button[kind="secondary"] {
    background: var(--surface2) !important; color: var(--text) !important;
    border: 1px solid var(--border) !important;
}

/* Hero */
.hero {
    text-align: center;
    padding: 3.5rem 1rem 1.5rem;
    max-width: 560px; margin: 0 auto;
}
.hero-hex {
    width: 60px; height: 60px; margin: 0 auto 1.3rem;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    clip-path: polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%);
}
.hero h1 {
    font-family: 'Syne', sans-serif; font-weight: 800;
    font-size: 2rem; letter-spacing: -1px; margin-bottom: .6rem;
    background: linear-gradient(135deg, #e2e8f0, #4e6180);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.hero p { color: var(--muted); font-size: .92rem; line-height: 1.65; }

.steps {
    max-width: 460px; margin: 2rem auto 0;
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 1.3rem 1.6rem;
}
.step {
    display: flex; align-items: baseline; gap: 12px;
    font-size: .87rem; line-height: 2;
}
.step-n {
    font-family: 'DM Mono', monospace; font-size: .65rem;
    color: var(--accent); flex-shrink: 0;
}

/* Stats */
.stats { display: flex; gap: .8rem; margin-bottom: 1.2rem; }
.stat {
    flex: 1; background: var(--surface);
    border: 1px solid var(--border); border-radius: var(--radius);
    padding: .9rem 1.1rem;
}
.stat-v {
    font-family: 'Syne', sans-serif; font-weight: 700;
    font-size: 1.35rem; color: var(--accent);
}
.stat-l {
    font-family: 'DM Mono', monospace; font-size: .62rem;
    text-transform: uppercase; letter-spacing: .1em;
    color: var(--muted); margin-top: 2px;
}

/* Chat */
.chat-wrap { display: flex; flex-direction: column; gap: .9rem; margin-bottom: 1.2rem; }

.msg-u {
    align-self: flex-end; max-width: 68%;
    background: linear-gradient(135deg,rgba(59,130,246,.18),rgba(59,130,246,.08));
    border: 1px solid rgba(59,130,246,.25);
    border-radius: 14px 14px 3px 14px;
    padding: .75rem 1rem; font-size: .88rem; line-height: 1.6;
}
.msg-a {
    align-self: flex-start; max-width: 78%;
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: 3px 14px 14px 14px;
    padding: .75rem 1rem; font-size: .88rem; line-height: 1.7;
}
.msg-lbl {
    font-family: 'DM Mono', monospace; font-size: .6rem;
    text-transform: uppercase; letter-spacing: .1em; margin-bottom: .35rem;
}
.msg-u .msg-lbl { color: var(--accent2); }
.msg-a .msg-lbl { color: var(--accent); }
.msg-src {
    margin-top: .6rem; padding-top: .5rem;
    border-top: 1px solid var(--border);
    font-family: 'DM Mono', monospace;
    font-size: .68rem; color: var(--muted);
}

/* Suggestions */
.sug-label {
    font-family: 'DM Mono', monospace; font-size: .62rem;
    text-transform: uppercase; letter-spacing: .12em;
    color: var(--muted); margin-bottom: .6rem;
}

/* Selectbox */
[data-testid="stSelectbox"] > div > div {
    background: var(--surface2) !important;
    border-color: var(--border) !important;
    color: var(--text) !important; border-radius: 8px !important;
}

/* Slider */
[data-testid="stSlider"] { color: var(--text) !important; }

/* Chat input */
.stChatInput > div {
    background: var(--surface) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
}
.stChatInput > div:focus-within { border-color: var(--accent) !important; }

/* Expander */
[data-testid="stExpander"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}

/* Scrollbar */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 99px; }

hr { border-color: var(--border) !important; margin: .8rem 0 !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──────────────────────────────────────────────────────────────
for k, v in {
    "messages": [], "vectorstore": None, "qa_chain": None,
    "indexed": False, "file_info": {}
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Load Groq API key from secrets ────────────────────────────────────────────
try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    GROQ_API_KEY = None

# ── Helpers ────────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_embeddings():
    from langchain_community.embeddings import HuggingFaceEmbeddings
    return HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

def index_file(file_bytes, filename, chunk_size, chunk_overlap):
    from docling.document_converter import DocumentConverter
    from langchain_community.vectorstores import Chroma
    from langchain.text_splitter import RecursiveCharacterTextSplitter

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    converter = DocumentConverter()
    result = converter.convert(tmp_path)
    markdown = result.document.export_to_markdown()
    os.unlink(tmp_path)

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size, chunk_overlap=chunk_overlap,
        separators=["\n\n", "\n", "|", " "]
    )
    chunks = splitter.split_text(markdown)

    embeddings = get_embeddings()
    vectorstore = Chroma.from_texts(
        texts=chunks, embedding=embeddings,
        metadatas=[{"source": filename, "chunk": i} for i in range(len(chunks))]
    )
    return vectorstore, len(chunks), markdown

def build_chain(vectorstore, model_name, k, temperature):
    from langchain_groq import ChatGroq
    from langchain.chains import RetrievalQA
    from langchain.prompts import PromptTemplate

    llm = ChatGroq(
        model=model_name,
        temperature=temperature,
        groq_api_key=GROQ_API_KEY
    )
    retriever = vectorstore.as_retriever(search_kwargs={"k": k})
    prompt = PromptTemplate(
        template="""You are an expert data analyst. Analyze the Excel data in the context and answer precisely.
If the answer is not in the context, say "I don't find this information in the provided data."
Be concise, structured, and factual. Format numbers clearly when relevant.

Context (Excel data):
{context}

Question: {question}

Answer:""",
        input_variables=["context", "question"]
    )
    return RetrievalQA.from_chain_type(
        llm=llm, retriever=retriever, chain_type="stuff",
        chain_type_kwargs={"prompt": prompt},
        return_source_documents=True,
    )

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="brand"><div class="hex"></div>ExcelRAG</div>', unsafe_allow_html=True)

    if not GROQ_API_KEY:
        st.error("⚠️ GROQ_API_KEY manquante dans les secrets Streamlit.")
        st.stop()

    # Status
    if st.session_state.indexed:
        st.markdown('<span class="badge badge-ok"><span class="dot"></span>Index prêt</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="badge badge-off"><span class="dot"></span>Aucun index</span>', unsafe_allow_html=True)

    st.markdown('<div class="slabel">📁 Fichier Excel</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload", type=["xlsx", "xls"], label_visibility="collapsed")

    st.markdown('<div class="slabel">🤖 Modèle Groq</div>', unsafe_allow_html=True)
    model = st.selectbox("Modèle", [
        "llama-3.1-8b-instant",
        "llama-3.3-70b-versatile",
        "mixtral-8x7b-32768",
        "gemma2-9b-it",
    ], label_visibility="collapsed")

    st.markdown('<div class="slabel">⚙️ Paramètres</div>', unsafe_allow_html=True)
    chunk_size    = st.slider("Chunk size",   200, 1500, 600, 50)
    chunk_overlap = st.slider("Chunk overlap",  0,  300,  80, 10)
    top_k         = st.slider("Top-K chunks",   1,   10,   5,  1)
    temperature   = st.slider("Temperature",  0.0,  1.0, 0.0, 0.05)

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        index_btn = st.button("⬡ Indexer", type="primary", use_container_width=True)
    with c2:
        clear_btn = st.button("Reset", type="secondary", use_container_width=True)

    if clear_btn:
        for k in ["messages","vectorstore","qa_chain","indexed","file_info"]:
            st.session_state[k] = [] if k == "messages" else (False if k == "indexed" else (None if k in ["vectorstore","qa_chain"] else {}))
        st.rerun()

    if index_btn:
        if not uploaded:
            st.error("Upload un fichier Excel d'abord.")
        else:
            with st.spinner("Parsing & indexation…"):
                try:
                    vs, n, md = index_file(uploaded.getvalue(), uploaded.name, chunk_size, chunk_overlap)
                    st.session_state.vectorstore = vs
                    st.session_state.qa_chain = build_chain(vs, model, top_k, temperature)
                    st.session_state.indexed = True
                    st.session_state.file_info = {
                        "name": uploaded.name,
                        "size": f"{uploaded.size/1024:.1f} KB",
                        "chunks": n, "preview": md[:800]
                    }
                    st.success(f"✓ {n} chunks indexés")
                except Exception as e:
                    st.error(f"Erreur : {e}")

    if st.session_state.file_info:
        fi = st.session_state.file_info
        st.markdown("---")
        st.markdown(f"""<div style="font-family:'DM Mono',monospace;font-size:.7rem;
            color:var(--muted,#4e6180);line-height:1.9;">
            📄 {fi['name']}<br>📦 {fi['size']}<br>🔢 {fi['chunks']} chunks
        </div>""", unsafe_allow_html=True)
        with st.expander("Aperçu du contenu"):
            st.code(fi["preview"], language="markdown")

# ── Main ───────────────────────────────────────────────────────────────────────
if not st.session_state.indexed:
    st.markdown("""
    <div class="hero">
        <div class="hero-hex"></div>
        <h1>Excel RAG · AI</h1>
        <p>Uploadez votre fichier Excel, indexez-le,<br>puis posez toutes vos questions en langage naturel.</p>
    </div>
    <div class="steps">
        <div style="font-family:'DM Mono',monospace;font-size:.62rem;text-transform:uppercase;
            letter-spacing:.12em;color:var(--muted,#4e6180);margin-bottom:.8rem;">Démarrage rapide</div>
        <div class="step"><span class="step-n">01</span> Uploadez votre <b>.xlsx</b> dans la sidebar</div>
        <div class="step"><span class="step-n">02</span> Choisissez un modèle Groq</div>
        <div class="step"><span class="step-n">03</span> Cliquez <b>⬡ Indexer</b></div>
        <div class="step"><span class="step-n">04</span> Posez vos questions</div>
    </div>
    """, unsafe_allow_html=True)

else:
    fi = st.session_state.file_info
    st.markdown(f"""<div class="stats">
        <div class="stat"><div class="stat-v" style="font-size:.95rem;color:var(--text)">{fi['name']}</div><div class="stat-l">Fichier actif</div></div>
        <div class="stat"><div class="stat-v">{fi['chunks']}</div><div class="stat-l">Chunks indexés</div></div>
        <div class="stat"><div class="stat-v">{len(st.session_state.messages)//2}</div><div class="stat-l">Questions posées</div></div>
        <div class="stat"><div class="stat-v" style="font-size:.85rem;color:var(--text)">{model.split("-")[0].upper()}</div><div class="stat-l">Modèle actif</div></div>
    </div>""", unsafe_allow_html=True)

    # Suggestions
    if not st.session_state.messages:
        st.markdown('<div class="sug-label">Suggestions</div>', unsafe_allow_html=True)
        sugs = [
            "Quelles sont les colonnes et leurs types de données ?",
            "Résume les statistiques clés de ce fichier.",
            "Quelles sont les 5 valeurs les plus élevées ?",
            "Y a-t-il des valeurs manquantes ou aberrantes ?",
        ]
        cols = st.columns(2)
        for i, s in enumerate(sugs):
            with cols[i % 2]:
                if st.button(s, key=f"s{i}", use_container_width=True):
                    st.session_state.messages.append({"role": "user", "content": s})
                    with st.spinner(""):
                        res = st.session_state.qa_chain.invoke({"query": s})
                    srcs = list({d.metadata.get("chunk","") for d in res["source_documents"]})
                    st.session_state.messages.append({"role":"assistant","content":res["result"],"sources":srcs})
                    st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)

    # Chat history
    if st.session_state.messages:
        st.markdown('<div class="chat-wrap">', unsafe_allow_html=True)
        for msg in st.session_state.messages:
            if msg["role"] == "user":
                st.markdown(f'<div class="msg-u"><div class="msg-lbl">Vous</div>{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                src_html = ""
                if msg.get("sources"):
                    src_html = f'<div class="msg-src">Sources · chunks {", ".join(str(s) for s in msg["sources"])}</div>'
                st.markdown(f'<div class="msg-a"><div class="msg-lbl">⬡ ExcelRAG</div>{msg["content"]}{src_html}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Input
    question = st.chat_input("Posez votre question sur le fichier Excel…")
    if question:
        st.session_state.messages.append({"role":"user","content":question})
        with st.spinner("Analyse en cours…"):
            try:
                res = st.session_state.qa_chain.invoke({"query": question})
                srcs = list({d.metadata.get("chunk","") for d in res["source_documents"]})
                st.session_state.messages.append({"role":"assistant","content":res["result"],"sources":srcs})
            except Exception as e:
                st.session_state.messages.append({"role":"assistant","content":f"⚠️ Erreur : {e}","sources":[]})
        st.rerun()
