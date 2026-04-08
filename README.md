# ⬡ Excel RAG · AI

RAG over Excel en langage naturel — 100% cloud, zéro config pour l'utilisateur.

## Stack
| Composant | Technologie |
|---|---|
| Parser Excel | Docling |
| LLM | Groq (Llama 3, Mixtral, Gemma) |
| Embeddings | sentence-transformers (local) |
| Vector store | ChromaDB (in-memory) |
| Interface | Streamlit |

## Déploiement Streamlit Cloud

1. Fork ce repo
2. Va sur [share.streamlit.io](https://share.streamlit.io)
3. Connecte ton repo GitHub
4. Dans **Settings → Secrets**, ajoute :

```toml
GROQ_API_KEY = "gsk_xxxxxxxxxxxxxxxxxxxx"
```

5. Deploy !

## Obtenir une clé Groq gratuite

→ [console.groq.com](https://console.groq.com) — gratuit, 14 400 req/jour

## Usage local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Ajoute ta clé dans .streamlit/secrets.toml
streamlit run rag_excel_app.py
```
