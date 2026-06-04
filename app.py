import os
import io

import streamlit as st
from docx import Document
from PIL import Image

st.set_page_config(page_title="Tesina", layout="centered")

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "DEFINITIVA")

# --- ELENCO DEI CAPITOLI ---
# Ogni voce: (nome_file, numero_romano, titolo_completo)
#   nome_file      → es. "terzo" → "capitolo terzo testo.docx" / "CAPITOLO TERZO.docx"
#   numero_romano  → etichetta nella sidebar  (es. "III")
#   titolo_completo → mostrato nella sidebar e come sottotitolo nella pagina
#
# Per AGGIUNGERE un capitolo: aggiungi una riga + crea i due file .docx.
# Per MODIFICARE un titolo: cambia solo la stringa titolo_completo.
CHAPTERS = [
    ("primo",      "I",    'Italo Calvino e il segreto del ritmo ne "Le mille e una notte"'),
    ("secondo",    "II",   "Le pont entre deux mondes: Antoine Galland et la magie de l'Orient"),
    ("terzo",      "III",  'Le radici geografiche de "Le mille e una notte"'),
    ("quarto",     "IV",   'La spiritualità islamica ne "Le mille e una notte"'),
    ("quinto",     "V",    "Il Medio Oriente: un'area senza pace"),
    ("sesto",      "VI",   "L'incantesimo dell'oro: dai tesori di Shehrazàd all'oro nero"),
    ("settimo",    "VII",  "Aladdin... dietro le quinte del pregiudizio"),
    ("ottavo",     "VIII", "Aladdin contro Newton: il conflitto tra incantesimo e gravità"),
    ("nono",       "IX",   "Shehrazàd, la sostanza dei sogni di Magritte"),
    ("decimo",     "X",    "Shehrazàd: la voce del violino"),
    ("undicesimo", "XI",   "Jasmine, l'icona della danza orientale"),
    ("dodicesimo", "XII",  "The courage to tell: Shehrazàd and Malala"),
]


# --- LETTURA TESTO DA .docx ---
# skip_chapter_label=True: salta "CAPITOLO [nome]" e deduplica heading in maiuscolo
# skip_label: salta qualsiasi paragrafo che corrisponde esattamente a questa stringa
@st.cache_data
def read_text(filename: str, skip_chapter_label: bool = False, skip_label: str = "") -> str:
    path = os.path.join(BASE, filename)
    if not os.path.exists(path):
        return f"_File non trovato: {filename}_"
    try:
        doc = Document(path)
        lines = []
        seen_caps = []
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if skip_label and text.upper() == skip_label.upper():
                continue
            if skip_chapter_label:
                words = text.split()
                if words[0].upper() == "CAPITOLO" and len(words) <= 3:
                    continue
                if len(lines) < 10 and text.upper() == text:
                    norm = text.upper()
                    if any(norm in seen for seen in seen_caps):
                        continue
                    seen_caps.append(norm)
            try:
                style = p.style.name
            except Exception:
                style = ""
            if style.startswith("Heading 1"):
                lines.append(f"# {text}")
            elif style.startswith("Heading 2"):
                lines.append(f"## {text}")
            elif style.startswith("Heading"):
                lines.append(f"### {text}")
            else:
                lines.append(text)
        if not lines:
            return "_Il file non contiene testo leggibile._"
        return "\n\n".join(lines)
    except Exception as e:
        return f"_Errore nella lettura di `{filename}`: {e}_"


# --- ESTRAZIONE IMMAGINI DA .docx ---
@st.cache_data
def extract_images(filename: str) -> list:
    path = os.path.join(BASE, filename)
    if not os.path.exists(path):
        return []
    try:
        doc = Document(path)
        images = []
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                try:
                    images.append(Image.open(io.BytesIO(rel.target_part.blob)))
                except Exception:
                    pass
        return images
    except Exception:
        return []


# --- STRUTTURA DI NAVIGAZIONE ---
# "label"  → testo mostrato nella sidebar
# "type"   → usato nel routing sottostante
# "name"   → (capitoli) nome file
# "roman"  → (capitoli) numero romano
nav = [
    {"label": "Copertina",    "type": "copertina"},
    {"label": "Dedica",       "type": "dedica"},
    {"label": "Introduzione", "type": "introduzione"},
]
for name, roman, title in CHAPTERS:
    nav.append({
        "label": f"Capitolo {roman} — {title}",
        "type":  "chapter",
        "name":  name,
        "roman": roman,
    })
nav += [
    {"label": "Conclusioni",    "type": "conclusioni"},
    {"label": "Ringraziamenti", "type": "ringraziamenti"},
]

with st.sidebar:
    st.markdown("## Indice")
    idx = st.radio(
        "",
        range(len(nav)),
        format_func=lambda i: nav[i]["label"],
        label_visibility="collapsed",
    )

section = nav[idx]


# =============================================================================
# RENDERING DELLE PAGINE
# =============================================================================

if section["type"] == "copertina":
    cover_path = os.path.join(BASE, "le_mille_e_una_notte.png")
    if os.path.exists(cover_path):
        st.image(cover_path, use_container_width=True)
    text = read_text("copertina con simbolo.docx")
    if text:
        st.markdown(text)

elif section["type"] == "dedica":
    st.title("Dedica")
    st.markdown(read_text("dedica.docx"))

elif section["type"] == "introduzione":
    st.title("Introduzione")
    st.markdown(read_text("Introduzione testo.docx", skip_label="INTRODUZIONE"))
    images = extract_images("INTRODUZIONE.docx")
    if images:
        st.divider()
        for img in images:
            st.image(img, use_container_width=True)

elif section["type"] == "chapter":
    name  = section["name"]
    roman = section["roman"]
    st.title(f"Capitolo {roman}")
    st.markdown(read_text(f"capitolo {name} testo.docx", skip_chapter_label=True))
    images = extract_images(f"CAPITOLO {name.upper()}.docx")
    if images:
        st.divider()
        for img in images:
            st.image(img, use_container_width=True)

elif section["type"] == "conclusioni":
    st.title("Conclusioni")
    st.markdown(read_text("CONCLUSIONI.docx", skip_label="CONCLUSIONI"))
    images = extract_images("CONCLUSIONI FOTO.docx")
    if images:
        st.divider()
        for img in images:
            st.image(img, use_container_width=True)

elif section["type"] == "ringraziamenti":
    st.title("Ringraziamenti")
    st.markdown(read_text("ringraziamenti finali.docx"))