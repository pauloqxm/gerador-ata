import io
import datetime as dt
from typing import List, Dict


import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


# ==========================
# Estilos e utilidades DOCX
# ==========================


def _set_default_styles(doc: Document):
styles = doc.styles
# Fonte padrão
for style_name in ["Normal", "Heading 1", "Heading 2", "Heading 3", "Title"]:
if style_name in styles:
style = styles[style_name]
font = style.font
font.name = "Calibri"
font.size = Pt(11)
# Para compatibilidade com MS Word
try:
style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
except Exception:
pass


if "Title" in styles:
styles["Title"].font.size = Pt(20)
styles["Title"].font.bold = True


if "Heading 1" in styles:
styles["Heading 1"].font.size = Pt(14)
styles["Heading 1"].font.bold = True


if "Heading 2" in styles:
styles["Heading 2"].font.size = Pt(12)
styles["Heading 2"].font.bold = True




def _add_header_block(doc: Document, org_name: str, logo_bytes: bytes | None):
if logo_bytes:
# Tentar inserir logo no cabeçalho superior
section = doc.sections[0]
header = section.header
header_para = header.paragraphs[0]
run = header_para.add_run()
try:
run.add_picture(io.BytesIO(logo_bytes), width=Inches(1.2))
except Exception:
pass
st.info("Preencha o formulário acima e clique em 'Gerar .DOCX da Ata'.")
