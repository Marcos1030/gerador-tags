import os
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import streamlit as st

def set_cell_margins(cell, top=20, bottom=20, left=20, right=20):
    """Define margens internas reduzidas na célula para otimizar espaço."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def gerar_word_personalizado(image_path, largura_tag, altura_tag, qtd_tags):
    doc = Document()
    
    # Margens da página otimizadas (0.4 cm)
    for section in doc.sections:
        section.top_margin = Cm(0.4)
        section.bottom_margin = Cm(0.4)
        section.left_margin = Cm(0.4)
        section.right_margin = Cm(0.4)

    # Calcular o número de colunas que cabem numa página A4 largura útil (~20 cm)
    largura_pagina_util = 20.0  # cm aproximados
    cols = max(1, int(largura_pagina_util // largura_tag))
    
    # Calcular quantas linhas são necessárias para a quantidade de tags escolhida
    import math
    rows = math.ceil(qtd_tags / cols)

    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tag_contador = 0
    for i in range(rows):
        row = table.rows[i]
        row.height = Cm(altura_tag)
        row.height_rule = 1
        for j in range(cols):
            if tag_contador >= qtd_tags:
                break
            cell = row.cells[j]
            cell.width = Cm(largura_tag)
            set_cell_margins(cell)
            
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run()
            # Deixar a imagem ligeiramente menor que a célula para dar margem de corte
            run.add_picture(image_path, width=Cm(largura_tag * 0.95))
            tag_contador += 1

    output_path = "tags_personalizadas.docx"
    doc.save(output_path)
    return output_path

# Interface Streamlit
st.title("🧁 Gerador Personalizado de Tags")
st.write("Personalize o tamanho e a quantidade de tags à sua medida para impressão.")

# Opções de Personalização na barra lateral ou no ecrã principal
st.sidebar.header("⚙️ Definições de Impressão")
largura_tag = st.sidebar.slider("Largura da Tag (cm)", min_value=2.0, max_value=10.0, value=4.0, step=0.5)
altura_tag = st.sidebar.slider("Altura da Tag (cm)", min_value=2.0, max_value=10.0, value=4.0, step=0.5)
qtd_tags = st.sidebar.number_input("Quantidade Total de Tags", min_value=1, max_value=100, value=35, step=1)

uploaded_file = st.file_uploader("Escolha a imagem da tag (PNG ou JPG)", type=["png", "jpg", "jpeg"], key="tag_uploader_custom")

if uploaded_file is not None:
    temp_image_path = "temp_tag.png"
    with open(temp_image_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.success("Imagem carregada com sucesso!")
    st.info(configuracao_txt := f"Configuração selecionada: **{qtd_tags} tags** com tamanho **{largura_tag}x{altura_tag} cm**.")
    
    if st.button("Gerar Documento Word Personalizado"):
        with st.spinner("A criar documento ajustado..."):
            docx_file = gerar_word_personalizado(temp_image_path, largura_tag, altura_tag, int(qtd_tags))
            
            with open(docx_file, "rb") as f:
                st.download_button(
                    label="📥 Descarregar Ficheiro Word (.docx)",
                    data=f,
                    file_name=f"Tags_{int(largura_tag)}x{int(altura_tag)}cm_{qtd_tags}un.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
