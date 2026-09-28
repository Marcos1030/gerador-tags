import streamlit as st
import os
from io import BytesIO
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from docx import Document
from docx.shared import Cm

st.set_page_config(page_title="Gerador de Tags - Cibele Doces", page_icon="🧁", layout="centered")

st.title("🧁 Gerador de Tags para Impressão")
st.write("Carregue a imagem da tag para gerar automaticamente a folha formatada em 4x4 cm (35 tags por página).")

# Upload da imagem único
uploaded_file = st.file_uploader("Escolha a imagem da tag (PNG ou JPG)", type=["png", "jpg", "jpeg"], key="tag_uploader")

if uploaded_file is not None:
    image = PILImage.open(uploaded_file)
    st.image(image, caption="Tag Carregada", width=200)
    
    st.write("---")
    st.subheader("Escolha o formato de saída:")
    
    col1, col2 = st.columns(2)
    
    # Opção para gerar PDF
    with col1:
        if st.button("Gerar em PDF"):
            pdf_buffer = BytesIO()
            largura_a4, altura_a4 = A4
            tamanho_pt = (4.0 / 2.54) * 72  # 4 cm em pontos
            margem = 10
            espacamento = 2
            
            c = canvas.Canvas(pdf_buffer, pagesize=A4)
            colunas = int((largura_a4 - (2 * margem)) // (tamanho_pt + espacamento))
            linhas = int((altura_a4 - (2 * margem)) // (tamanho_pt + espacamento))
            
            x_atual = margem
            y_atual = altura_a4 - margem - tamanho_pt
            
            temp_img_path = "temp_tag.png"
            image.save(temp_img_path)
            
            for i in range(colunas * linhas):
                c.drawImage(temp_img_path, x_atual, y_atual, width=tamanho_pt, height=tamanho_pt)
                x_atual += tamanho_pt + espacamento
                if (i + 1) % colunas == 0:
                    x_atual = margem
                    y_atual -= (tamanho_pt + espacamento)
            
            c.save()
            pdf_buffer.seek(0)
            
            st.success("PDF gerado com sucesso!")
            st.download_button(
                label="📥 Descarregar PDF",
                data=pdf_buffer,
                file_name="tags_4x4.pdf",
                mime="application/pdf"
            )

    # Opção para gerar Word (.docx) exatamente com 35 tags (5 colunas x 7 linhas)
    with col2:
        if st.button("Gerar em Word (.docx)"):
            doc = Document()
            
            # Margens milimetricamente ajustadas para caber perfeitamente a grelha 5x7
            for section in doc.sections:
                section.top_margin = Cm(0.5)
                section.bottom_margin = Cm(0.5)
                section.left_margin = Cm(0.5)
                section.right_margin = Cm(0.5)
            
            temp_img_path = "temp_tag.png"
            image.save(temp_img_path)
            
            # 5 colunas e 7 linhas = 35 tags por folha
            colunas = 5
            linhas = 7
            table = doc.add_table(rows=linhas, cols=colunas)
            
            for row in table.rows:
                for cell in row.cells:
                    p = cell.paragraphs[0]
                    run = p.add_run()
                    run.add_picture(temp_img_path, width=Cm(4.0), height=Cm(4.0))
            
            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            st.success("Documento Word gerado com sucesso!")
            st.download_button(
                label="📥 Descarregar Word",
                data=doc_buffer,
                file_name="tags_4x4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )