import streamlit as st
import tempfile
from pathlib import Path
from pdf_to_xlsx import convert, parse_page, extract_doc_meta
import pdfplumber

st.set_page_config(page_title="PDF → XLSX", page_icon="📊", layout="centered")

st.title("📄 PDF → XLSX")
st.caption("Converte ensaios agrícolas (GENVCE) em planilha formatada")

pdf_file  = st.file_uploader("Selecione o PDF", type="pdf")
logo_file = st.file_uploader("Logo (opcional — PNG ou JPEG)", type=["png", "jpg", "jpeg"])

if pdf_file and st.button("Converter", type="primary", use_container_width=True):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        pdf_path  = tmp / pdf_file.name
        xlsx_path = tmp / pdf_file.name.replace(".pdf", ".xlsx")
        logo_path = None

        pdf_path.write_bytes(pdf_file.read())

        if logo_file:
            logo_path = tmp / logo_file.name
            logo_path.write_bytes(logo_file.read())

        with st.spinner("Lendo o PDF..."):
            try:
                # Mostra o que foi extraído antes de gerar o XLSX
                with pdfplumber.open(pdf_path) as pdf:
                    pages_data = [parse_page(p) for p in pdf.pages]

                st.info(f"**{len(pages_data)} página(s) detectada(s)**")
                for i, pd_ in enumerate(pages_data):
                    st.write(f"- Pág {i+1} · **{pd_['section'] or '(sem seção)'}** "
                             f"· {len(pd_['headers'])} colunas "
                             f"· {len(pd_['rows'])} linhas de dados")

                convert(pdf_path, xlsx_path, logo_path=logo_path)

                st.success("Pronto!")
                st.download_button(
                    label="⬇ Baixar XLSX",
                    data=xlsx_path.read_bytes(),
                    file_name=xlsx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro: {e}")
                st.exception(e)
