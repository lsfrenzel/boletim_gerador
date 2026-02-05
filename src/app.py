import streamlit as st
import pandas as pd
import os
from main import generate_bulletins

st.set_page_config(page_title="Gerador de Boletins", layout="wide")

st.title("🎓 Gerador de Boletins Escolares")
st.markdown("""
Esta ferramenta gera boletins em PDF a partir de um arquivo Excel.
O arquivo deve conter as colunas: **Turma, Nome do Aluno, Unidade Curricular, Nota, Frequência Geral (%)**.
""")

uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx)", type="xlsx")

if uploaded_file is not None:
    # Salvar temporariamente o arquivo para processamento
    with open("temp_notas.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    df = pd.read_excel("temp_notas.xlsx")
    st.write("### Prévia dos Dados")
    st.dataframe(df.head())
    
    if st.button("Gerar Boletins PDF"):
        with st.spinner("Gerando PDF..."):
            try:
                generate_bulletins("temp_notas.xlsx")
                if os.path.exists("boletins_alunos.pdf"):
                    with open("boletins_alunos.pdf", "rb") as f:
                        st.download_button(
                            label="📥 Baixar Boletins (PDF)",
                            data=f,
                            file_name="boletins_alunos.pdf",
                            mime="application/pdf"
                        )
                    st.success("PDF gerado com sucesso!")
                else:
                    st.error("Erro ao gerar o PDF. Verifique os dados do arquivo.")
            except Exception as e:
                st.error(f"Erro inesperado: {str(e)}")
            finally:
                if os.path.exists("temp_notas.xlsx"):
                    os.remove("temp_notas.xlsx")

st.sidebar.header("Ajuda")
st.sidebar.info("""
1. Faça o upload do arquivo Excel.
2. Verifique se os dados estão corretos.
3. Clique em 'Gerar Boletins PDF'.
4. Baixe o arquivo gerado.
""")
