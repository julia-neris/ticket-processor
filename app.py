import streamlit as st
import pdfplumber
import re
import pandas as pd

st.title("Processador de NFS-e")

arquivos = st.file_uploader(
    "Selecione os PDFs",
    type="pdf",
    accept_multiple_files=True
)

dados = []

if arquivos:

    for arquivo in arquivos:

        with pdfplumber.open(arquivo) as pdf:

            texto = ""

            for pagina in pdf.pages:
                texto += pagina.extract_text()

            nfse = re.findall(r"\d+", arquivo.name)[0]

            razoes = re.findall(r"Nome/Razão Social:\s*(.+)", texto)
            razao = razoes[1] if len(razoes) >= 2 else razoes[0]

            match = re.search(r"Número DPS / Série DPS.*?(\d+)\s*/\s*(\d+)", texto, re.DOTALL)

            dps = ""
            if match:
                dps = f"{match.group(1)} / {match.group(2)}"

            dados.append({
                "Razão Social": razao,
                "NFS-e": nfse,
                "DPS / Série": dps
            })

    df = pd.DataFrame(dados)

    st.dataframe(df)

    excel = df.to_excel("resultado.xlsx", index=False)

    with open("resultado.xlsx", "rb") as f:
        st.download_button(
            "Baixar Excel",
            f,
            "resultado.xlsx"
        )