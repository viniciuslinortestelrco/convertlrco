import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

st.title("Conversor de PDF para Excel 📄➡️📊")

# Upload de múltiplos arquivos
uploaded_files = st.file_uploader("Selecione os arquivos PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    dados = []
    horario_re = r"\d{2}:\d{2}:\d{2}"
    registro_re = r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}"
    data_relatorio_re = r"\b\d{2}/\d{2}/\d{4}\b"

    for uploaded_file in uploaded_files:
        turma_atual = None
        nome_escola = "ESCOLA NÃO IDENTIFICADA"
        municipio = "MUNICÍPIO NÃO IDENTIFICADO"
        data_relatorio = "DATA NÃO IDENTIFICADA"

        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                texto = page.extract_text()
                if not texto:
                    continue
                linhas = texto.split("\n")

                if page_num == 0:
                    for i, linha in enumerate(linhas):
                        if "ESTADO DO PARANÁ" in linha:
                            match_data = re.search(data_relatorio_re, linha)
                            if match_data:
                                data_relatorio = match_data.group()
                        if "SECRETARIA DE ESTADO DA EDUCAÇÃO" in linha:
                            municipio = linha.split("SECRETARIA")[0].strip()
                            if i + 1 < len(linhas):
                                nome_escola = linhas[i + 1].strip()

                for linha in linhas:
                    linha = linha.strip()
                    if " - " in linha and "TURMA" not in linha and "LANÇAMENTO" not in linha:
                        turma_atual = linha
                        continue
                    if not turma_atual:
                        continue

                    horarios = re.findall(horario_re, linha)
                    registros = re.findall(registro_re, linha)

                    if not horarios:
                        continue

                    horario = horarios[0]
                    pos_horario = linha.find(horario)
                    pos_fim_horario = pos_horario + len(horario)

                    registro_aula = registros[0] if len(registros) >= 1 else "Sem registro"
                    registro_conteudo = registros[1] if len(registros) >= 2 else "Sem registro"

                    pos_registro = linha.find(registros[0]) if registros else len(linha)
                    disciplina = linha[pos_fim_horario:pos_registro].strip()

                    dados.append([
                        data_relatorio,
                        municipio,
                        nome_escola,
                        turma_atual,
                        horario,
                        disciplina,
                        registro_aula,
                        registro_conteudo
                    ])

    colunas = [
        "DATA DO RELATÓRIO", "MUNICÍPIO", "ESCOLA", "TURMA",
        "HORÁRIO", "DISCIPLINA", "REGISTRO DE AULA", "REGISTRO DE CONTEÚDO"
    ]
    df = pd.DataFrame(dados, columns=colunas)
    df = df[~df["DISCIPLINA"].str.contains("impresso por:", case=False, na=False)]

    # Excel em memória
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.success("Conversão concluída! 🎉")
    st.download_button("📥 Baixar Excel", data=output, file_name="relatorio_convertido.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
