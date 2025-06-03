import streamlit as st
import pandas as pd
import unicodedata
from difflib import get_close_matches
from zipfile import ZipFile
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="Gerador de Planilhas com Cabeçalho", layout="wide")
st.title("Gerador de Planilhas com Cabeçalho para Empresas")

st.markdown(
    """
    1. Faça upload da planilha **“DADOS DAS EMPRESAS.xlsx”**, que deve ter blocos de 6 linhas para cada empresa:
       - Coluna B da linha 1: Razão Social
       - Coluna B da linha 2: CNPJ
       - Coluna B da linha 3: Endereço (completo)
       - Coluna B da linha 4: Telefone
       - Coluna B da linha 5: E-mail

    2. Depois, faça upload das planilhas individuais (uma por empresa). O nome do arquivo deve corresponder, ao menos parcialmente, à razão social para que o app consiga fazer a correspondência.
    """
)

# 1) Upload da planilha de dados das empresas
dados_empresas_file = st.file_uploader(
    "Selecione o arquivo DADOS DAS EMPRESAS.xlsx",
    type="xlsx",
    key="dados_empresas"
)

# 2) Upload das planilhas separadas por empresa
arquivos_empresas = st.file_uploader(
    "Selecione as planilhas individuais por empresa (.xlsx)",
    type="xlsx",
    accept_multiple_files=True,
    key="arquivos_empresas"
)

def normalizar(texto: str) -> str:
    """Remove acentos e caracteres especiais, deixa tudo minúsculo e sem pontuação."""
    texto_nfkd = unicodedata.normalize("NFKD", texto)
    texto_ascii = texto_nfkd.encode("ASCII", "ignore").decode("utf-8")
    return "".join(c for c in texto_ascii if c.isalnum() or c.isspace()).lower().strip()

def extrair_blocos_empresas(df: pd.DataFrame) -> dict:
    """
    Recebe um DataFrame sem cabeçalho e cria um dicionário:
      chave = razão social (coluna B da linha 1 de cada bloco de 6)
      valor = { "RAZÂO_SOCIAL": ..., "CNPJ": ..., "ENDERECO": ..., "TELEFONE": ..., "EMAIL": ... }
    Cada bloco de 6 linhas corresponde a uma empresa.
    """
    dados = {}
    for i in range(0, len(df), 6):
        bloco = df.iloc[i : i + 6].reset_index(drop=True)
        # Se não existir ao menos 2 linhas ou se a célula B1 estiver vazia, pula
        if bloco.shape[0] < 2 or pd.isna(bloco.iloc[0, 1]):
            continue
        nome_real = str(bloco.iloc[0, 1]).strip()
        dados[nome_real] = {
            "RAZÃO_SOCIAL": nome_real,
            "CNPJ": str(bloco.iloc[1, 1]).strip() if bloco.shape[0] > 1 else "",
            "ENDEREÇO": str(bloco.iloc[2, 1]).strip() if bloco.shape[0] > 2 else "",
            "TELEFONE": str(bloco.iloc[3, 1]).strip() if bloco.shape[0] > 3 else "",
            "E-MAIL": str(bloco.iloc[4, 1]).strip() if bloco.shape[0] > 4 else "",
        }
    return dados

if dados_empresas_file and arquivos_empresas:
    try:
        # Lê o Excel (sem cabeçalho!), para poder iterar de 6 em 6 linhas
        df_empresas = pd.read_excel(dados_empresas_file, header=None)
    except Exception as e:
        st.error(f"❌ Erro ao ler “DADOS DAS EMPRESAS.xlsx”: {e}")
        st.stop()

    # Extrai dados em blocos de 6 linhas
    dados_empresas = extrair_blocos_empresas(df_empresas)
    if not dados_empresas:
        st.error("❌ Não foram encontrados blocos válidos na planilha de empresas.")
        st.stop()

    # Normaliza as chaves (razões sociais)
    dados_empresas_norm = {
        normalizar(nome): valores
        for nome, valores in dados_empresas.items()
    }

    st.subheader("Razões Sociais Detectadas")
    col1, col2 = st.columns(2)
    for idx, nome in enumerate(dados_empresas.keys()):
        if idx % 2 == 0:
            col1.write(f"- {nome}")
        else:
            col2.write(f"- {nome}")

    # ====== COMEÇA A GERAÇÃO DO ZIP EM MEMÓRIA ======
    output_zip = BytesIO()
    match_log = []

    with ZipFile(output_zip, "w") as zipf:
        for arquivo in arquivos_empresas:
            nome_arquivo = arquivo.name
            nome_empresa = nome_arquivo.replace(".xlsx", "").strip()
            nome_norm = normalizar(nome_empresa)

            # Tenta buscar a razão social correspondente
            match = get_close_matches(nome_norm, dados_empresas_norm.keys(), n=1, cutoff=0.3)
            if not match:
                match_log.append(f"❌ NÃO ENCONTRADO: {nome_empresa} (normalizado: {nome_norm})")
                continue

            chave = match[0]
            info = dados_empresas_norm[chave]
            match_log.append(f"✅ {nome_empresa} → {info['RAZÃO_SOCIAL']}")

            try:
                wb = load_workbook(arquivo)
            except Exception as e:
                match_log.append(f"⚠️ Erro ao abrir “{nome_arquivo}”: {e}")
                continue

            ws = wb.active

            # Remove mesclagens pré-existentes (caso existam)
            for m in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(m))

            # Insere 5 linhas vazias no topo para o cabeçalho
            ws.insert_rows(1, amount=5)

            # Mescla A1:H5 em uma única célula
            ws.merge_cells(start_row=1, start_column=1, end_row=5, end_column=8)
            cell = ws.cell(row=1, column=1)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            # Preenche a célula A1 com todos os dados, separados por nova linha
            texto_cabecalho = (
                f"RAZÃO SOCIAL: {info['RAZÃO_SOCIAL']}\n"
                f"CNPJ: {info['CNPJ']}\n"
                f"ENDEREÇO: {info['ENDEREÇO']}\n"
                f"TELEFONE: {info['TELEFONE']}\n"
                f"E-MAIL: {info['E-MAIL']}"
            )
            cell.value = texto_cabecalho

            # Salva a planilha modificada em memória e adiciona ao ZIP
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            zipf.writestr(nome_arquivo, buffer.read())

    # Exibe o log de correspondência
    st.subheader("Log de Correspondências")
    for linha in match_log:
        st.write(linha)

    # Botão de download do ZIP final
    output_zip.seek(0)
    st.download_button(
        label="📥 Baixar ZIP com Planilhas Formatadas",
        data=output_zip.getvalue(),
        file_name="Planilhas_Cabecalho_Formatado.zip",
        mime="application/zip"
    )

else:
    st.info("Aguardando upload:\n\n1. DADOS DAS EMPRESAS.xlsx\n2. Planilhas por empresa (.xlsx)")
