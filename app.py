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
    1. Faça upload da planilha **“DADOS DAS EMPRESAS.xlsx”**, que deve conter blocos de 6 linhas para cada empresa:
       - Coluna B da linha 1 de cada bloco: **Razão Social**
       - Coluna B da linha 2 de cada bloco: **CNPJ**
       - Coluna B da linha 3 de cada bloco: **Endereço (completo)**
       - Coluna B da linha 4 de cada bloco: **Telefone**
       - Coluna B da linha 5 de cada bloco: **E-mail**

    2. Depois, faça upload das planilhas individuais (uma por empresa).  
       O nome de cada arquivo deve corresponder, ao menos parcialmente, à razão social para que a correspondência funcione.

    A partir da linha 6 de cada planilha haverá dados em um certo número de colunas (por exemplo, colunas A até G, ou até H, ou até I).  
    Este script irá detectar quantas colunas realmente têm conteúdo a partir da linha 6 e mesclar apenas até esse último índice de coluna.
    """
)

# 1) Upload da planilha “DADOS DAS EMPRESAS.xlsx”
dados_empresas_file = st.file_uploader(
    "1) Selecione o arquivo DADOS DAS EMPRESAS.xlsx",
    type="xlsx",
    key="dados_empresas"
)

# 2) Upload das planilhas separadas por empresa
arquivos_empresas = st.file_uploader(
    "2) Selecione as planilhas individuais por empresa (.xlsx)",
    type="xlsx",
    accept_multiple_files=True,
    key="arquivos_empresas"
)

def normalizar(texto: str) -> str:
    """
    Remove acentos e caracteres não alfanuméricos, deixa tudo minúsculo e sem pontuação.
    """
    nfkd = unicodedata.normalize("NFKD", texto)
    ascii_txt = nfkd.encode("ASCII", "ignore").decode("utf-8")
    return "".join(c for c in ascii_txt if c.isalnum() or c.isspace()).lower().strip()

def extrair_blocos_empresas(df: pd.DataFrame) -> dict:
    """
    Recebe um DataFrame sem cabeçalho e agrupa de 6 em 6 linhas, extraindo da coluna B (índice 1):
      - Linha 1 de cada bloco: Razão Social
      - Linha 2 de cada bloco: CNPJ
      - Linha 3 de cada bloco: Endereço completo
      - Linha 4 de cada bloco: Telefone
      - Linha 5 de cada bloco: E-mail
    Retorna um dicionário:
      { "Razão Social": { "RAZAO_SOCIAL": ..., "CNPJ": ..., "ENDERECO": ..., "TELEFONE": ..., "EMAIL": ... }, ... }
    """
    dados = {}
    for i in range(0, len(df), 6):
        bloco = df.iloc[i : i + 6].reset_index(drop=True)
        if bloco.shape[0] < 2 or pd.isna(bloco.iloc[0, 1]):
            continue
        razao = str(bloco.iloc[0, 1]).strip()
        dados[razao] = {
            "RAZAO_SOCIAL": razao,
            "CNPJ": str(bloco.iloc[1, 1]).strip() if bloco.shape[0] > 1 else "",
            "ENDERECO": str(bloco.iloc[2, 1]).strip() if bloco.shape[0] > 2 else "",
            "TELEFONE": str(bloco.iloc[3, 1]).strip() if bloco.shape[0] > 3 else "",
            "E-MAIL": str(bloco.iloc[4, 1]).strip() if bloco.shape[0] > 4 else "",
        }
    return dados

if dados_empresas_file and arquivos_empresas:
    # Tenta ler o Excel sem cabeçalho
    try:
        df_empresas = pd.read_excel(dados_empresas_file, header=None)
    except Exception as e:
        st.error(f"❌ Erro ao ler ‘DADOS DAS EMPRESAS.xlsx’: {e}")
        st.stop()

    # Extrai blocos de 6 linhas
    dados_empresas = extrair_blocos_empresas(df_empresas)
    if not dados_empresas:
        st.error("❌ Não foram encontrados blocos válidos na planilha DADOS DAS EMPRESAS.xlsx.")
        st.stop()

    # Normaliza nomes para correspondência
    dados_empresas_norm = { normalizar(nome): info for nome, info in dados_empresas.items() }

    st.subheader("Razões Sociais Detectadas")
    coluna1, coluna2 = st.columns(2)
    for idx, razao in enumerate(dados_empresas.keys()):
        if idx % 2 == 0:
            coluna1.write(f"- {razao}")
        else:
            coluna2.write(f"- {razao}")

    # Preparar buffer ZIP em memória
    output_zip = BytesIO()
    match_log = []

    with ZipFile(output_zip, "w") as zipf:
        for arquivo in arquivos_empresas:
            nome_arquivo = arquivo.name
            base_nome = nome_arquivo.replace(".xlsx", "").strip()
            nome_norm = normalizar(base_nome)

            # Busca correspondência
            matches = get_close_matches(nome_norm, dados_empresas_norm.keys(), n=1, cutoff=0.3)
            if not matches:
                match_log.append(f"❌ NÃO ENCONTRADO: {base_nome} (normalizado: {nome_norm})")
                continue

            chave = matches[0]
            info = dados_empresas_norm[chave]
            match_log.append(f"✅ {base_nome} → {info['RAZAO_SOCIAL']}")

            # Carrega a planilha da empresa
            try:
                wb = load_workbook(arquivo)
            except Exception as e:
                match_log.append(f"⚠️ Erro ao abrir ‘{nome_arquivo}’: {e}")
                continue

            ws = wb.active

            # Remove mesclagens pré-existentes
            for m in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(m))

            # INSERIR 5 LINHAS PARA CABEÇALHO
            ws.insert_rows(1, amount=5)

            # Determinar até qual coluna há dados a partir da linha 6
            last_col = 1  # mínimo será 1
            # Verifica a linha 6 para encontrar última coluna não vazia
            for idx_col, val in enumerate(ws.iter_rows(min_row=6, max_row=6, values_only=True)[0], start=1):
                if val is not None and str(val).strip() != "":
                    last_col = idx_col

            # Mescla e preenche cada linha de cabeçalho de acordo com last_col
            # Linha 1: Razão Social (A1 até última coluna)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
            cell1 = ws.cell(row=1, column=1)
            cell1.alignment = Alignment(horizontal="left", vertical="center")
            cell1.value = f"RAZÃO SOCIAL: {info['RAZAO_SOCIAL']}"

            # Linha 2: CNPJ
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
            cell2 = ws.cell(row=2, column=1)
            cell2.alignment = Alignment(horizontal="left", vertical="center")
            cell2.value = f"CNPJ: {info['CNPJ']}"

            # Linha 3: Endereço (pode envolver quebra de texto se for muito longo)
            ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
            cell3 = ws.cell(row=3, column=1)
            cell3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell3.value = f"ENDEREÇO: {info['ENDERECO']}"

            # Linha 4: Telefone
            ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_col)
            cell4 = ws.cell(row=4, column=1)
            cell4.alignment = Alignment(horizontal="left", vertical="center")
            cell4.value = f"TELEFONE: {info['TELEFONE']}"

            # Linha 5: E-mail
            ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last_col)
            cell5 = ws.cell(row=5, column=1)
            cell5.alignment = Alignment(horizontal="left", vertical="center")
            cell5.value = f"E-MAIL: {info['E-MAIL']}"

            # Salva a planilha modificada em memória e adiciona ao ZIP
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            zipf.writestr(nome_arquivo, buffer.read())

    # Exibe o log de correspondências
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
