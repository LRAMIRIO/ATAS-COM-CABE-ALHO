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
    Carregue primeiro a planilha **“DADOS DAS EMPRESAS.xlsx”**, que deve conter blocos de 6 linhas cada empresa:
    - Coluna B da linha 1: Razão Social
    - Coluna B da linha 2: CNPJ
    - Coluna B da linha 3: Endereço (completo)
    - Coluna B da linha 4: Telefone
    - Coluna B da linha 5: E-mail

    Depois, faça o upload das planilhas individuais (uma por empresa) que serão processadas. 
    O nome do arquivo deve corresponder à razão social (ou conter parte dela) para que a correspondência funcione.
    """
)

# Upload “DADOS DAS EMPRESAS.xlsx”
dados_empresas_file = st.file_uploader(
    "1) Selecione o arquivo DADOS DAS EMPRESAS.xlsx", 
    type="xlsx", 
    key="dados_empresas"
)

# Upload das planilhas separadas por empresa
arquivos_empresas = st.file_uploader(
    "2) Selecione as planilhas individuais por empresa (.xlsx)", 
    type="xlsx", 
    accept_multiple_files=True, 
    key="arquivos_empresas"
)

def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    return "".join(c for c in texto if c.isalnum() or c.isspace()).lower().strip()

def extrair_blocos_empresas(df: pd.DataFrame) -> dict:
    """
    Recebe um DataFrame sem cabeçalho e agrupa de 6 em 6 linhas,
    extraindo dados da coluna B (índice 1) para montar o dicionário.
    """
    dados = {}
    for i in range(0, len(df), 6):
        bloco = df.iloc[i : i + 6].reset_index(drop=True)
        if bloco.shape[0] < 2:
            continue
        # Se coluna B da primeira linha estiver vazia, ignora o bloco
        if pd.isna(bloco.iloc[0, 1]):
            continue
        nome_real = str(bloco.iloc[0, 1]).strip()
        dados[nome_real] = {
            "RAZÃO SOCIAL": nome_real,
            "CNPJ": str(bloco.iloc[1, 1]).strip() if bloco.shape[0] > 1 else "",
            "ENDEREÇO": str(bloco.iloc[2, 1]).strip() if bloco.shape[0] > 2 else "",
            "TELEFONE": str(bloco.iloc[3, 1]).strip() if bloco.shape[0] > 3 else "",
            "E-MAIL": str(bloco.iloc[4, 1]).strip() if bloco.shape[0] > 4 else "",
        }
    return dados

if dados_empresas_file and arquivos_empresas:
    try:
        # Lê o Excel sem cabeçalho
        df_empresas = pd.read_excel(dados_empresas_file, header=None)
    except Exception as e:
        st.error(f"Erro ao ler DADOS DAS EMPRESAS.xlsx: {e}")
        st.stop()

    # Extrair informações em blocos
    dados_empresas = extrair_blocos_empresas(df_empresas)
    if not dados_empresas:
        st.error("Não foram encontrados blocos válidos na planilha DADOS DAS EMPRESAS.xlsx.")
        st.stop()

    # Normalizar chaves
    dados_empresas_norm = { normalizar(nome): valores for nome, valores in dados_empresas.items() }

    st.subheader("Razões Sociais Detectadas:")
    cols = st.columns(2)
    for idx, nome in enumerate(dados_empresas.keys()):
        cols[idx % 2].write(f"- {nome}")

    # Preparar buffer ZIP em memória
    output_zip = BytesIO()
    match_log = []

    with ZipFile(output_zip, "w") as zipf:
        for arquivo in arquivos_empresas:
            nome_arquivo = arquivo.name
            nome_empresa = nome_arquivo.replace(".xlsx", "").strip()
            nome_norm = normalizar(nome_empresa)

            match = get_close_matches(nome_norm, dados_empresas_norm.keys(), n=1, cutoff=0.3)
            if not match:
                match_log.append(f"❌ NÃO ENCONTRADO: {nome_empresa} (normalizado: {nome_norm})")
                continue

            encontrado = match[0]
            dados = dados_empresas_norm[encontrado]
            match_log.append(f"✅ {nome_empresa} ➜ {dados['RAZÃO SOCIAL']}")

            try:
                wb = load_workbook(arquivo)
            except Exception as e:
                match_log.append(f"⚠️ Erro ao abrir {nome_arquivo}: {e}")
                continue

            ws = wb.active

            # Remove mesclagens pré-existentes
            for m in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(m))

            # Insere 5 linhas no topo para o cabeçalho
            ws.insert_rows(1, amount=5)

            # Mescla A1:H5 em única célula
            ws.merge_cells(start_row=1, start_column=1, end_row=5, end_column=8)
            cell = ws["A1"]
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            # Preenche a célula A1 com todas as informações, quebrando linha
            texto_cabecalho = (
                f"RAZÃO SOCIAL: {dados['RAZÃO SOCIAL']}
"
                f"CNPJ: {dados['CNPJ']}
"
                f"ENDEREÇO: {dados['ENDEREÇO']}
"
                f"TELEFONE: {dados['TELEFONE']}
"
                f"E-MAIL: {dados['E-MAIL']}"
            )
            cell.value = texto_cabecalho

            # Salva em buffer e adiciona ao ZIP
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            zipf.writestr(nome_arquivo, buffer.read())

    st.subheader("Relatório de Correspondências:")
    for log in match_log:
        st.write(log)

    # Finaliza o ZIP e exibe botão de download
    output_zip.seek(0)
    st.download_button(
        label="📥 Baixar ZIP com Planilhas Formatadas",
        data=output_zip.getvalue(),
        file_name="Planilhas_Cabecalho_Formatado.zip",
        mime="application/zip"
    )
else:
    st.info("Aguardando upload dos arquivos: primeiro DADOS DAS EMPRESAS, depois as planilhas por empresa.")
