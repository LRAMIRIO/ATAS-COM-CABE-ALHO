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

    **Importante:**  
    O conteúdo de cada planilha “empresa.xlsx” começa efetivamente na **linha 6** (dados de colunas A, B, C etc.).  
    Este script irá “ler” qual é a última coluna não vazia nessa linha 6 antes de inserir os 5 cabeçalhos.  
    Assim, será capaz de mesclar as células de cabeçalho apenas até aquela última coluna dinâmica (A até, por exemplo, G ou H ou I).
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
    # Tenta ler “DADOS DAS EMPRESAS.xlsx” sem cabeçalho
    try:
        df_empresas = pd.read_excel(dados_empresas_file, header=None)
    except Exception as e:
        st.error(f"❌ Erro ao ler “DADOS DAS EMPRESAS.xlsx”: {e}")
        st.stop()

    # Extrai blocos de 6 linhas
    dados_empresas = extrair_blocos_empresas(df_empresas)
    if not dados_empresas:
        st.error("❌ Não foram encontrados blocos válidos na planilha “DADOS DAS EMPRESAS.xlsx”.")
        st.stop()

    # Normaliza nomes para correspondência
    dados_empresas_norm = { normalizar(nome): info for nome, info in dados_empresas.items() }

    st.subheader("Razões Sociais Detectadas")
    col1, col2 = st.columns(2)
    for idx, razao in enumerate(dados_empresas.keys()):
        if idx % 2 == 0:
            col1.write(f"- {razao}")
        else:
            col2.write(f"- {razao}")

    # ====== COMEÇA A GERAÇÃO DO ZIP EM MEMÓRIA ======
    output_zip = BytesIO()
    match_log = []

    with ZipFile(output_zip, "w") as zipf:
        for arquivo in arquivos_empresas:
            nome_arquivo = arquivo.name
            base_nome = nome_arquivo.replace(".xlsx", "").strip()
            nome_norm = normalizar(base_nome)

            # Busca correspondência entre nome de arquivo e razão social
            matches = get_close_matches(nome_norm, dados_empresas_norm.keys(), n=1, cutoff=0.3)
            if not matches:
                match_log.append(f"❌ NÃO ENCONTRADO: {base_nome} (normalizado: {nome_norm})")
                continue

            chave = matches[0]
            info = dados_empresas_norm[chave]
            match_log.append(f"✅ {base_nome} → {info['RAZAO_SOCIAL']}")

            # Carrega a planilha de cada empresa
            try:
                wb = load_workbook(arquivo)
            except Exception as e:
                match_log.append(f"⚠️ Erro ao abrir “{nome_arquivo}”: {e}")
                continue

            ws = wb.active

            # 1) Antes de inserir linhas, vamos detectar quantas colunas há a partir da linha 6
            #    Isso considera que a linha 6 (antes da inserção) contém dados em colunas A, B, C, ... até a última
            #    Após inserirmos cabeçalho, a linha 6 se tornará a linha 11, mas vamos buscar agora:
            #    - Faça um loop sobre as células da linha 6 original (ou seja, ws[“6”])
            #    - Descubra o índice da última coluna que não está em branco
            last_col = 1  # Mínimo: coluna A
            # Acesso direto às células da linha 6
            for cell in ws[6]:
                if cell.value is not None and str(cell.value).strip() != "":
                    # cell.column retorna a letra (“A”, “B” etc.); para índice numérico, use cell.column_letter e depois converter
                    # No openpyxl v3.x, use cell.column (inteiro). Caso seja string, converta letra para número.
                    if isinstance(cell.column, int):
                        idx_col = cell.column
                    else:
                        # Converter letra para índice: ord("A")→65, etc.
                        idx_col = ord(cell.column.upper()) - ord("A") + 1
                    last_col = max(last_col, idx_col)

            # 2) Agora inserimos 5 linhas em branco no topo para o cabeçalho
            ws.insert_rows(1, amount=5)

            # 3) Mesclamos cada uma das linhas 1 a 5, da coluna A até last_col
            #    (Se last_col = 7, mescla A1→G1, A2→G2, etc.)
            #    Ajusta alinhamento conforme pedido
            # LINHA 1 (Razão Social)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
            c1 = ws.cell(row=1, column=1)
            c1.alignment = Alignment(horizontal="left", vertical="center")
            c1.value = f"RAZÃO SOCIAL: {info['RAZAO_SOCIAL']}"

            # LINHA 2 (CNPJ)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
            c2 = ws.cell(row=2, column=1)
            c2.alignment = Alignment(horizontal="left", vertical="center")
            c2.value = f"CNPJ: {info['CNPJ']}"

            # LINHA 3 (Endereço)
            ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
            c3 = ws.cell(row=3, column=1)
            c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c3.value = f"ENDEREÇO: {info['ENDERECO']}"

            # LINHA 4 (Telefone)
            ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_col)
            c4 = ws.cell(row=4, column=1)
            c4.alignment = Alignment(horizontal="left", vertical="center")
            c4.value = f"TELEFONE: {info['TELEFONE']}"

            # LINHA 5 (E-mail)
            ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last_col)
            c5 = ws.cell(row=5, column=1)
            c5.alignment = Alignment(horizontal="left", vertical="center")
            c5.value = f"E-MAIL: {info['E-MAIL']}"

            # 4) Por fim, salva esta planilha em memória e adiciona ao ZIP
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
