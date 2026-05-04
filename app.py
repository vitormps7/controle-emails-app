import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Controle de E-mails",
    page_icon="📧",
    layout="wide"
)

st.title("📧 Controle de E-mails")
st.write("Aplicativo para acompanhamento e análise dos atendimentos registrados em planilha.")

st.sidebar.header("📁 Carregar base")
arquivo = st.sidebar.file_uploader(
    "Carregue a planilha Excel",
    type=["xlsx"]
)

if arquivo is None:
    st.info("Carregue uma planilha Excel na barra lateral para começar.")
    st.stop()

try:
    df = pd.read_excel(arquivo)
except Exception as erro:
    st.error("Não foi possível ler a planilha.")
    st.write(erro)
    st.stop()

st.success("Planilha carregada com sucesso!")

# Ajuste visual básico
df.columns = [str(col).strip() for col in df.columns]

st.sidebar.header("🔎 Filtros")

df_filtrado = df.copy()

# Filtro por Situação
if "Situação" in df.columns:
    situacoes = sorted(df["Situação"].dropna().astype(str).unique())
    filtro_situacao = st.sidebar.multiselect("Situação", situacoes)

    if filtro_situacao:
        df_filtrado = df_filtrado[
            df_filtrado["Situação"].astype(str).isin(filtro_situacao)
        ]

# Filtro por Servidor(a)
if "Servidor(a)" in df.columns:
    servidores = sorted(df["Servidor(a)"].dropna().astype(str).unique())
    filtro_servidor = st.sidebar.multiselect("Servidor(a)", servidores)

    if filtro_servidor:
        df_filtrado = df_filtrado[
            df_filtrado["Servidor(a)"].astype(str).isin(filtro_servidor)
        ]

# Filtro por Fonte
if "Fonte" in df.columns:
    fontes = sorted(df["Fonte"].dropna().astype(str).unique())
    filtro_fonte = st.sidebar.multiselect("Fonte", fontes)

    if filtro_fonte:
        df_filtrado = df_filtrado[
            df_filtrado["Fonte"].astype(str).isin(filtro_fonte)
        ]

# Filtro por Assunto
if "Assunto" in df.columns:
    assuntos = sorted(df["Assunto"].dropna().astype(str).unique())
    filtro_assunto = st.sidebar.multiselect("Assunto", assuntos)

    if filtro_assunto:
        df_filtrado = df_filtrado[
            df_filtrado["Assunto"].astype(str).isin(filtro_assunto)
        ]

# Busca livre
busca = st.sidebar.text_input("Pesquisar na base")

if busca:
    busca_lower = busca.lower()
    df_filtrado = df_filtrado[
        df_filtrado.astype(str).apply(
            lambda linha: linha.str.lower().str.contains(busca_lower, na=False).any(),
            axis=1
        )
    ]

# Indicadores
st.subheader("📊 Resumo geral")

total_registros = len(df)
total_filtrado = len(df_filtrado)

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total da base", total_registros)
col2.metric("Resultado filtrado", total_filtrado)

if "Situação" in df.columns:
    concluidos = df[df["Situação"].astype(str).str.contains("concl|finaliz|resol", case=False, na=False)]
    pendentes = df[df["Situação"].astype(str).str.contains("pend|abert|andamento", case=False, na=False)]

    col3.metric("Concluídos", len(concluidos))
    col4.metric("Pendentes / Em andamento", len(pendentes))
else:
    col3.metric("Concluídos", "-")
    col4.metric("Pendentes / Em andamento", "-")

st.divider()

# Tabela principal
st.subheader("📋 Registros encontrados")
st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

# Resumos
st.divider()
st.subheader("📌 Resumos")

aba1, aba2, aba3 = st.tabs(["Por situação", "Por servidor", "Por fonte"])

with aba1:
    if "Situação" in df_filtrado.columns:
        resumo_situacao = (
            df_filtrado["Situação"]
            .fillna("Não informado")
            .astype(str)
            .value_counts()
            .reset_index()
        )
        resumo_situacao.columns = ["Situação", "Quantidade"]
        st.dataframe(resumo_situacao, use_container_width=True, hide_index=True)
    else:
        st.warning("Coluna 'Situação' não encontrada na planilha.")

with aba2:
    if "Servidor(a)" in df_filtrado.columns:
        resumo_servidor = (
            df_filtrado["Servidor(a)"]
            .fillna("Não informado")
            .astype(str)
            .value_counts()
            .reset_index()
        )
        resumo_servidor.columns = ["Servidor(a)", "Quantidade"]
        st.dataframe(resumo_servidor, use_container_width=True, hide_index=True)
    else:
        st.warning("Coluna 'Servidor(a)' não encontrada na planilha.")

with aba3:
    if "Fonte" in df_filtrado.columns:
        resumo_fonte = (
            df_filtrado["Fonte"]
            .fillna("Não informado")
            .astype(str)
            .value_counts()
            .reset_index()
        )
        resumo_fonte.columns = ["Fonte", "Quantidade"]
        st.dataframe(resumo_fonte, use_container_width=True, hide_index=True)
    else:
        st.warning("Coluna 'Fonte' não encontrada na planilha.")

# Download da base filtrada
st.divider()
st.subheader("⬇️ Baixar resultado filtrado")

buffer = BytesIO()

with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name="Resultado filtrado")

st.download_button(
    label="Baixar em Excel",
    data=buffer.getvalue(),
    file_name="resultado_filtrado_controle_emails.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
