import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime
import re

# ============================================================
# CONTROLE DE ATENDIMENTOS - SEPRO
# Aplicativo Streamlit baseado na planilha de controle de e-mails.
# ============================================================

st.set_page_config(
    page_title="Controle de Atendimentos - SEPRO",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------
# Listas oficiais do aplicativo
# -----------------------------

SITUACOES = ["Pendente", "Em atendimento", "Respondido", "Outro", "Arquivado"]
FONTES = ["E-mail", "Telefone"]

# Assuntos extraídos da planilha-base.
ASSUNTOS = ['ADVOGADO DATIVO', 'AIJE', 'AIME', 'APOIAMENTO DE PARTIDO', 'ARQUIVAMENTO', 'ASE', 'Ação Penal', 'CADIN', 'CARTA DE ORDEM', 'CARTA PRECATÓRIA', 'CONTA DEPÓSITO JUDICIAL', 'CONTAGEM DE PRAZOS PROCESSUAIS', 'Cumprimento de Sentença', 'DECLINAÇÃO DE COMPETÊNCIA', 'DEFENSOR DATIVO', 'Depósito Judicial', 'Destinação Valores Feitos Criminais', 'Destinação Valores Procedimentos Criminais', 'Diplomação', 'DÚVIDAS PROCESSUAIS', 'ELO', 'EXECUÇÃO DE MULTA', 'EXECUÇÃO FISCAL', 'EXECUÇÃO PENAL', 'INFOJUD', 'Inquérito', 'JUIZ DE GARANTIA', 'LISTA DE APOIAMENTO', 'Medidas despenalizadoras', 'MULTA', 'Núcleo das Garantias', 'Outros', 'Parcelamento de Multa Eleitoral', 'PRESTAÇÃO DE CONTAS', 'Prestação de Contas Anual', 'Prestação de Contas Eleitorais', 'PROCESSO DE CRIAÇÃO DE PARTIDO', 'PROCESSO INVESTIGATÓRIO CRIMINAL', 'PROVIMENTO 01/2025', 'RENAJUD', 'REPRESENTAÇÃO', 'Representação - Direito de resposta', 'Representação - Propaganda', 'RROPCO', 'SERASAJUD', 'SICO', 'Sisbajud', 'Sistemas', 'TCO', 'TRANSAÇÃO PENAL', 'TRÂNSITO EM JULGADO']

# Zonas eleitorais da Bahia para seleção no sistema.
# Mantive "NA" para casos sem zona eleitoral definida.
ZONAS_ELEITORAIS_BA = ["NA"] + [str(i).zfill(3) for i in range(1, 206)]

SERVIDORES_INICIAIS = ['Andréa', 'Arivaldo', 'Arnaldo', 'Camille', 'Elisa', 'Equipe', 'Gabriele', 'José Carlos', 'Lorena', 'Mariana', 'Milena', 'Raphaela', 'Ricardo', 'Robelza', 'Sandra', 'Sandra Força', 'Vitor']

COLUNAS_BASE = [
    "SITUAÇÃO",
    "ZE",
    "REMETENTE",
    "DATA",
    "SERVIDOR(A)",
    "FONTE",
    "ASSUNTO",
    "OBSERVAÇÕES",
]


# -----------------------------
# Funções auxiliares
# -----------------------------

def limpar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    mapa = {
        "SITUAÇÃO": "SITUAÇÃO",
        "SITUACAO": "SITUAÇÃO",
        "ZE": "ZE",
        "ZONA": "ZE",
        "ZONA ELEITORAL": "ZE",
        "REMETENTE": "REMETENTE",
        "ORIGEM": "REMETENTE",
        "ORIGINADOR": "REMETENTE",
        "DATA": "DATA",
        "SERVIDOR(A)": "SERVIDOR(A)",
        "SERVIDOR": "SERVIDOR(A)",
        "FONTE": "FONTE",
        "ASSUNTO": "ASSUNTO",
        "OBSERVAÇÕES": "OBSERVAÇÕES",
        "OBSERVACOES": "OBSERVAÇÕES",
        "OBS": "OBSERVAÇÕES",
    }

    novas = []
    for c in df.columns:
        novas.append(mapa.get(c, c))
    df.columns = novas

    for c in COLUNAS_BASE:
        if c not in df.columns:
            df[c] = ""

    df = df[COLUNAS_BASE + [c for c in df.columns if c not in COLUNAS_BASE]]

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.date
    df["SITUAÇÃO"] = df["SITUAÇÃO"].fillna("").astype(str).str.strip()
    df["ZE"] = df["ZE"].fillna("NA").astype(str).str.strip()
    df["REMETENTE"] = df["REMETENTE"].fillna("").astype(str).str.strip()
    df["SERVIDOR(A)"] = df["SERVIDOR(A)"].fillna("").astype(str).str.strip()
    df["FONTE"] = df["FONTE"].fillna("").astype(str).str.strip()
    df["ASSUNTO"] = df["ASSUNTO"].fillna("").astype(str).str.strip()
    df["OBSERVAÇÕES"] = df["OBSERVAÇÕES"].fillna("").astype(str).str.strip()

    # Normaliza zona eleitoral: 1 -> 001, 42 -> 042, NA permanece NA.
    def normalizar_ze(valor):
        texto = str(valor).strip()
        if texto.upper() in ["", "NAN", "NA", "N/A", "NÃO SE APLICA"]:
            return "NA"
        numeros = re.sub(r"[^0-9]", "", texto)
        if numeros:
            return numeros.zfill(3)
        return texto

    df["ZE"] = df["ZE"].apply(normalizar_ze)
    return df


def criar_base_vazia() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUNAS_BASE)


def criar_servidores_iniciais() -> pd.DataFrame:
    return pd.DataFrame({
        "SERVIDOR(A)": SERVIDORES_INICIAIS,
        "ATIVO": [True] * len(SERVIDORES_INICIAIS)
    })


def obter_servidores_ativos() -> list:
    servidores = st.session_state.servidores.copy()
    if servidores.empty:
        return []
    servidores["SERVIDOR(A)"] = servidores["SERVIDOR(A)"].astype(str).str.strip()
    ativos = servidores[servidores["ATIVO"] == True]["SERVIDOR(A)"].dropna().tolist()
    return sorted(list(set([s for s in ativos if s])), key=str.casefold)


def gerar_excel_relatorio(df: pd.DataFrame, servidores: pd.DataFrame) -> bytes:
    buffer = BytesIO()

    base = df.copy()
    if "DATA" in base.columns:
        base["DATA"] = pd.to_datetime(base["DATA"], errors="coerce")

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        base.to_excel(writer, index=False, sheet_name="Base")

        if not base.empty:
            resumo_situacao = base["SITUAÇÃO"].fillna("Não informado").value_counts().reset_index()
            resumo_situacao.columns = ["Situação", "Quantidade"]
            resumo_situacao.to_excel(writer, index=False, sheet_name="Resumo por Situação")

            resumo_servidor = base["SERVIDOR(A)"].fillna("Não informado").value_counts().reset_index()
            resumo_servidor.columns = ["Servidor(a)", "Quantidade"]
            resumo_servidor.to_excel(writer, index=False, sheet_name="Resumo por Servidor")

            resumo_assunto = base["ASSUNTO"].fillna("Não informado").value_counts().reset_index()
            resumo_assunto.columns = ["Assunto", "Quantidade"]
            resumo_assunto.to_excel(writer, index=False, sheet_name="Resumo por Assunto")

            resumo_ze = base["ZE"].fillna("Não informado").value_counts().reset_index()
            resumo_ze.columns = ["Zona Eleitoral", "Quantidade"]
            resumo_ze.to_excel(writer, index=False, sheet_name="Resumo por ZE")

            resumo_fonte = base["FONTE"].fillna("Não informado").value_counts().reset_index()
            resumo_fonte.columns = ["Fonte", "Quantidade"]
            resumo_fonte.to_excel(writer, index=False, sheet_name="Resumo por Fonte")

            mensal = (
                base.dropna(subset=["DATA"])
                .assign(MÊS=lambda x: x["DATA"].dt.to_period("M").astype(str))
                .groupby("MÊS", as_index=False)
                .size()
                .rename(columns={"size": "Quantidade"})
            )
            mensal.to_excel(writer, index=False, sheet_name="Resumo Mensal")

            pendentes = base[base["SITUAÇÃO"].astype(str).str.contains("pend|andamento", case=False, na=False)]
            pendentes.to_excel(writer, index=False, sheet_name="Pendentes")

        servidores.to_excel(writer, index=False, sheet_name="Servidores")

        workbook = writer.book
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    valor = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(valor))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
                cell.alignment = cell.alignment.copy(horizontal="center")

    return buffer.getvalue()


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    filtrado = df.copy()

    st.sidebar.markdown("## 🔎 Filtros")

    if not filtrado.empty:
        data_min = pd.to_datetime(filtrado["DATA"], errors="coerce").min()
        data_max = pd.to_datetime(filtrado["DATA"], errors="coerce").max()

        if pd.notna(data_min) and pd.notna(data_max):
            intervalo = st.sidebar.date_input(
                "Período",
                value=(data_min.date(), data_max.date()),
                min_value=data_min.date(),
                max_value=data_max.date()
            )
            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                inicio, fim = intervalo
                datas = pd.to_datetime(filtrado["DATA"], errors="coerce").dt.date
                filtrado = filtrado[(datas >= inicio) & (datas <= fim)]

    situacoes = st.sidebar.multiselect("Situação", SITUACOES)
    if situacoes:
        filtrado = filtrado[filtrado["SITUAÇÃO"].isin(situacoes)]

    servidores = st.sidebar.multiselect("Servidor(a)", sorted(df["SERVIDOR(A)"].dropna().astype(str).unique(), key=str.casefold))
    if servidores:
        filtrado = filtrado[filtrado["SERVIDOR(A)"].isin(servidores)]

    fontes = st.sidebar.multiselect("Fonte", FONTES)
    if fontes:
        filtrado = filtrado[filtrado["FONTE"].isin(fontes)]

    assuntos = st.sidebar.multiselect("Assunto", ASSUNTOS)
    if assuntos:
        filtrado = filtrado[filtrado["ASSUNTO"].isin(assuntos)]

    zonas = st.sidebar.multiselect("Zona Eleitoral", ZONAS_ELEITORAIS_BA)
    if zonas:
        filtrado = filtrado[filtrado["ZE"].isin(zonas)]

    busca = st.sidebar.text_input("Busca livre")
    if busca:
        termo = busca.lower()
        filtrado = filtrado[
            filtrado.astype(str).apply(
                lambda linha: linha.str.lower().str.contains(termo, na=False).any(),
                axis=1
            )
        ]

    return filtrado


def card(label, valor, ajuda=None):
    st.metric(label, valor, help=ajuda)


# -----------------------------
# Estado inicial do app
# -----------------------------

if "base" not in st.session_state:
    st.session_state.base = criar_base_vazia()

if "servidores" not in st.session_state:
    st.session_state.servidores = criar_servidores_iniciais()


# -----------------------------
# Layout
# -----------------------------

st.title("📧 Controle de Atendimentos - SEPRO")
st.caption("Sistema para registrar, acompanhar, filtrar e extrair relatórios dos atendimentos de e-mail e telefone.")

st.sidebar.markdown("# Menu")
pagina = st.sidebar.radio(
    "Escolha uma área",
    [
        "Painel",
        "Novo atendimento",
        "Base de atendimentos",
        "Servidores",
        "Relatórios",
        "Importar / Exportar",
    ],
)

# -----------------------------
# Importação rápida pela lateral
# -----------------------------

st.sidebar.markdown("---")
st.sidebar.markdown("## 📁 Base")
arquivo = st.sidebar.file_uploader("Importar planilha Excel", type=["xlsx"], key="upload_global")
if arquivo is not None:
    try:
        df_importado = pd.read_excel(arquivo)
        df_importado = limpar_colunas(df_importado)
        if st.sidebar.button("Usar esta planilha como base"):
            st.session_state.base = df_importado
            servidores_planilha = sorted(
                [s for s in df_importado["SERVIDOR(A)"].dropna().astype(str).str.strip().unique() if s],
                key=str.casefold
            )
            servidores_existentes = set(st.session_state.servidores["SERVIDOR(A)"].astype(str))
            novos = [s for s in servidores_planilha if s not in servidores_existentes]
            if novos:
                st.session_state.servidores = pd.concat([
                    st.session_state.servidores,
                    pd.DataFrame({"SERVIDOR(A)": novos, "ATIVO": [True] * len(novos)})
                ], ignore_index=True)
            st.sidebar.success("Base importada com sucesso.")
            st.rerun()
    except Exception as e:
        st.sidebar.error("Não foi possível importar a planilha.")
        st.sidebar.write(e)


# -----------------------------
# PÁGINA: PAINEL
# -----------------------------

if pagina == "Painel":
    df = st.session_state.base.copy()
    df_filtrado = aplicar_filtros(df)

    total = len(df_filtrado)
    pendentes = df_filtrado["SITUAÇÃO"].astype(str).str.contains("pend|andamento", case=False, na=False).sum() if not df_filtrado.empty else 0
    respondidos = df_filtrado["SITUAÇÃO"].astype(str).str.contains("respond|concl|final", case=False, na=False).sum() if not df_filtrado.empty else 0
    outros = df_filtrado["SITUAÇÃO"].astype(str).str.contains("outro|arquiv", case=False, na=False).sum() if not df_filtrado.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card("Total filtrado", total)
    with c2:
        card("Respondidos", int(respondidos))
    with c3:
        card("Pendentes / em atendimento", int(pendentes))
    with c4:
        card("Outros / arquivados", int(outros))

    st.divider()

    if df_filtrado.empty:
        st.info("Nenhum atendimento encontrado. Importe uma planilha ou registre um novo atendimento.")
    else:
        col_a, col_b = st.columns(2)

        with col_a:
            st.subheader("Atendimentos por situação")
            tabela = df_filtrado["SITUAÇÃO"].fillna("Não informado").value_counts()
            st.bar_chart(tabela)

        with col_b:
            st.subheader("Atendimentos por fonte")
            tabela = df_filtrado["FONTE"].fillna("Não informado").value_counts()
            st.bar_chart(tabela)

        col_c, col_d = st.columns(2)

        with col_c:
            st.subheader("Top servidores")
            top_servidores = df_filtrado["SERVIDOR(A)"].fillna("Não informado").value_counts().head(15)
            tabela_servidores = top_servidores.rename_axis("Servidor(a)").reset_index(name="Quantidade")
            st.dataframe(tabela_servidores, use_container_width=True, hide_index=True)

        with col_d:
            st.subheader("Top assuntos")
            top_assuntos = df_filtrado["ASSUNTO"].fillna("Não informado").value_counts().head(15)
            tabela_assuntos = top_assuntos.rename_axis("Assunto").reset_index(name="Quantidade")
            st.dataframe(tabela_assuntos, use_container_width=True, hide_index=True)

        st.subheader("Base filtrada")
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)


# -----------------------------
# PÁGINA: NOVO ATENDIMENTO
# -----------------------------

elif pagina == "Novo atendimento":
    st.subheader("➕ Registrar novo atendimento")

    servidores_ativos = obter_servidores_ativos()
    if not servidores_ativos:
        st.warning("Cadastre pelo menos um servidor ativo antes de registrar atendimentos.")

    with st.form("form_novo_atendimento", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            data_atendimento = st.date_input("Data", value=date.today())
            situacao = st.selectbox("Situação", SITUACOES, index=0)
            fonte = st.selectbox("Fonte", FONTES)

        with col2:
            ze = st.selectbox("Zona Eleitoral - Bahia", ZONAS_ELEITORAIS_BA)
            servidor = st.selectbox("Servidor(a) responsável", servidores_ativos if servidores_ativos else [""])

        with col3:
            assunto = st.selectbox("Assunto", ASSUNTOS)
            remetente = st.text_input("Quem originou a chamada / remetente", placeholder="Digite livremente")

        observacoes = st.text_area("Observações", placeholder="Registre detalhes do atendimento, protocolo, encaminhamento, retorno etc.")

        salvar = st.form_submit_button("Salvar atendimento")

    if salvar:
        novo = pd.DataFrame([{ 
            "SITUAÇÃO": situacao,
            "ZE": ze,
            "REMETENTE": remetente.strip(),
            "DATA": data_atendimento,
            "SERVIDOR(A)": servidor,
            "FONTE": fonte,
            "ASSUNTO": assunto,
            "OBSERVAÇÕES": observacoes.strip(),
        }])
        st.session_state.base = pd.concat([st.session_state.base, novo], ignore_index=True)
        st.success("Atendimento registrado com sucesso.")


# -----------------------------
# PÁGINA: BASE DE ATENDIMENTOS
# -----------------------------

elif pagina == "Base de atendimentos":
    st.subheader("📋 Base de atendimentos")

    df = st.session_state.base.copy()

    if df.empty:
        st.info("A base está vazia. Importe uma planilha ou registre um novo atendimento.")
    else:
        st.warning("Edite com cuidado. Após alterar, clique em 'Salvar alterações da tabela'.")
        editado = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "SITUAÇÃO": st.column_config.SelectboxColumn("Situação", options=SITUACOES),
                "ZE": st.column_config.SelectboxColumn("ZE", options=ZONAS_ELEITORAIS_BA),
                "SERVIDOR(A)": st.column_config.SelectboxColumn("Servidor(a)", options=obter_servidores_ativos()),
                "FONTE": st.column_config.SelectboxColumn("Fonte", options=FONTES),
                "ASSUNTO": st.column_config.SelectboxColumn("Assunto", options=ASSUNTOS),
                "DATA": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
            },
            key="editor_base"
        )

        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("Salvar alterações da tabela", type="primary"):
                st.session_state.base = limpar_colunas(editado)
                st.success("Alterações salvas.")
                st.rerun()

        with col2:
            st.caption("Para excluir uma linha, use o editor da tabela e depois salve as alterações.")


# -----------------------------
# PÁGINA: SERVIDORES
# -----------------------------

elif pagina == "Servidores":
    st.subheader("👥 Cadastro e descadastro de servidores")

    with st.form("form_servidor"):
        novo_servidor = st.text_input("Nome do servidor(a)")
        incluir = st.form_submit_button("Cadastrar servidor")

    if incluir:
        nome = novo_servidor.strip()
        if not nome:
            st.error("Informe o nome do servidor.")
        else:
            servidores = st.session_state.servidores.copy()
            existentes = servidores["SERVIDOR(A)"].astype(str).str.casefold().tolist()
            if nome.casefold() in existentes:
                st.warning("Esse servidor já existe.")
            else:
                st.session_state.servidores = pd.concat([
                    servidores,
                    pd.DataFrame([{"SERVIDOR(A)": nome, "ATIVO": True}])
                ], ignore_index=True)
                st.success("Servidor cadastrado.")
                st.rerun()

    st.markdown("### Lista de servidores")
    st.caption("Para descadastrar, desmarque a coluna ATIVO. O servidor sai das novas escolhas, mas permanece nos registros antigos.")

    servidores_editados = st.data_editor(
        st.session_state.servidores,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "SERVIDOR(A)": st.column_config.TextColumn("Servidor(a)", required=True),
            "ATIVO": st.column_config.CheckboxColumn("Ativo")
        },
        key="editor_servidores"
    )

    if st.button("Salvar cadastro de servidores", type="primary"):
        servidores_editados["SERVIDOR(A)"] = servidores_editados["SERVIDOR(A)"].astype(str).str.strip()
        servidores_editados = servidores_editados[servidores_editados["SERVIDOR(A)"] != ""]
        servidores_editados = servidores_editados.drop_duplicates(subset=["SERVIDOR(A)"], keep="last")
        st.session_state.servidores = servidores_editados.reset_index(drop=True)
        st.success("Cadastro de servidores atualizado.")
        st.rerun()


# -----------------------------
# PÁGINA: RELATÓRIOS
# -----------------------------

elif pagina == "Relatórios":
    st.subheader("📑 Relatórios")

    df = st.session_state.base.copy()
    df_filtrado = aplicar_filtros(df)

    if df_filtrado.empty:
        st.info("Não há dados para relatório com os filtros atuais.")
    else:
        opcao = st.selectbox(
            "Tipo de relatório",
            [
                "Base filtrada",
                "Resumo por situação",
                "Resumo por servidor",
                "Resumo por assunto",
                "Resumo por zona eleitoral",
                "Resumo por fonte",
                "Resumo mensal",
                "Pendentes / em atendimento",
            ]
        )

        if opcao == "Base filtrada":
            rel = df_filtrado
        elif opcao == "Resumo por situação":
            rel = df_filtrado["SITUAÇÃO"].fillna("Não informado").value_counts().reset_index()
            rel.columns = ["Situação", "Quantidade"]
        elif opcao == "Resumo por servidor":
            rel = df_filtrado["SERVIDOR(A)"].fillna("Não informado").value_counts().reset_index()
            rel.columns = ["Servidor(a)", "Quantidade"]
        elif opcao == "Resumo por assunto":
            rel = df_filtrado["ASSUNTO"].fillna("Não informado").value_counts().reset_index()
            rel.columns = ["Assunto", "Quantidade"]
        elif opcao == "Resumo por zona eleitoral":
            rel = df_filtrado["ZE"].fillna("Não informado").value_counts().reset_index()
            rel.columns = ["Zona Eleitoral", "Quantidade"]
        elif opcao == "Resumo por fonte":
            rel = df_filtrado["FONTE"].fillna("Não informado").value_counts().reset_index()
            rel.columns = ["Fonte", "Quantidade"]
        elif opcao == "Resumo mensal":
            temp = df_filtrado.copy()
            temp["DATA"] = pd.to_datetime(temp["DATA"], errors="coerce")
            rel = (
                temp.dropna(subset=["DATA"])
                .assign(MÊS=lambda x: x["DATA"].dt.to_period("M").astype(str))
                .groupby("MÊS", as_index=False)
                .size()
                .rename(columns={"size": "Quantidade"})
            )
        else:
            rel = df_filtrado[df_filtrado["SITUAÇÃO"].astype(str).str.contains("pend|andamento", case=False, na=False)]

        st.dataframe(rel, use_container_width=True, hide_index=True)

        csv = rel.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            "Baixar este relatório em CSV",
            data=csv,
            file_name="relatorio_controle_atendimentos.csv",
            mime="text/csv"
        )

        excel_completo = gerar_excel_relatorio(df_filtrado, st.session_state.servidores)
        st.download_button(
            "Baixar relatório completo em Excel",
            data=excel_completo,
            file_name="relatorio_completo_controle_atendimentos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# -----------------------------
# PÁGINA: IMPORTAR / EXPORTAR
# -----------------------------

elif pagina == "Importar / Exportar":
    st.subheader("📦 Importar / Exportar dados")

    st.markdown("### Exportar backup completo")
    st.write("Use este arquivo para guardar a base atual, os relatórios e o cadastro de servidores.")

    excel_completo = gerar_excel_relatorio(st.session_state.base, st.session_state.servidores)
    st.download_button(
        "Baixar backup completo em Excel",
        data=excel_completo,
        file_name="backup_controle_atendimentos_sepro.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    st.markdown("### Importar planilha")
    st.write("Use a opção da barra lateral para carregar uma planilha Excel antiga ou um backup exportado pelo sistema.")

    st.info(
        "Importante: no Streamlit Community Cloud, os dados podem não ficar gravados para sempre no servidor. "
        "Por isso, baixe o backup Excel ao final do expediente ou sempre que fizer alterações importantes."
    )
    
