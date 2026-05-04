import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
import json
import hashlib
import secrets
import smtplib
from email.message import EmailMessage

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

st.set_page_config(
    page_title="Sistema SEPRO - Controle de Atendimentos",
    page_icon="📧",
    layout="wide"
)

ARQ_USUARIOS = Path("usuarios_controle_emails_v3.json")
ARQ_ATENDIMENTOS = Path("atendimentos_controle_emails_v1.json")
ARQ_ASSUNTOS = Path("assuntos_controle_emails_v1.json")

DOMINIO_INSTITUCIONAL = "@tre-ba.jus.br"

STATUS_CADASTRADO = "Triagem"
STATUS_EM_ATENDIMENTO = "Em atendimento"
STATUS_REALIZADO = "Atendimento realizado"

STATUS_OPCOES = [
    STATUS_CADASTRADO,
    STATUS_EM_ATENDIMENTO,
    STATUS_REALIZADO,
]

FONTES = [
    "E-mail",
    "Telefone",
    "WhatsApp",
    "Outro",
]

ASSUNTOS_PADRAO = [
    "Cumprimento de Sentença",
    "Prestação de Contas Eleitorais",
    "Prestação de Contas Anual",
    "Sisbajud",
    "AJUE",
    "Ação Penal",
    "AIME",
    "Filiação Partidária",
    "Propaganda Eleitoral",
    "Registro de Candidatura",
    "Direitos Políticos",
    "Regularização de Situação Eleitoral",
    "Zona Eleitoral",
    "Outro",
]

ZONAS_BAHIA = ["Não informado"] + [f"{i:03d}ª Zona Eleitoral - Bahia" for i in range(1, 206)]


# ============================================================
# ESTILO VISUAL
# ============================================================

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #174A7C 0%, #1F5F99 55%, #2C8C6A 100%);
        padding: 18px 22px;
        border-radius: 10px;
        color: white;
        margin-bottom: 18px;
    }
    .main-header h1 {
        margin: 0;
        font-size: 28px;
    }
    .main-header p {
        margin: 4px 0 0 0;
        opacity: 0.92;
    }
    .card {
        border: 1px solid #D9E2EF;
        border-radius: 10px;
        padding: 16px;
        background: #FFFFFF;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
        margin-bottom: 12px;
    }
    .card-title {
        font-weight: 700;
        color: #174A7C;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .status-triagem {
        background: #E8F1FC;
        color: #174A7C;
        padding: 4px 10px;
        border-radius: 999px;
        font-weight: 700;
        display: inline-block;
    }
    .status-em {
        background: #FFF3D9;
        color: #8A5A00;
        padding: 4px 10px;
        border-radius: 999px;
        font-weight: 700;
        display: inline-block;
    }
    .status-realizado {
        background: #E7F6EC;
        color: #176B3A;
        padding: 4px 10px;
        border-radius: 999px;
        font-weight: 700;
        display: inline-block;
    }
    .small-muted {
        color: #64748B;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def agora_iso():
    return datetime.now().isoformat(timespec="seconds")


def hoje_ddmmaaaa():
    return datetime.now().strftime("%d/%m/%Y")


def parse_data(valor):
    if valor is None or valor == "":
        return None

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except Exception:
            pass

    try:
        return pd.to_datetime(texto, dayfirst=True, errors="coerce").date()
    except Exception:
        return None


def data_para_exibir(valor):
    d = parse_data(valor)
    if d:
        return d.strftime("%d/%m/%Y")
    return ""


def iso_para_exibir(valor):
    if not valor:
        return ""
    try:
        return datetime.fromisoformat(str(valor)).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(valor)


def senha_hash(senha):
    return hashlib.sha256(str(senha).encode("utf-8")).hexdigest()


def carregar_json(caminho, padrao):
    if not caminho.exists():
        return padrao
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except Exception:
        return padrao


def salvar_json(caminho, dados):
    caminho.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def normalizar_email(email):
    return str(email or "").strip().lower()


def usuarios():
    return carregar_json(ARQ_USUARIOS, [])


def salvar_usuarios(lista):
    salvar_json(ARQ_USUARIOS, lista)


def atendimentos():
    return carregar_json(ARQ_ATENDIMENTOS, [])


def salvar_atendimentos(lista):
    salvar_json(ARQ_ATENDIMENTOS, lista)


def assuntos():
    lista = carregar_json(ARQ_ASSUNTOS, [])
    if not lista:
        lista = ASSUNTOS_PADRAO
        salvar_json(ARQ_ASSUNTOS, lista)
    return sorted(set([str(x).strip() for x in lista if str(x).strip()]))


def salvar_assuntos(lista):
    salvar_json(ARQ_ASSUNTOS, sorted(set([str(x).strip() for x in lista if str(x).strip()])))


def email_institucional(email):
    return normalizar_email(email).endswith(DOMINIO_INSTITUCIONAL)


def usuario_logado():
    return st.session_state.get("usuario")


def eh_admin():
    u = usuario_logado()
    return bool(u and u.get("perfil") == "Administrador")


def usuarios_ativos_validados():
    return [
        u for u in usuarios()
        if u.get("ativo", True) and u.get("validado", False)
    ]


def nomes_usuarios_ativos():
    nomes = []
    for u in usuarios_ativos_validados():
        nome = u.get("nome") or u.get("email")
        if nome:
            nomes.append(nome)
    return sorted(set(nomes))


def proximo_id(lista):
    if not lista:
        return 1
    ids = []
    for item in lista:
        try:
            ids.append(int(item.get("id", 0)))
        except Exception:
            pass
    return max(ids or [0]) + 1


def status_badge(status):
    if status == STATUS_CADASTRADO:
        return '<span class="status-triagem">Triagem</span>'
    if status == STATUS_EM_ATENDIMENTO:
        return '<span class="status-em">Em atendimento</span>'
    if status == STATUS_REALIZADO:
        return '<span class="status-realizado">Atendimento realizado</span>'
    return f"<b>{status}</b>"


def atendimentos_df(lista=None):
    dados = lista if lista is not None else atendimentos()
    if not dados:
        return pd.DataFrame(columns=[
            "id", "data", "status", "servidor", "fonte", "assunto",
            "zona_eleitoral", "origem", "descricao", "observacoes",
            "criado_por", "criado_em", "atualizado_em", "data_realizacao"
        ])

    df = pd.DataFrame(dados)

    for col in [
        "id", "data", "status", "servidor", "fonte", "assunto",
        "zona_eleitoral", "origem", "descricao", "observacoes",
        "criado_por", "criado_em", "atualizado_em", "data_realizacao"
    ]:
        if col not in df.columns:
            df[col] = ""

    df["Data"] = df["data"].apply(data_para_exibir)
    df["Atualizado em"] = df["atualizado_em"].apply(iso_para_exibir)
    df["Criado em"] = df["criado_em"].apply(iso_para_exibir)
    df["Data de realização"] = df["data_realizacao"].apply(data_para_exibir)

    df = df.rename(columns={
        "id": "ID",
        "status": "Status",
        "servidor": "Servidor(a)",
        "fonte": "Fonte",
        "assunto": "Assunto",
        "zona_eleitoral": "Zona eleitoral",
        "origem": "Origem da demanda",
        "descricao": "Descrição",
        "observacoes": "Observações",
        "criado_por": "Criado por",
    })

    colunas = [
        "ID", "Data", "Status", "Servidor(a)", "Fonte", "Assunto",
        "Zona eleitoral", "Origem da demanda", "Descrição", "Observações",
        "Criado por", "Criado em", "Atualizado em", "Data de realização"
    ]

    return df[colunas]


def enviar_email(destinatario, assunto_email, corpo):
    try:
        smtp_host = st.secrets.get("SMTP_HOST", "")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
        smtp_user = st.secrets.get("SMTP_USER", "")
        smtp_password = st.secrets.get("SMTP_PASSWORD", "")
        remetente = st.secrets.get("EMAIL_REMETENTE", smtp_user)

        if not all([smtp_host, smtp_port, smtp_user, smtp_password, remetente]):
            return False, "Envio de e-mail não configurado nos Secrets do Streamlit."

        msg = EmailMessage()
        msg["From"] = remetente
        msg["To"] = destinatario
        msg["Subject"] = assunto_email
        msg.set_content(corpo)

        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as servidor:
            servidor.starttls()
            servidor.login(smtp_user, smtp_password)
            servidor.send_message(msg)

        return True, "E-mail enviado com sucesso."
    except Exception as e:
        return False, f"Não foi possível enviar o e-mail: {e}"


def app_base_url():
    return st.secrets.get("APP_BASE_URL", "").rstrip("/")


def gerar_link_validacao(token):
    base = app_base_url()
    if base:
        return f"{base}/?validar={token}"
    return f"?validar={token}"


def gerar_link_recuperacao(token):
    base = app_base_url()
    if base:
        return f"{base}/?recuperar={token}"
    return f"?recuperar={token}"


# ============================================================
# PROCESSAMENTO DE TOKENS VIA QUERY PARAMS
# ============================================================

def get_query_param(nome):
    try:
        valor = st.query_params.get(nome)
        if isinstance(valor, list):
            return valor[0] if valor else None
        return valor
    except Exception:
        return None


def processar_validacao():
    token = get_query_param("validar")
    if not token:
        return

    lista = usuarios()
    alterou = False

    for u in lista:
        if u.get("token_validacao") == token:
            u["validado"] = True
            u["token_validacao"] = ""
            u["ativo"] = True
            alterou = True
            break

    if alterou:
        salvar_usuarios(lista)
        st.success("Cadastro validado com sucesso. Faça login para acessar o sistema.")
    else:
        st.error("Link de validação inválido ou expirado.")


def processar_recuperacao():
    token = get_query_param("recuperar")
    if not token:
        return

    lista = usuarios()
    usuario = None

    for u in lista:
        if u.get("token_recuperacao") == token:
            usuario = u
            break

    if not usuario:
        st.error("Link de recuperação inválido ou expirado.")
        return

    st.subheader("Redefinir senha")
    nova = st.text_input("Nova senha", type="password")
    confirmar = st.text_input("Confirmar nova senha", type="password")

    if st.button("Salvar nova senha"):
        if not nova or len(nova) < 6:
            st.warning("A senha deve ter pelo menos 6 caracteres.")
        elif nova != confirmar:
            st.warning("As senhas não conferem.")
        else:
            usuario["senha_hash"] = senha_hash(nova)
            usuario["token_recuperacao"] = ""
            salvar_usuarios(lista)
            st.success("Senha redefinida com sucesso. Faça login novamente.")


# ============================================================
# AUTENTICAÇÃO
# ============================================================

def tela_login():
    st.markdown(
        """
        <div class="main-header">
            <h1>📧 Sistema SEPRO - Controle de Atendimentos</h1>
            <p>Controle de demandas, movimentação por status e relatórios.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    processar_validacao()
    processar_recuperacao()

    aba_login, aba_cadastro, aba_recuperar = st.tabs(["Entrar", "Cadastrar usuário", "Recuperar senha"])

    with aba_login:
        st.subheader("Acesso ao sistema")
        email = normalizar_email(st.text_input("E-mail"))
        senha = st.text_input("Senha", type="password")

        if st.button("Entrar", type="primary"):
            lista = usuarios()
            encontrado = None

            for u in lista:
                if normalizar_email(u.get("email")) == email:
                    encontrado = u
                    break

            if not encontrado:
                st.error("Usuário não encontrado.")
            elif not encontrado.get("ativo", True):
                st.error("Usuário inativo. Procure o administrador.")
            elif not encontrado.get("validado", False):
                st.error("Usuário ainda não validado por e-mail.")
            elif encontrado.get("senha_hash") != senha_hash(senha):
                st.error("Senha inválida.")
            else:
                st.session_state["usuario"] = encontrado
                st.success("Login realizado com sucesso.")
                st.rerun()

    with aba_cadastro:
        st.subheader("Cadastro de usuário")

        lista = usuarios()
        primeiro_usuario = len(lista) == 0

        nome = st.text_input("Nome completo", key="cad_nome")
        email = normalizar_email(st.text_input("E-mail institucional", key="cad_email"))
        senha = st.text_input("Senha", type="password", key="cad_senha")
        confirmar = st.text_input("Confirmar senha", type="password", key="cad_conf")

        if primeiro_usuario:
            st.info("O primeiro usuário triagem será o Administrador do sistema.")
        else:
            st.info("Após o cadastro, será enviado um link de validação para o e-mail institucional.")

        if st.button("Cadastrar", key="btn_cadastrar"):
            if not nome.strip():
                st.warning("Informe o nome.")
            elif not email_institucional(email):
                st.warning(f"O e-mail deve terminar com {DOMINIO_INSTITUCIONAL}.")
            elif len(senha) < 6:
                st.warning("A senha deve ter pelo menos 6 caracteres.")
            elif senha != confirmar:
                st.warning("As senhas não conferem.")
            elif any(normalizar_email(u.get("email")) == email for u in lista):
                st.warning("Este e-mail já está triagem.")
            else:
                token = secrets.token_urlsafe(32)

                novo = {
                    "id": proximo_id(lista),
                    "nome": nome.strip(),
                    "email": email,
                    "senha_hash": senha_hash(senha),
                    "perfil": "Administrador" if primeiro_usuario else "Usuário",
                    "ativo": True,
                    "validado": True if primeiro_usuario else False,
                    "token_validacao": "" if primeiro_usuario else token,
                    "token_recuperacao": "",
                    "criado_em": agora_iso(),
                }

                lista.append(novo)
                salvar_usuarios(lista)

                if primeiro_usuario:
                    st.success("Administrador triagem com sucesso. Faça login.")
                else:
                    link = gerar_link_validacao(token)
                    ok, msg = enviar_email(
                        email,
                        "Validação de cadastro - Sistema SEPRO",
                        f"Olá, {nome}.\n\nPara validar seu cadastro no Sistema SEPRO, acesse o link abaixo:\n\n{link}\n\nCaso não tenha solicitado o cadastro, ignore esta mensagem."
                    )
                    if ok:
                        st.success("Usuário triagem. Link de validação enviado ao e-mail informado.")
                    else:
                        st.warning(f"{msg} Link de validação gerado: {link}")

    with aba_recuperar:
        st.subheader("Recuperação de senha")
        email_rec = normalizar_email(st.text_input("E-mail triagem", key="rec_email"))

        if st.button("Gerar link de recuperação"):
            lista = usuarios()
            encontrado = None

            for u in lista:
                if normalizar_email(u.get("email")) == email_rec:
                    encontrado = u
                    break

            if not encontrado:
                st.error("E-mail não encontrado.")
            else:
                token = secrets.token_urlsafe(32)
                encontrado["token_recuperacao"] = token
                salvar_usuarios(lista)

                link = gerar_link_recuperacao(token)
                ok, msg = enviar_email(
                    email_rec,
                    "Recuperação de senha - Sistema SEPRO",
                    f"Olá.\n\nPara redefinir sua senha, acesse o link abaixo:\n\n{link}\n\nCaso não tenha solicitado, ignore esta mensagem."
                )

                if ok:
                    st.success("Link de recuperação enviado ao e-mail triagem.")
                else:
                    st.warning(f"{msg} Link de recuperação gerado: {link}")


# ============================================================
# COMPONENTES DO SISTEMA
# ============================================================

def cabecalho():
    u = usuario_logado()
    st.markdown(
        f"""
        <div class="main-header">
            <h1>📧 Sistema SEPRO - Controle de Atendimentos</h1>
            <p>Usuário: {u.get("nome", "")} | Perfil: {u.get("perfil", "")}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


def sidebar_menu():
    st.sidebar.title("Menu")

    opcoes = [
        "Dashboard",
        "Novo atendimento",
        "Triagem",
        "Em atendimento",
        "Atendimento realizado",
        "Base geral",
        "Relatórios e exportação",
    ]

    if eh_admin():
        opcoes += ["Assuntos", "Usuários"]

    escolha = st.sidebar.radio("Navegação", opcoes)

    st.sidebar.divider()
    if st.sidebar.button("Sair"):
        st.session_state.pop("usuario", None)
        st.rerun()

    return escolha


def filtros_base(lista):
    df = atendimentos_df(lista)

    st.sidebar.subheader("Filtros")

    busca = st.sidebar.text_input("Busca livre")
    status = st.sidebar.multiselect("Status", STATUS_OPCOES)
    fontes = st.sidebar.multiselect("Fonte", FONTES)
    lista_assuntos = assuntos()
    filtro_assuntos = st.sidebar.multiselect("Assunto", lista_assuntos)

    servidores = sorted(df["Servidor(a)"].dropna().astype(str).unique()) if not df.empty else []
    filtro_servidores = st.sidebar.multiselect("Servidor(a)", servidores)

    if df.empty:
        return []

    ids_validos = set(df["ID"].tolist())

    if busca:
        termo = busca.lower()
        df = df[
            df.astype(str).apply(
                lambda linha: linha.str.lower().str.contains(termo, na=False).any(),
                axis=1
            )
        ]

    if status:
        df = df[df["Status"].isin(status)]

    if fontes:
        df = df[df["Fonte"].isin(fontes)]

    if filtro_assuntos:
        df = df[df["Assunto"].isin(filtro_assuntos)]

    if filtro_servidores:
        df = df[df["Servidor(a)"].isin(filtro_servidores)]

    ids_filtrados = set(df["ID"].tolist()) & ids_validos

    return [a for a in lista if a.get("id") in ids_filtrados]


def card_atendimento(atendimento, chave_prefixo, permitir_edicao=True):
    status = atendimento.get("status", STATUS_CADASTRADO)

    with st.container(border=True):
        col1, col2, col3 = st.columns([1.2, 2.2, 1.3])

        with col1:
            st.markdown(f"**ID:** {atendimento.get('id')}")
            st.markdown(f"**Data:** {data_para_exibir(atendimento.get('data'))}")
            st.markdown(status_badge(status), unsafe_allow_html=True)

        with col2:
            st.markdown(f"**Assunto:** {atendimento.get('assunto', '')}")
            st.markdown(f"**Origem:** {atendimento.get('origem', '')}")
            st.markdown(f"**Zona:** {atendimento.get('zona_eleitoral', '')}")
            st.markdown(f"**Descrição:** {atendimento.get('descricao', '')}")

        with col3:
            st.markdown(f"**Servidor(a):** {atendimento.get('servidor', '')}")
            st.markdown(f"**Fonte:** {atendimento.get('fonte', '')}")
            st.markdown(f"**Atualizado:** {iso_para_exibir(atendimento.get('atualizado_em'))}")

        if atendimento.get("observacoes"):
            st.markdown(f"**Observações:** {atendimento.get('observacoes')}")

        if not permitir_edicao:
            return

        st.divider()

        col_a, col_b, col_c, col_d = st.columns([1.1, 1.1, 1.1, 1.2])

        lista = atendimentos()

        def salvar_status(novo_status):
            for item in lista:
                if item.get("id") == atendimento.get("id"):
                    item["status"] = novo_status
                    item["atualizado_em"] = agora_iso()
                    if novo_status == STATUS_REALIZADO:
                        item["data_realizacao"] = hoje_ddmmaaaa()
                    salvar_atendimentos(lista)
                    st.success(f"Atendimento movido para: {novo_status}")
                    st.rerun()

        with col_a:
            if status != STATUS_CADASTRADO:
                if st.button("Mover para Triagem", key=f"{chave_prefixo}_cad_{atendimento.get('id')}"):
                    salvar_status(STATUS_CADASTRADO)

        with col_b:
            if status != STATUS_EM_ATENDIMENTO:
                if st.button("Mover para Em atendimento", key=f"{chave_prefixo}_em_{atendimento.get('id')}"):
                    salvar_status(STATUS_EM_ATENDIMENTO)

        with col_c:
            if status != STATUS_REALIZADO:
                if st.button("Mover para Realizado", key=f"{chave_prefixo}_real_{atendimento.get('id')}"):
                    salvar_status(STATUS_REALIZADO)

        with col_d:
            with st.popover("Editar observação/status"):
                nova_obs = st.text_area(
                    "Observações",
                    value=atendimento.get("observacoes", ""),
                    key=f"{chave_prefixo}_obs_{atendimento.get('id')}"
                )
                novo_status = st.selectbox(
                    "Status",
                    STATUS_OPCOES,
                    index=STATUS_OPCOES.index(status) if status in STATUS_OPCOES else 0,
                    key=f"{chave_prefixo}_status_{atendimento.get('id')}"
                )
                novo_servidor = st.selectbox(
                    "Servidor(a)",
                    nomes_usuarios_ativos() or [atendimento.get("servidor", "") or "Não definido"],
                    index=0,
                    key=f"{chave_prefixo}_serv_{atendimento.get('id')}"
                )

                if st.button("Salvar alterações", key=f"{chave_prefixo}_salvar_{atendimento.get('id')}"):
                    for item in lista:
                        if item.get("id") == atendimento.get("id"):
                            item["observacoes"] = nova_obs
                            item["status"] = novo_status
                            item["servidor"] = novo_servidor
                            item["atualizado_em"] = agora_iso()
                            if novo_status == STATUS_REALIZADO and not item.get("data_realizacao"):
                                item["data_realizacao"] = hoje_ddmmaaaa()
                            salvar_atendimentos(lista)
                            st.success("Alterações salvas.")
                            st.rerun()


def tela_dashboard():
    st.subheader("Dashboard")

    lista = filtros_base(atendimentos())
    df = atendimentos_df(lista)

    total = len(df)
    qtd_triagem = int((df["Status"] == STATUS_CADASTRADO).sum()) if not df.empty else 0
    qtd_em = int((df["Status"] == STATUS_EM_ATENDIMENTO).sum()) if not df.empty else 0
    qtd_realizado = int((df["Status"] == STATUS_REALIZADO).sum()) if not df.empty else 0
    percentual = (qtd_realizado / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de atendimentos", total)
    c2.metric("Triagem", qtd_triagem)
    c3.metric("Em atendimento", qtd_em)
    c4.metric("% realizados", f"{percentual:.2f}%".replace(".", ","))

    st.divider()

    c1, c2, c3 = st.columns(3)
    c1.metric("Atendimentos realizados", qtd_realizado)
    c2.metric("Pendentes totais", qtd_triagem + qtd_em)
    c3.metric("Fontes utilizadas", df["Fonte"].nunique() if not df.empty else 0)

    st.divider()

    if df.empty:
        st.info("Nenhum registro encontrado com os filtros atuais.")
        return

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Situação")
        st.dataframe(
            df["Status"].value_counts().rename_axis("Status").reset_index(name="Quantidade"),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("#### Top servidores")
        st.dataframe(
            df["Servidor(a)"].replace("", "Não informado").value_counts().head(10).rename_axis("Servidor(a)").reset_index(name="Quantidade"),
            use_container_width=True,
            hide_index=True
        )

    with col2:
        st.markdown("#### Fonte do atendimento")
        st.dataframe(
            df["Fonte"].replace("", "Não informado").value_counts().rename_axis("Fonte").reset_index(name="Quantidade"),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("#### Top assuntos")
        st.dataframe(
            df["Assunto"].replace("", "Não informado").value_counts().head(10).rename_axis("Assunto").reset_index(name="Quantidade"),
            use_container_width=True,
            hide_index=True
        )


def tela_novo_atendimento():
    st.subheader("Novo atendimento")

    st.info("Todo novo atendimento entra primeiro na base **Triagem**. Depois ele pode ser movido para **Em atendimento** e, ao final, para **Atendimento realizado**.")

    servidores = nomes_usuarios_ativos()
    if not servidores:
        servidores = [usuario_logado().get("nome", "Não definido")]

    with st.form("form_novo_atendimento", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            data_atendimento = st.date_input("Data do atendimento", value=date.today(), format="DD/MM/YYYY")
            fonte = st.selectbox("Fonte", FONTES)
            servidor = st.selectbox("Servidor(a) responsável", servidores)

        with col2:
            assunto = st.selectbox("Assunto", assuntos())
            zona = st.selectbox("Zona eleitoral", ZONAS_BAHIA)
            origem = st.text_input("Quem originou a demanda/chamada")

        with col3:
            protocolo = st.text_input("Protocolo ou referência, se houver")
            prioridade = st.selectbox("Prioridade", ["Normal", "Alta", "Urgente", "Baixa"])
            status_inicial = st.selectbox("Status inicial", [STATUS_CADASTRADO, STATUS_EM_ATENDIMENTO])

        descricao = st.text_area("Descrição da demanda")
        observacoes = st.text_area("Observações internas")

        enviar = st.form_submit_button("Cadastrar atendimento", type="primary")

        if enviar:
            lista = atendimentos()
            novo = {
                "id": proximo_id(lista),
                "data": data_atendimento.strftime("%d/%m/%Y"),
                "status": status_inicial,
                "servidor": servidor,
                "fonte": fonte,
                "assunto": assunto,
                "zona_eleitoral": zona,
                "origem": origem.strip(),
                "protocolo": protocolo.strip(),
                "prioridade": prioridade,
                "descricao": descricao.strip(),
                "observacoes": observacoes.strip(),
                "criado_por": usuario_logado().get("email", ""),
                "criado_em": agora_iso(),
                "atualizado_em": agora_iso(),
                "data_realizacao": "",
            }
            lista.append(novo)
            salvar_atendimentos(lista)
            st.success(f"Atendimento nº {novo['id']} triagem com sucesso na base: {status_inicial}.")


def tela_status(nome_status, titulo, texto_ajuda):
    st.subheader(titulo)
    st.caption(texto_ajuda)

    lista = [a for a in filtros_base(atendimentos()) if a.get("status") == nome_status]

    st.metric("Total nesta base", len(lista))

    if not lista:
        st.info("Nenhum atendimento nesta etapa.")
        return

    lista_ordenada = sorted(lista, key=lambda x: int(x.get("id", 0)), reverse=True)

    for item in lista_ordenada:
        card_atendimento(item, f"card_{nome_status.replace(' ', '_')}")


def tela_base_geral():
    st.subheader("Base geral de atendimentos")

    lista = filtros_base(atendimentos())
    df = atendimentos_df(lista)

    st.write(f"Total encontrado: **{len(df)}**")

    if df.empty:
        st.info("Nenhum atendimento encontrado.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)

        with st.expander("Visualizar cartões dos atendimentos"):
            for item in sorted(lista, key=lambda x: int(x.get("id", 0)), reverse=True):
                card_atendimento(item, "base_geral")


def tela_relatorios_exportacao():
    st.subheader("Relatórios e exportação")

    lista = filtros_base(atendimentos())
    df = atendimentos_df(lista)

    if df.empty:
        st.info("Nenhum registro para exportar.")
        return

    st.markdown("#### Base filtrada")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Relatório por status")
        st.dataframe(df["Status"].value_counts().rename_axis("Status").reset_index(name="Quantidade"), use_container_width=True, hide_index=True)

        st.markdown("#### Relatório por servidor")
        st.dataframe(df["Servidor(a)"].value_counts().rename_axis("Servidor(a)").reset_index(name="Quantidade"), use_container_width=True, hide_index=True)

    with col2:
        st.markdown("#### Relatório por assunto")
        st.dataframe(df["Assunto"].value_counts().rename_axis("Assunto").reset_index(name="Quantidade"), use_container_width=True, hide_index=True)

        st.markdown("#### Relatório por fonte")
        st.dataframe(df["Fonte"].value_counts().rename_axis("Fonte").reset_index(name="Quantidade"), use_container_width=True, hide_index=True)

    st.divider()

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Atendimentos")
        df["Status"].value_counts().rename_axis("Status").reset_index(name="Quantidade").to_excel(writer, index=False, sheet_name="Por status")
        df["Servidor(a)"].value_counts().rename_axis("Servidor(a)").reset_index(name="Quantidade").to_excel(writer, index=False, sheet_name="Por servidor")
        df["Assunto"].value_counts().rename_axis("Assunto").reset_index(name="Quantidade").to_excel(writer, index=False, sheet_name="Por assunto")
        df["Fonte"].value_counts().rename_axis("Fonte").reset_index(name="Quantidade").to_excel(writer, index=False, sheet_name="Por fonte")

    st.download_button(
        "Baixar relatório em Excel",
        data=buffer.getvalue(),
        file_name=f"relatorio_atendimentos_sepro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        "Baixar base em CSV",
        data=csv,
        file_name=f"base_atendimentos_sepro_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )


def tela_assuntos():
    st.subheader("Assuntos")
    st.caption("Área exclusiva do administrador.")

    lista = assuntos()

    with st.form("form_assunto"):
        novo = st.text_input("Novo assunto")
        if st.form_submit_button("Adicionar assunto"):
            if not novo.strip():
                st.warning("Informe um assunto.")
            else:
                lista.append(novo.strip())
                salvar_assuntos(lista)
                st.success("Assunto adicionado.")
                st.rerun()

    st.divider()
    st.markdown("#### Assuntos triagems")

    for assunto in lista:
        col1, col2 = st.columns([4, 1])
        col1.write(assunto)
        if col2.button("Excluir", key=f"del_assunto_{assunto}"):
            nova_lista = [a for a in lista if a != assunto]
            salvar_assuntos(nova_lista)
            st.success("Assunto removido da lista de opções. Registros antigos foram preservados.")
            st.rerun()


def tela_usuarios():
    st.subheader("Usuários")
    st.caption("Área exclusiva do administrador.")

    lista = usuarios()
    df = pd.DataFrame(lista)

    if not df.empty:
        exibir = df.copy()
        for col in ["senha_hash", "token_validacao", "token_recuperacao"]:
            if col in exibir.columns:
                exibir = exibir.drop(columns=[col])
        st.dataframe(exibir, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum usuário triagem.")

    st.divider()

    st.markdown("#### Gerenciar usuário")

    opcoes = [f"{u.get('nome')} - {u.get('email')}" for u in lista]
    if not opcoes:
        return

    selecionado = st.selectbox("Selecione o usuário", opcoes)
    email_sel = selecionado.split(" - ")[-1]

    usuario = next((u for u in lista if u.get("email") == email_sel), None)
    if not usuario:
        return

    col1, col2, col3 = st.columns(3)

    with col1:
        novo_perfil = st.selectbox(
            "Perfil",
            ["Administrador", "Usuário"],
            index=0 if usuario.get("perfil") == "Administrador" else 1
        )

    with col2:
        ativo = st.checkbox("Ativo", value=usuario.get("ativo", True))

    with col3:
        validado = st.checkbox("Validado", value=usuario.get("validado", False))

    if st.button("Salvar alterações do usuário"):
        usuario["perfil"] = novo_perfil
        usuario["ativo"] = ativo
        usuario["validado"] = validado
        salvar_usuarios(lista)
        st.success("Usuário atualizado.")
        st.rerun()

    st.divider()
    st.markdown("#### Excluir usuário")

    email_logado = usuario_logado().get("email")

    if usuario.get("email") == email_logado:
        st.warning("Você não pode excluir o próprio usuário enquanto estiver logado.")
    else:
        admins_ativos = [
            u for u in lista
            if u.get("perfil") == "Administrador" and u.get("ativo", True) and u.get("email") != usuario.get("email")
        ]

        if usuario.get("perfil") == "Administrador" and not admins_ativos:
            st.warning("Não é permitido excluir o último administrador ativo.")
        else:
            confirmar = st.checkbox(f"Confirmo a exclusão de {usuario.get('email')}")
            if st.button("Excluir usuário", type="secondary"):
                if not confirmar:
                    st.warning("Marque a confirmação antes de excluir.")
                else:
                    nova_lista = [u for u in lista if u.get("email") != usuario.get("email")]
                    salvar_usuarios(nova_lista)
                    st.success("Usuário excluído.")
                    st.rerun()


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

def main():
    if not usuario_logado():
        tela_login()
        return

    cabecalho()
    escolha = sidebar_menu()

    if escolha == "Dashboard":
        tela_dashboard()

    elif escolha == "Novo atendimento":
        tela_novo_atendimento()

    elif escolha == "Triagem":
        tela_status(
            STATUS_CADASTRADO,
            "Triagem",
            "Atendimentos cadastrados, ainda aguardando triagem ou início do atendimento."
        )

    elif escolha == "Em atendimento":
        tela_status(
            STATUS_EM_ATENDIMENTO,
            "Em atendimento",
            "Atendimentos em análise ou execução pela equipe responsável."
        )

    elif escolha == "Atendimento realizado":
        tela_status(
            STATUS_REALIZADO,
            "Atendimento realizado",
            "Atendimentos concluídos/realizados."
        )

    elif escolha == "Base geral":
        tela_base_geral()

    elif escolha == "Relatórios e exportação":
        tela_relatorios_exportacao()

    elif escolha == "Assuntos" and eh_admin():
        tela_assuntos()

    elif escolha == "Usuários" and eh_admin():
        tela_usuarios()


if __name__ == "__main__":
    main()
