# app.py
import os
import io
import time
import zipfile
import threading
import queue
from datetime import date
from typing import List, Dict, Any

import pandas as pd
import streamlit as st


import auth
import config as cfgmod
import data_store

st.set_page_config(
    page_title="Portal NFS-e",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    r'''
<style>
/* ===== Portal NFS-e - UI skin (corporativo, clean) ===== */
:root{
  --card-bg: rgba(255,255,255,0.06);
  --card-brd: rgba(255,255,255,0.10);
  --muted: rgba(255,255,255,0.68);
  --muted2: rgba(255,255,255,0.56);
  --ok: #22c55e;
  --warn: #f59e0b;
  --bad: #ef4444;
}

/* Page padding */
.block-container { padding-top: 1.0rem; padding-bottom: 2.2rem; }

/* Hide Streamlit default header & footer */
header[data-testid="stHeader"] { visibility: hidden; height: 0px; }
footer { visibility: hidden; height: 0px; }

/* Sidebar polish */
section[data-testid="stSidebar"] > div { padding-top: 1.2rem; }
section[data-testid="stSidebar"] .stRadio label p { font-size: 0.95rem; }
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: var(--muted); }

/* Brand bar */
.portal-topbar{
  display:flex; align-items:center; justify-content:space-between;
  padding: 14px 16px;
  border: 1px solid var(--card-brd);
  background: linear-gradient(90deg, rgba(255,255,255,0.07), rgba(255,255,255,0.03));
  border-radius: 16px;
  margin: 0 0 14px 0;
}
.portal-title{
  font-size: 1.2rem; font-weight: 700;
  letter-spacing: 0.2px;
}
.portal-sub{
  font-size: 0.88rem; color: var(--muted2);
}
.portal-pill{
  font-size: 0.78rem; color: rgba(255,255,255,0.75);
  padding: 6px 10px; border-radius: 999px;
  border: 1px solid var(--card-brd);
  background: rgba(255,255,255,0.05);
}

/* Cards */
.portal-card{
  border: 1px solid var(--card-brd);
  background: var(--card-bg);
  border-radius: 16px;
  padding: 14px 14px;
}
.portal-card h4{ margin: 0 0 6px 0; }
.portal-muted{ color: var(--muted2); }

/* Metrics */
div[data-testid="stMetric"] {
  border: 1px solid var(--card-brd);
  background: var(--card-bg);
  padding: 14px;
  border-radius: 16px;
}
div[data-testid="stMetric"] > div { gap: 4px; }

/* Buttons */
.stButton > button, .stDownloadButton > button{
  border-radius: 14px !important;
  padding: 0.65rem 0.9rem !important;
  border: 1px solid var(--card-brd) !important;
}
.stButton > button:hover, .stDownloadButton > button:hover{
  border-color: rgba(255,255,255,0.22) !important;
}

/* Inputs */
div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div{
  border-radius: 14px !important;
}

/* Tables */
div[data-testid="stDataFrame"]{
  border: 1px solid var(--card-brd);
  border-radius: 16px;
  overflow: hidden;
}
</style>
''',
    unsafe_allow_html=True,
)



def render_topbar(user=None):
    # Cabeçalho padrão (apresentável) — mantém o app com cara de produto
    nome = getattr(user, "name", "Operador") if user else "Operador"
    role = getattr(user, "role", "user") if user else "user"
    st.markdown(
        f'''
<div class="portal-topbar">
  <div>
    <div class="portal-title">🧾 Portal NFS-e</div>
    <div class="portal-sub">Pipeline operacional • extração • auditoria • evidências</div>
  </div>
  <div class="portal-pill">Usuário: <b>{nome}</b> • Perfil: <b>{role}</b></div>
</div>
''',
        unsafe_allow_html=True,
    )

# ----------------------------
# Auth
# ----------------------------
def require_login() -> auth.User:
    if "user" not in st.session_state:
        st.session_state.user = None
    if st.session_state.user:
        return st.session_state.user

    # Topbar (tela de login)
    st.markdown(
        '''
<div class="portal-topbar">
  <div>
    <div class="portal-title">🔐 Acesso</div>
    <div class="portal-sub">Controle de acesso e trilha de auditoria</div>
  </div>
  <div class="portal-pill">Ambiente local</div>
</div>
''',
        unsafe_allow_html=True,
    )

    # Card (form)
    st.markdown('<div class="portal-card">', unsafe_allow_html=True)
    st.markdown("#### Entre com suas credenciais")
    st.caption("Admin padrão é criado na primeira execução. Boas práticas: troque a senha após o primeiro acesso.")

    u = st.text_input("Usuário", placeholder="ex.: admin")
    p = st.text_input("Senha", type="password", placeholder="••••••••")

    if st.button("Entrar", use_container_width=True):
        try:
            user = auth.authenticate(u.strip(), p, users_path="users.json")
        except RuntimeError as e:
            st.error(str(e))
            st.markdown("</div>", unsafe_allow_html=True)
            st.stop()
        except Exception as e:
            st.error(f"Falha no login: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
            st.stop()

        if not user:
            st.error("Usuário ou senha inválidos.")
        else:
            st.session_state.user = user
            st.markdown("</div>", unsafe_allow_html=True)
            st.rerun()

    # fecha card e bloqueia acesso sem login
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


def sidebar(user: auth.User) -> str:
    with st.sidebar:
        st.markdown("### 🧾 Portal NFS-e")
        st.caption("Operação padronizada • evidências • compliance")
        st.markdown(f"**Usuário:** {user.name}")
        st.markdown(f"**Perfil:** {user.role}")
        st.divider()
        page = st.radio(
            "Menu",
            ["📊 Painel", "🧾 Processar NFS-e", "🧩 Clientes", "⚙️ Configurações", "🧾 Auditoria"],
            index=1,
        )
        st.divider()
        if st.button("Sair", use_container_width=True):
            st.session_state.user = None
            st.rerun()
    return page


# ----------------------------
# Bot helpers
# ----------------------------
COLUNAS_LOG_ORDEM = [
    "NUMERO_NF",
    "DATA_EMISSAO",
    "DATA_COMPETENCIA",
    "CNPJ_PRESTADOR",
    "RAZAO_PRESTADOR",
    "CNPJ_TOMADOR",
    "RAZAO_TOMADOR",
    "OPTANTE_SN",
    "CODIGO_TRIBUTACAO_NACIONAL",
    "VALOR_SERVICO",
    "IR",
    "ISS",
    "ISS_RETIDO",
    "CSLL",
    "DEDUCOES",
    "PIS",
    "COFINS",
    "INSS",
    "DESC_INCOND",
    "DESC_COND",
    "OUTRAS_RET",
    "ALIQUOTA",
    "BASE_CALCULO",
    "VALOR_LIQUIDO",
    "SITUACAO",
]


def _patch_bot_paths(cfg: cfgmod.AppConfig):
    """
    Mantém o bot original e só padroniza caminhos.
    """
    import bot_nfse

    # Padrão de produção: usar Selenium Manager (Selenium>=4.6) e evitar binário errado/corrompido do webdriver_manager
    # Isso mitiga o clássico [WinError 193] ao iniciar o ChromeDriver.
    try:
        bot_nfse.USE_WEBDRIVER_MANAGER = False
    except Exception:
        pass

    bot_nfse.CAMINHO_PLANILHA = os.path.abspath(cfg.caminho_planilha)
    bot_nfse.PASTA_BASE_SAIDA = os.path.abspath(cfg.pasta_base_saida)
    bot_nfse.PASTA_DOWNLOAD_TEMP = os.path.abspath(cfg.pasta_download_temp)
    bot_nfse.PASTA_IMAGENS_CERT = os.path.abspath(cfg.pasta_imagens_cert)
    bot_nfse.DELAY_ACAO = float(cfg.delay_acao)

    os.makedirs(bot_nfse.PASTA_DOWNLOAD_TEMP, exist_ok=True)
    os.makedirs(bot_nfse.PASTA_BASE_SAIDA, exist_ok=True)
    os.makedirs(bot_nfse.PASTA_IMAGENS_CERT, exist_ok=True)

    return bot_nfse


def _zip_dir(folder: str) -> bytes:
    """
    Compacta a pasta para download.

    Hardening (Windows):
    - ignora arquivos temporários do Excel (prefixo '~$')
    - ignora arquivos bloqueados (PermissionError) para não quebrar a auditoria
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, folder)
                try:
                    z.write(full, rel)
                except PermissionError:
                    continue
                except FileNotFoundError:
                    continue
    buf.seek(0)
    return buf.read()


def _emit(q: "queue.Queue", evt: Dict[str, Any]) -> None:
    try:
        q.put(evt, block=False)
    except Exception:
        pass


def run_bot_job(cfg: cfgmod.AppConfig, ano: int, mes: int, empresas: List[str], events: "queue.Queue", stop_evt: threading.Event):
    """
    Worker thread: NÃO chama st.* (evita 'missing ScriptRunContext').
    Progresso é comunicado por eventos na fila.
    """
    try:
        bot_nfse = _patch_bot_paths(cfg)
        data_store.garantir_planilha_modelo(cfg.caminho_planilha)

        clientes = bot_nfse.carregar_clientes_da_planilha()
        clientes = [c for c in clientes if c.get("EMPRESA") in empresas]

        if not clientes:
            _emit(events, {"type": "error", "message": "Nenhum cliente selecionado (ou nenhum ATIVO)."})
            return

        bot = bot_nfse.NFSePortalBot(ano, mes)
        _emit(events, {"type": "init", "output_folder": bot.pasta_competencia})

        for c in clientes:
            if stop_evt.is_set():
                _emit(events, {"type": "log", "message": "[INFO] Execução interrompida pelo operador."})
                break

            empresa = c.get("EMPRESA", "")
            _emit(
                events,
                {
                    "type": "client_start",
                    "empresa": empresa,
                    "row": {
                        "EMPRESA": empresa,
                        "CNPJ": c.get("CNPJ", ""),
                        "PREFEITURA": c.get("PREFEITURA", ""),
                        "TIPO_ACESSO": c.get("TIPO_ACESSO", ""),
                        "STATUS": "EM EXECUÇÃO",
                        "DETALHE": "",
                    },
                },
            )

            try:
                bot._processar_cliente(c)
                _emit(events, {"type": "client_end", "empresa": empresa, "status": "OK", "detalhe": ""})
            except Exception as e:
                _emit(events, {"type": "client_end", "empresa": empresa, "status": "FALHA", "detalhe": str(e)})

        if bot.registros_log:
            df_log = pd.DataFrame(bot.registros_log)
            for col in COLUNAS_LOG_ORDEM:
                if col not in df_log.columns:
                    df_log[col] = ""
            df_log = df_log.reindex(columns=COLUNAS_LOG_ORDEM)
            os.makedirs(bot.pasta_competencia, exist_ok=True)
            caminho_log = os.path.join(bot.pasta_competencia, f"LOG_NFSE_{bot.competencia_str}.xlsx")
            df_log.to_excel(caminho_log, index=False)
            _emit(events, {"type": "log", "message": f"[INFO] Log salvo em: {caminho_log}"})

        _emit(events, {"type": "done"})

    except Exception as e:
        _emit(events, {"type": "error", "message": str(e)})


# ----------------------------
# Job state (session)
# ----------------------------
def job_init_if_needed():
    st.session_state.setdefault("job", {"active": False, "status": [], "logs": [], "error": None, "output_folder": None})
    st.session_state.setdefault("job_index", {})  # empresa -> idx
    st.session_state.setdefault("job_events", None)
    st.session_state.setdefault("job_stop_evt", None)
    st.session_state.setdefault("job_thread", None)


def drain_events():
    q = st.session_state.get("job_events")
    if not q:
        return

    while True:
        try:
            evt = q.get_nowait()
        except Exception:
            break

        t = evt.get("type")
        if t == "init":
            st.session_state.job["output_folder"] = evt.get("output_folder")
            st.session_state.job["active"] = True

        elif t == "client_start":
            row = evt.get("row", {})
            empresa = evt.get("empresa", row.get("EMPRESA", ""))
            st.session_state.job["status"].append(row)
            st.session_state.job_index[empresa] = len(st.session_state.job["status"]) - 1

        elif t == "client_end":
            empresa = evt.get("empresa", "")
            idx = st.session_state.job_index.get(empresa)
            if idx is not None and idx < len(st.session_state.job["status"]):
                st.session_state.job["status"][idx]["STATUS"] = evt.get("status", "")
                st.session_state.job["status"][idx]["DETALHE"] = evt.get("detalhe", "")

        elif t == "log":
            st.session_state.job["logs"].append(evt.get("message", ""))

        elif t == "error":
            st.session_state.job["error"] = evt.get("message", "Erro desconhecido.")
            st.session_state.job["active"] = False

        elif t == "done":
            st.session_state.job["active"] = False


# ----------------------------
# App
# ----------------------------
user = require_login()
page = sidebar(user)
render_topbar(user)
cfg = cfgmod.load_config()

job_init_if_needed()
drain_events()

if page == "📊 Painel":
    st.title("📊 Painel")
    df = data_store.ler_clientes(cfg.caminho_planilha)
    ativos = (df["ATIVO"].astype(str).str.upper().str.strip() == "S").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Clientes (total)", len(df))
    c2.metric("Clientes ativos", int(ativos))
    c3.metric("Competência padrão", "Mês anterior")
    c4.metric("Execução em andamento", "SIM" if st.session_state.job["active"] else "NÃO")

    st.subheader("Clientes (visão rápida)")
    st.dataframe(df, use_container_width=True, hide_index=True)

elif page == "🧾 Processar NFS-e":
    st.title("🧾 Processar NFS-e")

    hoje = date.today()
    ano_padrao = hoje.year
    mes_padrao = hoje.month - 1 or 12
    if hoje.month == 1:
        ano_padrao -= 1

    colA, colB, colC = st.columns([1, 1, 2])
    ano = colA.number_input("Ano (competência)", 2020, 2100, int(ano_padrao), 1)
    mes = colB.number_input("Mês (competência)", 1, 12, int(mes_padrao), 1)
    somente_login = colC.checkbox("Somente clientes com LOGIN/SENHA", value=True)

    df = data_store.ler_clientes(cfg.caminho_planilha)
    df["ATIVO"] = df["ATIVO"].astype(str).str.upper().str.strip()
    df_vis = df[df["ATIVO"] == "S"].copy()
    if somente_login:
        df_vis = df_vis[df_vis["TIPO_ACESSO"].astype(str).str.upper().str.strip() == "LOGIN_SENHA"]

    empresas = df_vis["EMPRESA"].tolist()

    st.subheader("Selecionar clientes")
    selecionadas = st.multiselect("Clientes ativos", options=empresas, default=empresas[:10])

    c1, c2, _ = st.columns([1, 1, 2])
    iniciar = c1.button("✅ Iniciar", use_container_width=True, disabled=st.session_state.job["active"])
    parar = c2.button("⛔ Parar", use_container_width=True, disabled=not st.session_state.job["active"])

    if iniciar:
        st.session_state.job = {"active": True, "status": [], "logs": [], "error": None, "output_folder": None}
        st.session_state.job_index = {}
        st.session_state.job_events = queue.Queue()
        st.session_state.job_stop_evt = threading.Event()

        t = threading.Thread(
            target=run_bot_job,
            args=(cfg, int(ano), int(mes), selecionadas, st.session_state.job_events, st.session_state.job_stop_evt),
            daemon=True,
        )
        st.session_state.job_thread = t
        t.start()
        st.success("Execução iniciada.")
        st.rerun()

    if parar and st.session_state.job_stop_evt:
        st.session_state.job_stop_evt.set()
        st.warning("Sinal de parada enviado (para no próximo cliente).")

    if st.session_state.job.get("error"):
        st.error(st.session_state.job["error"])

    st.subheader("Status por cliente")
    st.dataframe(pd.DataFrame(st.session_state.job.get("status", [])), use_container_width=True, hide_index=True)

    st.subheader("Logs (resumo)")
    logs = st.session_state.job.get("logs", [])
    if logs:
        st.code("\n".join(logs[-200:]))
    else:
        st.caption("Sem logs ainda.")

    out_folder = st.session_state.job.get("output_folder")
    if out_folder and os.path.isdir(out_folder) and not st.session_state.job["active"]:
        zip_bytes = _zip_dir(out_folder)
        st.download_button(
            "📥 Baixar resultado (ZIP)",
            data=zip_bytes,
            file_name=os.path.basename(out_folder) + ".zip",
            mime="application/zip",
            use_container_width=True,
        )

    if st.session_state.job["active"]:
        time.sleep(1.0)
        st.rerun()

elif page == "🧩 Clientes":
    st.title("🧩 Clientes")
    if user.role != "admin":
        st.warning("Somente ADMIN pode editar clientes.")
        st.stop()

    df = data_store.ler_clientes(cfg.caminho_planilha)
    st.subheader("Lista")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Editar / Novo")

    empresas = ["(novo)"] + df["EMPRESA"].tolist()
    escolhido = st.selectbox("Cliente", options=empresas)

    if escolhido == "(novo)":
        rec = {c: "" for c in data_store.COLUNAS_OBRIGATORIAS}
        rec["ATIVO"] = "S"
        rec["TIPO_ACESSO"] = "LOGIN_SENHA"
    else:
        rec = df[df["EMPRESA"] == escolhido].iloc[0].to_dict()

    tabs = st.tabs(["Dados", "NFS-e (Acesso)"])
    with tabs[0]:
        rec["EMPRESA"] = st.text_input("Empresa", value=rec.get("EMPRESA", ""))
        rec["CNPJ"] = st.text_input("CNPJ", value=rec.get("CNPJ", ""))
        rec["PREFEITURA"] = st.text_input("Prefeitura", value=rec.get("PREFEITURA", ""))
        rec["ATIVO"] = st.selectbox("Ativo (S/N)", ["S", "N"], index=0 if rec.get("ATIVO", "S") == "S" else 1)

    with tabs[1]:
        rec["TIPO_ACESSO"] = st.selectbox(
            "Tipo de acesso",
            ["LOGIN_SENHA", "CERTIFICADO"],
            index=0 if rec.get("TIPO_ACESSO", "LOGIN_SENHA") == "LOGIN_SENHA" else 1,
        )
        rec["LOGIN"] = st.text_input("Login (se houver)", value=rec.get("LOGIN", ""))
        rec["SENHA"] = st.text_input("Senha (se houver)", value=rec.get("SENHA", ""), type="password")
        rec["IDENT_CERT"] = st.text_input("IDENT_CERT (certificado)", value=rec.get("IDENT_CERT", ""))
        rec["IMG_CERT"] = st.text_input("IMG_CERT (arquivo imagem)", value=rec.get("IMG_CERT", ""))

    a, b = st.columns([1, 1])
    if a.button("💾 Salvar", use_container_width=True):
        df2 = df.copy()
        if escolhido != "(novo)":
            df2 = df2[df2["EMPRESA"] != escolhido]
        df2 = pd.concat([df2, pd.DataFrame([rec])], ignore_index=True).fillna("")
        df2 = df2[data_store.COLUNAS_OBRIGATORIAS]
        data_store.salvar_clientes(cfg.caminho_planilha, df2)
        st.success("Salvo.")
        st.rerun()

    if escolhido != "(novo)" and b.button("🗑️ Excluir", use_container_width=True):
        df2 = df[df["EMPRESA"] != escolhido].copy()
        data_store.salvar_clientes(cfg.caminho_planilha, df2)
        st.success("Excluído.")
        st.rerun()

elif page == "⚙️ Configurações":
    st.title("⚙️ Configurações")
    if user.role != "admin":
        st.warning("Somente ADMIN pode alterar configurações.")
        st.stop()

    st.caption("Caminhos do seu ambiente. Salva em `config.local.json`.")
    c1, c2 = st.columns(2)
    caminho_planilha = c1.text_input("Caminho da planilha", value=cfg.caminho_planilha)
    pasta_saida = c2.text_input("Pasta base de saída", value=cfg.pasta_base_saida)
    c3, c4 = st.columns(2)
    pasta_download = c3.text_input("Pasta de download temporário", value=cfg.pasta_download_temp)
    pasta_imagens = c4.text_input("Pasta imagens (certificado)", value=cfg.pasta_imagens_cert)
    delay = st.number_input("Delay ações (segundos)", 0.0, 10.0, float(cfg.delay_acao), 0.1)

    if st.button("💾 Salvar configurações", use_container_width=True):
        novo = cfgmod.AppConfig(caminho_planilha, pasta_saida, pasta_download, pasta_imagens, float(delay))
        cfgmod.save_config(novo)
        st.success("Config salvo.")
        st.rerun()

elif page == "🧾 Auditoria":
    st.title("🧾 Auditoria")
    base = os.path.abspath(cfg.pasta_base_saida)
    st.write(f"Base de saída: `{base}`")

    if not os.path.isdir(base):
        st.warning("Pasta base de saída não existe ainda.")
        st.stop()

    pastas = sorted([p for p in os.listdir(base) if os.path.isdir(os.path.join(base, p))], reverse=True)
    sel = st.selectbox("Competências", options=pastas) if pastas else None

    if sel:
        folder = os.path.join(base, sel)
        files = []
        for root, _, fns in os.walk(folder):
            for fn in fns:
                files.append(os.path.relpath(os.path.join(root, fn), folder))
        st.subheader("Arquivos")
        st.dataframe(pd.DataFrame({"arquivo": files}), use_container_width=True, hide_index=True)

        zip_bytes = _zip_dir(folder)
        st.download_button(
            "📥 Baixar competência (ZIP)",
            data=zip_bytes,
            file_name=sel + ".zip",
            mime="application/zip",
            use_container_width=True,
        )
