"""
iG2P · Formulário de Inscrição — Processo Seletivo Público Simplificado nº 04/2026
Secretaria Municipal de Saúde de Cássia - MG · Centro de Especialidades Odontológicas (CEO)
Padrão visual dos dashboards iG2P (modo claro/escuro, faixas indigo, cards).

Regras (iguais ao app de Cássia médico):
  - Gravação ATÔMICA/idempotente por token (evita duplicidade por duplo clique).
  - Protocolo sequencial lido da própria planilha.
  - Tela de loading + rolagem ao topo + botão desabilitado durante o envio.
  - E-mail de confirmação com o número de inscrição.
  - Anexos enviados ao Google Drive; link vai para a planilha.
"""

import os
import re
import io
import ssl
import uuid
import base64
import random
import smtplib
import unicodedata
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage

import streamlit as st
import streamlit.components.v1 as components

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PLANILHA_LOCAL = os.path.join(APP_DIR, "respostas_ceo.xlsx")
BANDEIRA_PATH = os.path.join(APP_DIR, "cássiabandeira.png")

# ----------------------------------------------------------------------------
# JANELA DE INSCRIÇÕES  ——  conforme o Edital nº 04/2026 (inscrições 01 a 07/12/2026).
# Fuso de Brasília. Use None em ambas para deixar BLOQUEADO ("em breve").
# ----------------------------------------------------------------------------
TZ = timezone(timedelta(hours=-3))  # Horário de Brasília (UTC-3)
DATA_ABERTURA = datetime(2026, 12, 1, 0, 0, tzinfo=TZ)      # 01/12/2026 00h00
DATA_ENCERRAMENTO = datetime(2026, 12, 7, 23, 59, tzinfo=TZ)  # 07/12/2026 23h59

# Numeração sequencial do protocolo (acompanha a ordem das inscrições).
NUMERO_INICIAL = 1000

# Anexos (documentos)
EXTS_ANEXO = ["pdf", "jpg", "jpeg", "png"]
TAMANHO_MAX_MB = 10

# Funções pretendidas (seleção única) — exatamente como a Ficha de Inscrição (Anexo I)
FUNCOES = [
    "Cirurgião-Dentista – Área de atuação em Endodontia (CEO)",
    "Cirurgião-Dentista – Área de atuação em Cirurgia Oral e Odontologia para Pacientes com Necessidades Especiais (CEO)",
    "Cirurgião-Dentista – Área de atuação em Periodontia e Odontologia para Pacientes com Necessidades Especiais (CEO)",
    "Cirurgião-Dentista – 40 horas semanais (Atenção Primária à Saúde)",
    "Auxiliar em Saúde Bucal (ASB) – 40 horas semanais (CEO)",
]


def e_dentista(funcao):
    return (funcao or "").startswith("Cirurgião-Dentista")


# Anexos: (chave, rótulo, obrigatório, somente_dentista)
ANEXOS = [
    ("diploma", "Diploma de graduação em Odontologia (Cirurgião-Dentista) ou comprovante de habilitação em Auxiliar em Saúde Bucal (ASB)", True, False),
    ("cro_cert", "Certidão de Regularidade Profissional emitida pelo CRO (inscrição ativa e regular)", True, False),
    ("especializacao", "Pós-Graduação lato sensu (Especialização) na área da função", False, True),
    ("mestrado", "Pós-Graduação stricto sensu (Mestrado ou Doutorado) na área da função", False, True),
    ("cursos", "Cursos de aperfeiçoamento, atualização e extensão", False, False),
    ("experiencia", "Comprovação de experiência profissional (declaração/certidão de tempo de serviço ou CTPS)", False, False),
]

COLUNAS_PLANILHA = [
    "Data/Hora do Envio", "Protocolo", "Função Pretendida", "Nome Completo", "CPF",
    "Data de Nascimento", "E-mail", "Telefone/WhatsApp", "Endereço",
    "Registro CRO (nº)", "Registro CRO (UF)",
    "Diploma/Habilitação (link)", "Certidão Regularidade CRO (link)",
    "Pós-Grad. Especialização (link)", "Mestrado/Doutorado (link)",
    "Cursos/Aperfeiçoamento (link)", "Experiência Profissional (link)",
    "Declaração - Ciência do Edital", "Declaração - Veracidade", "Declaração - LGPD",
]

# ----------------------------------------------------------------------------
# Tema (padrão iG2P) — dicionário claro/escuro dirige todo o CSS inline
# ----------------------------------------------------------------------------
TEMAS = {
    "claro": {
        "bg": "#F4F5FB", "card": "#FFFFFF", "texto": "#1E1F36",
        "texto2": "#5A5C77", "borda": "#E3E5F0", "primaria": "#5B5BD6",
        "primaria_grad": "#7C5CFF", "faixa_txt": "#FFFFFF", "input_bg": "#FFFFFF",
        "sombra": "0 8px 24px rgba(91,91,214,0.08)", "ok": "#1FA971", "erro": "#E5484D",
    },
    "escuro": {
        "bg": "#0F1020", "card": "#1A1B33", "texto": "#ECEDF7",
        "texto2": "#9FA1C0", "borda": "#2C2E4A", "primaria": "#7C7CFF",
        "primaria_grad": "#9D7BFF", "faixa_txt": "#FFFFFF", "input_bg": "#23244180",
        "sombra": "0 8px 24px rgba(0,0,0,0.35)", "ok": "#3DD68C", "erro": "#FF6369",
    },
}

st.set_page_config(
    page_title="Inscrição · Processo Seletivo CEO · Cássia",
    page_icon="🦷",
    layout="centered",
)

for k, v in {"modo": "claro", "enviado": False, "protocolo": None, "destino": None,
             "motivo": None, "email_ok": None, "email_dest": None,
             "processando": False, "dados_pendentes": None, "erro_envio": None,
             "token_gravado": None, "duplicado": False}.items():
    if k not in st.session_state:
        st.session_state[k] = v

T = TEMAS[st.session_state.modo]


# ----------------------------------------------------------------------------
# CSS
# ----------------------------------------------------------------------------
def aplicar_css(t):
    st.markdown(f"""
    <style>
    .stApp {{ background:{t['bg']}; }}
    #MainMenu, footer, header {{ visibility:hidden; }}
    .block-container {{ padding-top:1.2rem; padding-bottom:3rem; max-width:820px; }}

    .ig2p-header {{
        background:linear-gradient(135deg,{t['primaria']} 0%,{t['primaria_grad']} 100%);
        border-radius:20px; padding:26px 30px; color:#fff;
        box-shadow:{t['sombra']}; display:flex; align-items:center; gap:20px;
        margin-bottom:6px;
    }}
    .ig2p-flag {{
        width:62px; height:62px; border-radius:14px; flex:0 0 auto;
        background:#fff; display:flex; align-items:center;
        justify-content:center; font-weight:800; font-size:20px; letter-spacing:1px;
        color:{t['primaria']}; border:2px solid rgba(255,255,255,.35);
        overflow:hidden; padding:7px; box-sizing:border-box;
    }}
    .ig2p-flag img {{ width:100%; height:100%; object-fit:contain; }}
    .ig2p-header h1 {{ font-size:1.38rem; margin:0; line-height:1.18; font-weight:800; }}
    .ig2p-header p {{ margin:4px 0 0; opacity:.92; font-size:.86rem; }}
    .ig2p-chip {{
        display:inline-block; margin-top:8px; padding:3px 12px; border-radius:30px;
        background:rgba(255,255,255,.20); font-size:.72rem; font-weight:600;
    }}

    .ig2p-intro {{
        background:{t['card']}; color:{t['texto2']}; border:1px solid {t['borda']};
        border-radius:16px; padding:16px 20px; font-size:.84rem; line-height:1.5;
        margin:14px 0 4px; box-shadow:{t['sombra']};
    }}
    .ig2p-intro b {{ color:{t['texto']}; }}

    .ig2p-faixa {{
        background:linear-gradient(135deg,{t['primaria']} 0%,{t['primaria_grad']} 100%);
        color:{t['faixa_txt']}; padding:12px 18px; border-radius:12px;
        font-weight:700; font-size:1rem; margin:26px 0 10px;
        box-shadow:{t['sombra']};
    }}
    .ig2p-legal {{
        background:{t['card']}; border:1px dashed {t['borda']}; border-radius:12px;
        padding:14px 18px; font-size:.78rem; color:{t['texto2']}; line-height:1.55;
        margin-bottom:6px;
    }}
    .ig2p-legal b {{ color:{t['texto']}; }}

    label, .stMarkdown, .stRadio, p {{ color:{t['texto']} !important; }}
    .stRadio label p {{ color:{t['texto']} !important; }}

    .stTextInput input, .stDateInput input, .stTextArea textarea {{
        background:{t['input_bg']} !important; color:{t['texto']} !important;
        border:1px solid {t['borda']} !important; border-radius:10px !important;
    }}
    .stTextInput input:focus, .stTextArea textarea:focus {{
        border-color:{t['primaria']} !important;
        box-shadow:0 0 0 3px {t['primaria']}33 !important;
    }}
    .stSelectbox div[data-baseweb="select"] > div {{
        background:{t['input_bg']} !important; border:1px solid {t['borda']} !important;
        border-radius:10px !important; color:{t['texto']} !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background:{t['input_bg']} !important; border:1px dashed {t['borda']} !important;
        border-radius:12px !important;
    }}

    .stButton > button {{
        background:linear-gradient(135deg,{t['primaria']} 0%,{t['primaria_grad']} 100%);
        color:#fff !important; border:none; border-radius:12px; padding:12px 28px;
        font-weight:700; font-size:.98rem; width:100%; transition:transform .15s ease;
        box-shadow:{t['sombra']};
    }}
    .stButton > button:hover {{ transform:translateY(-2px); }}

    .ig2p-card {{
        background:{t['card']}; border:1px solid {t['borda']}; border-radius:16px;
        padding:22px 24px; margin-top:10px; box-shadow:{t['sombra']};
    }}
    .ig2p-success {{
        background:{t['card']}; border:1px solid {t['ok']}; border-radius:18px;
        padding:30px; text-align:center; box-shadow:{t['sombra']};
    }}
    .ig2p-success .check {{
        width:64px; height:64px; border-radius:50%; background:{t['ok']}22;
        color:{t['ok']}; display:flex; align-items:center; justify-content:center;
        font-size:34px; margin:0 auto 14px;
    }}
    .ig2p-prot {{
        font-family:ui-monospace,monospace; font-size:1.4rem; font-weight:800;
        color:{t['primaria']}; letter-spacing:2px; margin:6px 0 2px;
    }}
    .ig2p-rodape {{ text-align:center; color:{t['texto2']}; font-size:.72rem; margin-top:26px; }}

    /* Loading */
    .ig2p-loading {{
        background:{t['card']}; border:2px solid {t['primaria']}; border-radius:18px;
        padding:60px 30px; text-align:center; box-shadow:{t['sombra']};
        min-height:52vh; display:flex; flex-direction:column;
        align-items:center; justify-content:center;
    }}
    .ig2p-spinner {{
        width:76px; height:76px; margin:0 auto 24px; border-radius:50%;
        border:7px solid {t['borda']}; border-top-color:{t['primaria']};
        animation:ig2pspin 0.9s linear infinite;
    }}
    @keyframes ig2pspin {{ to {{ transform:rotate(360deg); }} }}
    .ig2p-loading h2 {{ font-size:1.6rem; margin:0; }}
    .ig2p-pulse {{ animation:ig2ppulse 1.6s ease-in-out infinite; }}
    @keyframes ig2ppulse {{ 0%,100% {{ opacity:1; }} 50% {{ opacity:.55; }} }}

    [data-testid="stElementContainer"]:has(iframe),
    .element-container:has(iframe), [data-testid="stIFrame"] {{
        height:0!important; min-height:0!important; margin:0!important;
        padding:0!important; overflow:hidden!important;
    }}
    </style>
    """, unsafe_allow_html=True)


aplicar_css(T)


def flag_html():
    if os.path.exists(BANDEIRA_PATH):
        b64 = base64.b64encode(open(BANDEIRA_PATH, "rb").read()).decode()
        return f'<div class="ig2p-flag"><img src="data:image/png;base64,{b64}"/></div>'
    return '<div class="ig2p-flag">CA</div>'


def scroll_top():
    components.html("""<script>
    (function(){try{var d=window.parent.document;
    [d.querySelector('section.main'),d.querySelector('[data-testid="stMain"]'),
     d.querySelector('[data-testid="stAppViewContainer"]'),
     d.scrollingElement,d.documentElement,d.body
    ].forEach(function(e){if(e)e.scrollTop=0;});
    if(window.parent.scrollTo)window.parent.scrollTo(0,0);}catch(e){}})();
    </script>""", height=1)


# ----------------------------------------------------------------------------
# Validações
# ----------------------------------------------------------------------------
def so_digitos(v):
    return re.sub(r"\D", "", v or "")


def email_valido(e):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", (e or "").strip()))


def cpf_valido(cpf):
    cpf = so_digitos(cpf)
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for i in (9, 10):
        soma = sum(int(cpf[n]) * ((i + 1) - n) for n in range(i))
        dig = (soma * 10) % 11
        dig = 0 if dig == 10 else dig
        if dig != int(cpf[i]):
            return False
    return True


def telefone_valido(tel):
    return 10 <= len(so_digitos(tel)) <= 11


def data_valida(s):
    try:
        datetime.strptime((s or "").strip(), "%d/%m/%Y")
        return True
    except ValueError:
        return False


def slug(texto):
    t = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9]+", "_", t).strip("_")
    return t[:40] or "candidato"


# ----------------------------------------------------------------------------
# Máscaras
# ----------------------------------------------------------------------------
def mascara_cpf(texto):
    d = so_digitos(texto)[:11]
    if len(d) > 9:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) > 6:
        return f"{d[:3]}.{d[3:6]}.{d[6:]}"
    if len(d) > 3:
        return f"{d[:3]}.{d[3:]}"
    return d


def mascara_data(texto):
    d = so_digitos(texto)[:8]
    if len(d) > 4:
        return f"{d[:2]}/{d[2:4]}/{d[4:]}"
    if len(d) > 2:
        return f"{d[:2]}/{d[2:]}"
    return d


def mascara_tel(texto):
    d = so_digitos(texto)[:11]
    if len(d) > 10:
        return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    if len(d) > 6:
        return f"({d[:2]}) {d[2:6]}-{d[6:]}"
    if len(d) > 2:
        return f"({d[:2]}) {d[2:]}"
    if len(d) > 0:
        return f"({d}"
    return d


def fmt_cpf():
    st.session_state.cpf = mascara_cpf(st.session_state.cpf)


def fmt_nascimento():
    st.session_state.nascimento = mascara_data(st.session_state.nascimento)


def fmt_tel():
    st.session_state.telefone = mascara_tel(st.session_state.telefone)


def fmt_cro_uf():
    st.session_state.cro_uf = re.sub(r"[^A-Za-z]", "", st.session_state.cro_uf).upper()[:2]


# ----------------------------------------------------------------------------
# Janela de inscrições
# ----------------------------------------------------------------------------
def agora():
    return datetime.now(TZ)


def status_inscricoes():
    if DATA_ABERTURA is None:
        return "antes"
    n = agora()
    if n < DATA_ABERTURA:
        return "antes"
    if DATA_ENCERRAMENTO is not None and n > DATA_ENCERRAMENTO:
        return "encerrado"
    return "aberto"


def fmt_data_br(dt):
    return dt.strftime("%d/%m/%Y às %Hh%M") if dt else "a definir"


# ----------------------------------------------------------------------------
# Anexos — upload para o Google Drive (pasta em [drive].folder_id)
# ----------------------------------------------------------------------------
def _drive_service():
    from google.oauth2.service_account import Credentials  # type: ignore
    from googleapiclient.discovery import build  # type: ignore
    escopos = ["https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=escopos)
    return build("drive", "v3", credentials=cred, cache_discovery=False)


def salvar_anexo(uploaded, token, nome_candidato, chave):
    """Envia um anexo ao Drive e retorna (link, destino, motivo)."""
    if uploaded is None:
        return "", "vazio", None
    ext = os.path.splitext(uploaded.name)[1].lower()
    nome_final = f"{token}_{chave}_{slug(nome_candidato)}{ext}"
    folder_id = st.secrets.get("drive", {}).get("folder_id", "")
    motivo_drive = ""
    if "gcp_service_account" in st.secrets and folder_id:
        try:
            from googleapiclient.http import MediaIoBaseUpload  # type: ignore
            service = _drive_service()
            media = MediaIoBaseUpload(
                io.BytesIO(uploaded.getvalue()),
                mimetype=uploaded.type or "application/octet-stream", resumable=False)
            arq = service.files().create(
                body={"name": nome_final, "parents": [folder_id]},
                media_body=media, fields="id,webViewLink",
                supportsAllDrives=True).execute()
            return arq.get("webViewLink", arq.get("id", "")), "drive", None
        except Exception as e:
            motivo_drive = f"{type(e).__name__}: {e}"
    else:
        motivo_drive = "Drive não configurado ([drive].folder_id ausente nos secrets)."
    # Fallback local (apenas testes; no Streamlit Cloud o disco é efêmero)
    try:
        pasta = os.path.join(APP_DIR, "anexos")
        os.makedirs(pasta, exist_ok=True)
        caminho = os.path.join(pasta, nome_final)
        with open(caminho, "wb") as fp:
            fp.write(uploaded.getvalue())
        return caminho, "local", motivo_drive
    except Exception as e:
        return "", "erro", f"Drive: {motivo_drive} | Local: {type(e).__name__}: {e}"


# ----------------------------------------------------------------------------
# E-mail de confirmação (SMTP, via st.secrets[email])
# ----------------------------------------------------------------------------
def enviar_email_confirmacao(destinatario, nome, protocolo, funcao):
    if "email" not in st.secrets:
        return False, "Seção [email] não configurada nos secrets."
    cfg = st.secrets["email"]
    try:
        remetente = cfg["remetente"]
        usuario = cfg.get("usuario", remetente)
        senha = cfg["senha"]
        host = cfg.get("host", "smtp.gmail.com")
        porta = int(cfg.get("porta", 465))
        nome_exib = cfg.get("nome_exibicao", "Processo Seletivo CEO - Cássia")
    except Exception as e:
        return False, f"Configuração de e-mail incompleta: {e}"

    assunto = "Confirmação de Inscrição — Processo Seletivo nº 04/2026 (CEO) · Cássia-MG"
    texto = (
        f"Olá, {nome}!\n\n"
        f"Sua inscrição no Processo Seletivo Público Simplificado nº 04/2026 "
        f"(Centro de Especialidades Odontológicas) da Secretaria Municipal de Saúde de "
        f"Cássia - MG foi recebida com sucesso.\n\n"
        f"Número de inscrição (protocolo): {protocolo}\n"
        f"Função pretendida: {funcao}\n"
        f"Data/hora: {agora().strftime('%d/%m/%Y às %Hh%M')}\n\n"
        f"Guarde este número — ele é o comprovante da sua inscrição.\n\n"
        f"As informações prestadas são de inteira responsabilidade do candidato.\n\n"
        f"Este é um e-mail automático, não responda.\n— {nome_exib}"
    )
    html = (
        f'<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;'
        f'border:1px solid #E3E5F0;border-radius:14px;overflow:hidden">'
        f'<div style="background:linear-gradient(135deg,#5B5BD6,#7C5CFF);color:#fff;padding:22px 26px">'
        f'<h2 style="margin:0;font-size:18px">Inscrição confirmada ✓</h2>'
        f'<p style="margin:4px 0 0;opacity:.9;font-size:13px">Processo Seletivo nº 04/2026 · CEO · Cássia-MG</p>'
        f'</div><div style="padding:24px 26px;color:#1E1F36;font-size:14px;line-height:1.6">'
        f'<p>Olá, <b>{nome}</b>! Sua inscrição foi recebida com sucesso.</p>'
        f'<p style="margin:18px 0;text-align:center"><span style="display:inline-block;'
        f'background:#5B5BD61a;color:#5B5BD6;font-weight:800;font-size:20px;letter-spacing:2px;'
        f'padding:12px 22px;border-radius:10px">{protocolo}</span></p>'
        f'<p style="margin:6px 0"><b>Função pretendida:</b> {funcao}</p>'
        f'<p style="margin:6px 0"><b>Data/hora:</b> {agora().strftime("%d/%m/%Y às %Hh%M")}</p>'
        f'<p style="margin-top:16px;color:#5A5C77;font-size:13px">Guarde este número — '
        f'ele é o comprovante da sua inscrição.</p>'
        f'<p style="color:#9FA1C0;font-size:12px;margin-top:18px">Este é um e-mail automático, não responda.</p>'
        f'</div></div>'
    )

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = f"{nome_exib} <{remetente}>"
    msg["To"] = destinatario
    msg.set_content(texto)
    msg.add_alternative(html, subtype="html")
    try:
        ctx = ssl.create_default_context()
        if porta == 465:
            with smtplib.SMTP_SSL(host, porta, context=ctx, timeout=20) as s:
                s.login(usuario, senha)
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, porta, timeout=20) as s:
                s.starttls(context=ctx)
                s.login(usuario, senha)
                s.send_message(msg)
        return True, None
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


# ----------------------------------------------------------------------------
# Persistência — Google Sheets (fonte oficial)
# ----------------------------------------------------------------------------
def conectar_sheets():
    if "gcp_service_account" not in st.secrets:
        return None, ("Secrets não foram lidos. No Streamlit Cloud: cole o conteúdo do "
                      "secrets.toml em Manage app → Settings → Secrets.")
    sheet_id = st.secrets.get("planilha", {}).get("sheet_id", "")
    if not sheet_id or "COLE_AQUI" in sheet_id:
        return None, "sheet_id não configurado em [planilha] no secrets.toml."
    try:
        import gspread  # type: ignore
        from google.oauth2.service_account import Credentials  # type: ignore
    except ImportError:
        return None, "Bibliotecas gspread/google-auth não instaladas. Rode: pip install gspread google-auth"
    try:
        escopos = ["https://www.googleapis.com/auth/spreadsheets"]
        cred = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=escopos)
        gc = gspread.authorize(cred)
        return gc.open_by_key(sheet_id).sheet1, None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def buscar_inscricao_por_cpf(cpf):
    """Consulta (sem gravar) se o CPF já tem inscrição. Retorna o protocolo existente ou None.
    CPF está na coluna 5 (índice 4 na linha)."""
    cpf = so_digitos(cpf)
    if not cpf:
        return None
    try:
        ws, _ = conectar_sheets()
        if ws is not None:
            for r in ws.get_all_values()[1:]:  # pula o cabeçalho
                if len(r) > 4 and so_digitos(r[4]) == cpf:
                    return r[1] if len(r) > 1 and r[1] else "—"
    except Exception:
        pass
    return None


def salvar_google_sheets(linha, token):
    """Gravação IDEMPOTENTE e à prova de corrida, usando o token na última coluna.
    (1) checa antes; (2) grava; (3) recheca e remove excedentes do mesmo token.
    Depois do append, nenhum erro posterior retorna 'erro'. (status, motivo, protocolo)."""
    ws, motivo = conectar_sheets()
    if ws is None:
        return "erro", motivo, None

    col_token = len(COLUNAS_PLANILHA) + 1

    def linhas_do_token():
        try:
            col = ws.col_values(col_token)
        except Exception:
            col = []
        return [i + 1 for i, v in enumerate(col) if v and v == token]

    try:
        if not ws.row_values(1):
            ws.append_row(COLUNAS_PLANILHA + ["Token"], value_input_option="USER_ENTERED")
        if token:
            existentes = linhas_do_token()
            if existentes:
                idx = existentes[0]
                proto = ws.cell(idx, 2).value or str(NUMERO_INICIAL + (idx - 2))
                return "ok", None, str(proto)
        # Rede de segurança anti-duplicidade por CPF (coluna 5 = índice 4 na linha).
        cpf_novo = so_digitos(linha[4]) if len(linha) > 4 else ""
        if cpf_novo:
            for r in ws.get_all_values()[1:]:
                if len(r) > 4 and so_digitos(r[4]) == cpf_novo:
                    return "duplicado", None, (r[1] if len(r) > 1 and r[1] else "—")
        ws.append_row(linha + [token], value_input_option="USER_ENTERED")
    except Exception as e:
        return "erro", f"Falha ao gravar a linha: {type(e).__name__}: {e}", None

    protocolo = None
    try:
        existentes = linhas_do_token() if token else []
        if len(existentes) > 1:
            for idx in sorted(existentes[1:], reverse=True):
                try:
                    ws.delete_rows(idx)
                except Exception:
                    pass
            existentes = linhas_do_token()
        row = existentes[0] if existentes else len(ws.col_values(1))
        atual = ws.cell(row, 2).value
        if atual:
            protocolo = str(atual)
        else:
            protocolo = str(NUMERO_INICIAL + (row - 2))
            ws.update_cell(row, 2, protocolo)
    except Exception:
        try:
            existentes = linhas_do_token()
            if existentes:
                row = existentes[0]
                protocolo = str(ws.cell(row, 2).value or (NUMERO_INICIAL + (row - 2)))
        except Exception:
            pass
    return "ok", None, protocolo


def salvar_resposta(linha, token):
    status, motivo, protocolo = salvar_google_sheets(linha, token)
    if status == "ok":
        return "google", None, protocolo
    if status == "duplicado":
        return "duplicado", None, protocolo
    return "erro", motivo, None


# ----------------------------------------------------------------------------
# Tela de sucesso
# ----------------------------------------------------------------------------
if st.session_state.enviado:
    scroll_top()
    if st.session_state.duplicado:
        st.markdown(
            f'<div class="ig2p-success" style="border-color:{T["primaria"]};">'
            f'<div class="check" style="background:{T["primaria"]}1a;color:{T["primaria"]};">ℹ️</div>'
            f'<h2 style="color:{T["texto"]};margin:0;">Você já possui uma inscrição</h2>'
            f'<p style="color:{T["texto2"]};margin:8px 0 0;">Encontramos uma inscrição com este CPF. '
            f'Não é necessário se inscrever novamente. Seu número de inscrição é:</p>'
            f'<div class="ig2p-prot">{st.session_state.protocolo}</div>'
            f'<p style="color:{T["texto2"]};font-size:.8rem;margin-top:14px;">'
            f'Se você acredita que isso é um engano, entre em contato pelo e-mail '
            f'<a href="mailto:smscassiamg@gmail.com" style="color:{T["primaria"]};font-weight:700;'
            f'text-decoration:none;">smscassiamg@gmail.com</a>.</p></div>', unsafe_allow_html=True)
        if st.button("Voltar"):
            for k in ("enviado", "protocolo", "duplicado"):
                st.session_state[k] = False if k in ("enviado", "duplicado") else None
            st.rerun()
        st.markdown('<div class="ig2p-rodape">iG2P · Inteligência em Gestão Pública · G3ST4O · Prefeitura de Cássia - MG</div>',
                    unsafe_allow_html=True)
        st.stop()
    if st.session_state.email_ok is True:
        email_badge = (f'<p style="color:{T["texto2"]};font-size:.85rem;margin-top:14px;">'
                       f'📧 Um e-mail de confirmação foi enviado para '
                       f'<b>{st.session_state.email_dest}</b>.</p>')
    else:
        email_badge = ""
    contato_box = (
        f'<div style="margin-top:18px;padding:14px 16px;border-radius:12px;'
        f'background:{T["primaria"]}12;border:1px solid {T["borda"]};">'
        f'<p style="color:{T["texto"]};font-size:.85rem;margin:0;font-weight:600;">Ficou com alguma dúvida?</p>'
        f'<p style="color:{T["texto2"]};font-size:.82rem;margin:4px 0 0;">'
        f'Em caso de dúvidas ou problemas com sua inscrição, entre em contato pelo e-mail:<br>'
        f'<a href="mailto:smscassiamg@gmail.com" style="color:{T["primaria"]};font-weight:700;'
        f'text-decoration:none;">smscassiamg@gmail.com</a></p></div>'
    )
    st.markdown(
        f'<div class="ig2p-success">'
        f'<div class="check">✓</div>'
        f'<h2 style="color:{T["texto"]};margin:0;">Inscrição enviada com sucesso!</h2>'
        f'<p style="color:{T["texto2"]};margin:8px 0 0;">Guarde o número de inscrição abaixo:</p>'
        f'<div class="ig2p-prot">{st.session_state.protocolo}</div>'
        f'<p style="color:{T["texto2"]};font-size:.8rem;margin-top:14px;">'
        f'Suas informações e documentos foram registrados. A constatação de falsidade ou omissão '
        f'de dados implicará na eliminação do candidato.</p>'
        f'{email_badge}{contato_box}'
        f'</div>', unsafe_allow_html=True)
    if st.button("Nova inscrição"):
        for k in ("enviado", "protocolo", "duplicado"):
            st.session_state[k] = False if k in ("enviado", "duplicado") else None
        st.rerun()
    st.markdown('<div class="ig2p-rodape">iG2P · Inteligência em Gestão Pública · G3ST4O · Prefeitura de Cássia - MG</div>',
                unsafe_allow_html=True)
    st.stop()


# ----------------------------------------------------------------------------
# Tela de processamento (loading) — grava e vai ao sucesso
# ----------------------------------------------------------------------------
if st.session_state.processando:
    scroll_top()
    st.markdown(f"""
    <div class="ig2p-header" style="margin-bottom:18px;">
        {flag_html()}
        <div><h1>Processo Seletivo nº 04/2026 — CEO</h1>
        <p>Secretaria Municipal de Saúde · Município de Cássia - MG</p></div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.erro_envio:
        st.markdown(
            f'<div class="ig2p-card" style="border-color:{T["erro"]};">'
            f'<b style="color:{T["erro"]};">Não foi possível registrar sua inscrição.</b><br>'
            f'<span style="color:{T["texto2"]};font-size:.85rem;">'
            f'Nenhum dado foi gravado. Tente novamente em instantes.<br>'
            f'<b>Detalhe técnico:</b> {st.session_state.erro_envio}</span></div>',
            unsafe_allow_html=True)
        if st.button("Tentar novamente"):
            st.session_state.erro_envio = None
            st.rerun()
        st.stop()

    st.markdown(
        f'<div class="ig2p-loading">'
        f'<div class="ig2p-spinner"></div>'
        f'<h2 class="ig2p-pulse" style="color:{T["texto"]};">Inscrição sendo realizada…</h2>'
        f'<p style="color:{T["texto2"]};margin:12px 0 0;font-size:1rem;">Aguarde alguns segundos.</p>'
        f'<p style="color:{T["texto2"]};margin:6px 0 0;font-size:.9rem;">'
        f'<b>Não feche a página e não clique novamente.</b></p>'
        f'</div>', unsafe_allow_html=True)

    dados = st.session_state.dados_pendentes or {}
    token_atual = dados.get("token", "")

    if token_atual and st.session_state.token_gravado == token_atual:
        st.session_state.dados_pendentes = None
        st.session_state.processando = False
        st.session_state.enviado = True
        st.rerun()

    destino, motivo, protocolo = salvar_resposta(dados.get("linha", []), token_atual)
    if destino == "erro":
        st.session_state.erro_envio = motivo
        st.rerun()
    elif destino == "duplicado":
        st.session_state.protocolo = protocolo
        st.session_state.duplicado = True
        st.session_state.email_ok = None
        st.session_state.dados_pendentes = None
        st.session_state.processando = False
        st.session_state.enviado = True
        st.rerun()
    else:
        email_ok, _ = enviar_email_confirmacao(
            dados.get("email", ""), dados.get("nome", ""), protocolo, dados.get("funcao", ""))
        st.session_state.email_ok = email_ok
        st.session_state.email_dest = dados.get("email", "")
        st.session_state.destino = destino
        st.session_state.motivo = motivo
        st.session_state.protocolo = protocolo
        st.session_state.token_gravado = token_atual
        st.session_state.dados_pendentes = None
        st.session_state.processando = False
        st.session_state.enviado = True
        st.rerun()
    st.stop()


# ----------------------------------------------------------------------------
# Bloqueio por janela de datas
# ----------------------------------------------------------------------------
_status = status_inscricoes()
if _status != "aberto":
    if _status == "antes":
        icone, titulo = "🔒", "Inscrições ainda não abertas"
        msg = ("O período de inscrições será divulgado em breve." if DATA_ABERTURA is None
               else f"As inscrições abrem em <b>{fmt_data_br(DATA_ABERTURA)}</b>.")
    else:
        icone, titulo = "⛔", "Inscrições encerradas"
        msg = f"O período de inscrições foi encerrado em <b>{fmt_data_br(DATA_ENCERRAMENTO)}</b>."
    periodo = ""
    if DATA_ABERTURA is not None:
        periodo = (f'<div style="margin-top:16px;padding:12px 16px;border-radius:12px;'
                   f'background:{T["primaria"]}14;color:{T["texto2"]};font-size:.82rem;">'
                   f'<b style="color:{T["texto"]};">Período de inscrições</b><br>'
                   f'Abertura: {fmt_data_br(DATA_ABERTURA)}<br>'
                   f'Encerramento: {fmt_data_br(DATA_ENCERRAMENTO)}</div>')
    st.markdown(f"""
    <div class="ig2p-header" style="margin-bottom:18px;">
        {flag_html()}
        <div><h1>Processo Seletivo nº 04/2026 — CEO</h1>
        <p>Secretaria Municipal de Saúde · Município de Cássia - MG</p></div>
    </div>
    <div class="ig2p-success">
        <div class="check" style="background:{T['primaria']}1a;color:{T['primaria']};">{icone}</div>
        <h2 style="color:{T['texto']};margin:0;">{titulo}</h2>
        <p style="color:{T['texto2']};margin:10px 0 0;font-size:.9rem;">{msg}</p>
        {periodo}
    </div>
    <div class="ig2p-rodape">iG2P · Inteligência em Gestão Pública · G3ST4O · Prefeitura de Cássia - MG</div>
    """, unsafe_allow_html=True)
    st.stop()


# ----------------------------------------------------------------------------
# Cabeçalho + tema
# ----------------------------------------------------------------------------
col_a, col_b = st.columns([6, 1])
with col_b:
    if st.button("🌙" if st.session_state.modo == "claro" else "☀️", help="Alternar tema"):
        st.session_state.modo = "escuro" if st.session_state.modo == "claro" else "claro"
        st.rerun()

st.markdown(f"""
<div class="ig2p-header">
    {flag_html()}
    <div>
        <h1>Formulário de Inscrição — Processo Seletivo nº 04/2026</h1>
        <p>Secretaria Municipal de Saúde · Município de Cássia · Estado de Minas Gerais</p>
        <span class="ig2p-chip">Centro de Especialidades Odontológicas (CEO)</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="ig2p-intro">
Processo Seletivo Público Simplificado para seleção e contratação de caráter temporário,
destinado ao <b>Centro de Especialidades Odontológicas (CEO)</b> e à Atenção Primária à Saúde
da Secretaria Municipal de Saúde de Cássia - MG.<br><br>
<b>As informações prestadas neste formulário são de inteira responsabilidade do candidato.</b>
A constatação, a qualquer tempo, de falsidade ou omissão de dados implicará na eliminação
do candidato. Não haverá conferência de documentos no ato da inscrição, nem será aceita a
entrega posterior dos mesmos.<br><br>
A inscrição é <b>única por CPF</b> — se já houver uma inscrição com o seu CPF, o sistema mostrará o número existente.<br><br>
<span style="color:{T['erro']};font-weight:600;">*</span> Indica uma pergunta obrigatória.
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 1. FUNÇÃO PRETENDIDA
# ----------------------------------------------------------------------------
st.markdown('<div class="ig2p-faixa">1. Função Pretendida</div>', unsafe_allow_html=True)
funcao = st.selectbox("1. Selecione a função para inscrição *", ["— Selecione —"] + FUNCOES)
is_dent = e_dentista(funcao)

# ----------------------------------------------------------------------------
# 2. DADOS PESSOAIS
# ----------------------------------------------------------------------------
st.markdown('<div class="ig2p-faixa">2. Dados Pessoais</div>', unsafe_allow_html=True)
nome = st.text_input("2. Nome Completo (sem abreviações) *")
cpf = st.text_input("3. CPF — digite apenas os números *",
                    key="cpf", max_chars=14, on_change=fmt_cpf,
                    help="A pontuação (000.000.000-00) é aplicada automaticamente.")
nascimento = st.text_input("4. Data de Nascimento (dia/mês/ano) — Ex: 01/01/1990 *",
                           key="nascimento", max_chars=10, on_change=fmt_nascimento,
                           help="Digite só os números — a barra é inserida automaticamente.")
email = st.text_input("5. E-mail *",
                      help="Será usado para o envio da confirmação e das comunicações do processo.")
telefone = st.text_input("6. Telefone para contato / WhatsApp (com DDD) *",
                         key="telefone", max_chars=16, on_change=fmt_tel)
endereco = st.text_area("7. Endereço residencial completo *", height=90,
                        help="Logradouro, número, complemento, bairro, município, estado e CEP.")

# ----------------------------------------------------------------------------
# 3. FORMAÇÃO E DOCUMENTAÇÃO
# ----------------------------------------------------------------------------
st.markdown('<div class="ig2p-faixa">3. Formação e Documentação</div>', unsafe_allow_html=True)
st.markdown(f"""
<div class="ig2p-legal">
Anexe os documentos nos formatos <b>PDF, JPG, JPEG ou PNG</b> (até {TAMANHO_MAX_MB} MB cada).
Os documentos de <b>formação/habilitação</b> e de <b>regularidade no CRO</b> são obrigatórios;
os títulos acadêmicos e a experiência profissional são opcionais e servem para a pontuação
curricular, conforme o Anexo V do Edital.
</div>
""", unsafe_allow_html=True)

st.markdown("**Registro profissional no Conselho Regional de Odontologia (CRO)**")
c1, c2 = st.columns([3, 1])
with c1:
    cro_num = st.text_input("8. Número do registro no CRO *")
with c2:
    cro_uf = st.text_input("UF *", key="cro_uf", max_chars=2, on_change=fmt_cro_uf,
                           help="Sigla do estado, ex.: MG")

arquivos = {}
n = 9
for chave, rotulo, obrig, so_dent in ANEXOS:
    if so_dent and not is_dent:
        arquivos[chave] = None
        continue
    marca = " *" if obrig else " (opcional)"
    arquivos[chave] = st.file_uploader(f"{n}. {rotulo}{marca}",
                                       type=EXTS_ANEXO, key=f"file_{chave}")
    n += 1

# ----------------------------------------------------------------------------
# 4. DECLARAÇÕES
# ----------------------------------------------------------------------------
st.markdown('<div class="ig2p-faixa">4. Declarações</div>', unsafe_allow_html=True)
dec_edital = st.checkbox(
    "Declaro que li e estou ciente das disposições estabelecidas no Edital do Processo Seletivo "
    "Público Simplificado da Secretaria Municipal de Saúde de Cássia/MG, comprometendo-me a cumprir "
    "integralmente suas normas, condições e exigências. *")
dec_verdade = st.checkbox(
    "Declaro, sob as penas da lei, que todas as informações prestadas e os documentos apresentados "
    "no ato da inscrição são verdadeiros e autênticos, estando ciente de que a constatação de "
    "informações falsas ou irregularidades poderá resultar no indeferimento da inscrição ou na "
    "eliminação do processo seletivo. *")
dec_lgpd = st.checkbox(
    "Declaro estar ciente de que os dados pessoais e documentos fornecidos serão tratados pela "
    "Prefeitura Municipal de Cássia/MG para fins de inscrição, análise documental, classificação, "
    "divulgação de resultados e demais procedimentos administrativos do Processo Seletivo. *")

# ----------------------------------------------------------------------------
# ENVIO + VALIDAÇÃO
# ----------------------------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
if st.button("Enviar inscrição", disabled=st.session_state.processando):
    erros = []
    if funcao == "— Selecione —":
        erros.append("Selecione a função pretendida.")
    if not nome.strip() or len(nome.split()) < 2:
        erros.append("Informe o nome completo, sem abreviações.")
    if not cpf_valido(cpf):
        erros.append("CPF inválido — informe os 11 números.")
    if not data_valida(nascimento):
        erros.append("Data de nascimento inválida — use o formato dia/mês/ano (ex: 01/01/1990).")
    if not email_valido(email):
        erros.append("Informe um e-mail válido.")
    if not telefone_valido(telefone):
        erros.append("Telefone inválido — informe DDD + número.")
    if not endereco.strip():
        erros.append("Informe o endereço residencial completo.")
    if not cro_num.strip():
        erros.append("Informe o número do registro no CRO.")
    if len(cro_uf.strip()) != 2:
        erros.append("Informe a UF do registro no CRO (2 letras).")
    for chave, rotulo, obrig, so_dent in ANEXOS:
        if so_dent and not is_dent:
            continue
        arq = arquivos.get(chave)
        if arq is None:
            if obrig:
                erros.append(f"Anexe: {rotulo}.")
        elif len(arq.getvalue()) > TAMANHO_MAX_MB * 1024 * 1024:
            erros.append(f"O arquivo '{rotulo}' excede {TAMANHO_MAX_MB} MB.")
    if not dec_edital:
        erros.append("É necessário aceitar a declaração de ciência e concordância com o edital.")
    if not dec_verdade:
        erros.append("É necessário aceitar a declaração de veracidade das informações.")
    if not dec_lgpd:
        erros.append("É necessário aceitar a declaração de ciência sobre o tratamento de dados pessoais.")

    if erros:
        st.markdown(
            f'<div class="ig2p-card" style="border-color:{T["erro"]};">'
            f'<b style="color:{T["erro"]};">Corrija os campos abaixo:</b><ul>'
            + "".join(f"<li style=\"color:{T['texto2']};\">{e}</li>" for e in erros)
            + "</ul></div>", unsafe_allow_html=True)
    else:
        # Anti-duplicidade por CPF: se já existe inscrição, não cria outra nem sobe anexos.
        ja_protocolo = buscar_inscricao_por_cpf(cpf)
        if ja_protocolo is not None:
            st.session_state.protocolo = ja_protocolo
            st.session_state.duplicado = True
            st.session_state.email_ok = None
            st.session_state.enviado = True
            st.rerun()

        token = uuid.uuid4().hex
        # Uploads AGORA (o file_uploader não sobrevive ao st.rerun()).
        links = {}
        with st.spinner("Enviando documentos..."):
            for chave, rotulo, obrig, so_dent in ANEXOS:
                if so_dent and not is_dent:
                    links[chave] = ""
                    continue
                link, _dest, _mot = salvar_anexo(arquivos.get(chave), token, nome.strip(), chave)
                links[chave] = link

        st.session_state.dados_pendentes = {
            "linha": [
                agora().strftime("%d/%m/%Y %H:%M:%S"), "",  # protocolo definido ao gravar
                funcao, nome.strip(), cpf.strip(), nascimento.strip(),
                email.strip(), telefone.strip(), endereco.strip(),
                cro_num.strip(), cro_uf.strip().upper(),
                links.get("diploma", ""), links.get("cro_cert", ""),
                links.get("especializacao", ""), links.get("mestrado", ""),
                links.get("cursos", ""), links.get("experiencia", ""),
                "Sim" if dec_edital else "Não",
                "Sim" if dec_verdade else "Não",
                "Sim" if dec_lgpd else "Não",
            ],
            "email": email.strip(), "nome": nome.strip(), "funcao": funcao,
            "token": token,
        }
        st.session_state.processando = True
        st.rerun()

st.markdown(
    f'<div style="margin-top:26px;padding:16px 18px;border-radius:14px;'
    f'background:{T["card"]};border:1px solid {T["borda"]};box-shadow:{T["sombra"]};text-align:center;">'
    f'<p style="color:{T["texto"]};font-size:.88rem;margin:0;font-weight:600;">'
    f'Está com dúvidas ou problemas para se inscrever?</p>'
    f'<p style="color:{T["texto2"]};font-size:.83rem;margin:5px 0 0;">Entre em contato pelo e-mail: '
    f'<a href="mailto:smscassiamg@gmail.com" style="color:{T["primaria"]};font-weight:700;'
    f'text-decoration:none;">smscassiamg@gmail.com</a></p></div>', unsafe_allow_html=True)

st.markdown('<div class="ig2p-rodape">iG2P · Inteligência em Gestão Pública · G3ST4O · Prefeitura de Cássia - MG</div>',
            unsafe_allow_html=True)