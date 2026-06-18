"""
Download de NFSe/ XML/ Boleto — Front Streamlit (Sillion)
Encaminha período (data inicial/final) + email + CNPJ (opcional) para o backend
N8N via POST JSON. O backend retorna o relatório de extração por e-mail.

Arquitetura:
- app.py        → lógica Python (config, envio, widgets de input)
- styles/       → CSS (visual)
- templates/    → HTML estrutural (header, hero, footer, etc.)
"""

import re
from datetime import datetime, date
from pathlib import Path

import requests
import streamlit as st

import separar_detalhamento as sd

# ============================================================
# Caminhos
# ============================================================
BASE_DIR = Path(__file__).parent
CSS_PATH = BASE_DIR / "styles" / "main.css"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

# ============================================================
# Recursos externos
# ============================================================
LOGO_EXTERNO = "https://www.sillion.com.br/wp-content/themes/sillion/images/logo-black-tm.svg"
LOGO_LOCAL_FILE = STATIC_DIR / "logo-sillion.svg"


def resolver_logo_url() -> str:
    """
    Retorna o caminho do logo:
    - Se houver `static/logo-sillion.svg`, usa a versão local (mais rápida e offline).
    - Caso contrário, cai para a URL externa do site da Sillion.
    Streamlit sanitiza o atributo `onerror` em HTML, então o fallback
    precisa ser feito no Python, não no navegador.
    """
    if LOGO_LOCAL_FILE.exists():
        return "app/static/logo-sillion.svg"
    return LOGO_EXTERNO

# ============================================================
# Config da página
# ============================================================
st.set_page_config(
    page_title="Sillion · Download de NFSe/ XML/ Boleto",
    page_icon="",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ============================================================
# Constantes
# ============================================================
# Apenas emails @sillion.com.br são aceitos (case-insensitive)
DOMINIO_PERMITIDO = "sillion.com.br"
EMAIL_REGEX = re.compile(
    rf"^[A-Za-z0-9._%+\-]+@{re.escape(DOMINIO_PERMITIDO)}$",
    re.IGNORECASE,
)

TIMEOUT_REQ = 120  # segundos

# Opções de empresa/fonte a serem extraídas
EMPRESAS_OPCOES = ["TOT", "VALE"]

# Empresa solicitante (quem esta acionando o webhook)
EMPRESA_SOLICITANTE_OPCOES = ["Sitrack", "Sillion"]

# String identificadora enviada no HTTP quando a opção de Detalhamento
# da Medição é selecionada (caso de cliente específico).
OPCAO_DETALHAMENTO = "enviar_email"


# ============================================================
# Helpers de renderização (templates + CSS)
# ============================================================
def render_template(nome: str, **variaveis) -> str:
    """
    Lê um arquivo .html em templates/ e substitui placeholders no
    formato {{nome_da_variavel}} pelos valores passados.
    """
    caminho = TEMPLATES_DIR / f"{nome}.html"
    html = caminho.read_text(encoding="utf-8")
    for chave, valor in variaveis.items():
        html = html.replace(f"{{{{{chave}}}}}", str(valor))
    return html


def inject(html: str) -> None:
    """Injeta um trecho HTML na página."""
    st.markdown(html, unsafe_allow_html=True)


def carregar_css(caminho: Path) -> None:
    """Lê o arquivo CSS e injeta na página via st.markdown."""
    try:
        css = caminho.read_text(encoding="utf-8")
        inject(f"<style>{css}</style>")
    except FileNotFoundError:
        st.warning(f"Arquivo de estilos não encontrado: {caminho}")


# Carrega meta tags + CSS antes de qualquer conteúdo
inject(render_template("meta"))
carregar_css(CSS_PATH)


# ============================================================
# Configuração segura: URL do webhook
# ============================================================
try:
    WEBHOOK_URL = st.secrets["N8N_WEBHOOK_URL"]
except (KeyError, FileNotFoundError):
    WEBHOOK_URL = None

# Webhook do LOTE de relatórios de medição.
# Se não houver chave própria, reaproveita o webhook principal.
try:
    WEBHOOK_LOTE_URL = st.secrets["N8N_WEBHOOK_LOTE_URL"]
except (KeyError, FileNotFoundError):
    WEBHOOK_LOTE_URL = WEBHOOK_URL


# ============================================================
# Helpers de negócio
# ============================================================
def email_valido(email: str) -> bool:
    return bool(EMAIL_REGEX.match(email.strip()))


def somente_digitos(valor: str) -> str:
    """Remove qualquer caractere que não seja dígito (útil para CNPJ)."""
    return re.sub(r"\D", "", valor or "")


def formatar_cnpj(valor: str) -> str:
    """
    Normaliza o CNPJ para o formato XX.XXX.XXX/YYYY-ZZ.
    Retorna string vazia se o valor estiver vazio.
    Retorna apenas os dígitos se não tiver exatamente 14 dígitos
    (a validação a montante já barra esse caso).
    """
    d = somente_digitos(valor)
    if not d:
        return ""
    if len(d) != 14:
        return d
    return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"


def montar_payload(
    email: str,
    empresa_solicitante: str,
    data_inicial: date,
    data_final: date,
    empresas: list,
    cnpj: str = "",
    opcao: str = "",
) -> dict:
    """
    Payload enviado ao backend N8N.
    - empresa_solicitante: "Sitrack" ou "Sillion" (obrigatório)
    - data_inicial / data_final: ISO 8601 (YYYY-MM-DD)
    - empresas: lista com uma ou mais entre EMPRESAS_OPCOES (ex.: ["TOT", "VALE"])
    - cnpj: normalizado para XX.XXX.XXX/YYYY-ZZ ou "" se não informado
    """
    return {
        "email": email.strip(),
        "empresa": empresa_solicitante,
        "data_inicial": data_inicial.isoformat(),
        "data_final": data_final.isoformat(),
        "empresas": list(empresas),
        "cnpj": formatar_cnpj(cnpj),
        "opcao": opcao,
    }


def enviar_para_n8n(url: str, payload: dict) -> requests.Response:
    return requests.post(
        url,
        json=payload,
        timeout=TIMEOUT_REQ,
        headers={"Content-Type": "application/json"},
    )


# ============================================================
# UI — Header + Hero (vindos dos templates HTML)
# ============================================================
inject(render_template("header", logo_url=resolver_logo_url()))
inject(render_template(
    "hero",
    titulo="Download de NFSe/ XML/ Boleto",
    subtitulo="Extração de arquivos PDF e XML de acordo com o período selecionado.",
))


# ============================================================
# Verificação de configuração
# ============================================================
if not WEBHOOK_URL:
    st.error(
        "⚠️ A URL do webhook N8N não foi configurada. "
        "Crie o arquivo `.streamlit/secrets.toml` com a chave `N8N_WEBHOOK_URL` "
        "ou configure-a no painel do Streamlit Community Cloud."
    )
    st.stop()


# ============================================================
# UI — Formulário (widgets Streamlit — precisam falar com Python)
# ============================================================
email = st.text_input(
    "Email corporativo",
    placeholder=f"usuario@{DOMINIO_PERMITIDO}",
    help=f"Apenas emails do domínio @{DOMINIO_PERMITIDO} são aceitos. "
         "O relatório processado será enviado para este endereço.",
)

empresa_solicitante = st.radio(
    "Selecione a empresa:",
    options=EMPRESA_SOLICITANTE_OPCOES,
    index=None,
    horizontal=True,
    help="Empresa que está acionando esta extração (campo obrigatório).",
)

cnpj = st.text_input(
    "CNPJ do cliente (opcional)",
    placeholder="00.000.000/0000-00",
    help="Filtre a extração por um CNPJ específico. "
         "Deixe em branco para considerar todos os CNPJs.",
)

st.markdown("**Período selecionado:**")
col_ini, col_fim = st.columns(2)
with col_ini:
    data_inicial = st.date_input(
        "Data inicial",
        value=None,
        format="DD/MM/YYYY",
    )
with col_fim:
    data_final = st.date_input(
        "Data final",
        value=None,
        format="DD/MM/YYYY",
    )

empresas = st.multiselect(
    "Tipo de faturamento",
    options=EMPRESAS_OPCOES,
    default=[],
    placeholder="Selecione TOT, VALE ou ambas",
    help="Selecione uma ou as duas opções. Pelo menos uma é obrigatória.",
)

st.write("")
detalhar = st.toggle(
    "Habilitar envio por email",
    value=False,
    help="Ative para subir o Banco de Dados Sillion (aba 'Detalhamentos PDF'). "
         "O painel de upload aparece no fim da página e o identificador "
         f"'{OPCAO_DETALHAMENTO}' é enviado no payload (Download e lote).",
)

st.write("")
# Com o envio por email ativo (lote de detalhamento), o Download não se aplica — some.
if detalhar:
    enviar = False
else:
    enviar = st.button("Download", type="primary", use_container_width=True)


# ============================================================
# Lógica de envio
# ============================================================
if enviar:
    erros = []

    if not email.strip():
        erros.append("Informe o email.")
    elif not email_valido(email):
        erros.append(
            f"Email inválido. Use um endereço corporativo @{DOMINIO_PERMITIDO} "
            "(ex: seu.nome@" + DOMINIO_PERMITIDO + ")."
        )

    if not empresa_solicitante:
        erros.append("Selecione a empresa (Sitrack ou Sillion).")

    if data_inicial is None:
        erros.append("Informe a data inicial.")
    if data_final is None:
        erros.append("Informe a data final.")
    if data_inicial and data_final and data_inicial > data_final:
        erros.append("A data inicial não pode ser posterior à data final.")

    if not empresas:
        erros.append("Selecione pelo menos uma empresa (TOT e/ou VALE).")

    # CNPJ é opcional — só valida se foi preenchido
    cnpj_digitos = somente_digitos(cnpj)
    if cnpj.strip() and len(cnpj_digitos) != 14:
        erros.append("CNPJ inválido. Deve conter 14 dígitos.")

    if erros:
        for e in erros:
            st.error(e)
    else:
        with st.spinner("Solicitando extração ao backend..."):
            try:
                opcao = OPCAO_DETALHAMENTO if detalhar else ""
                payload = montar_payload(email, empresa_solicitante, data_inicial, data_final, empresas, cnpj, opcao)
                resp = enviar_para_n8n(WEBHOOK_URL, payload)

                if 200 <= resp.status_code < 300:
                    @st.dialog("Solicitação enviada")
                    def confirmacao():
                        st.success("Solicitação enviada com sucesso!")
                        st.write(
                            f"O relatório com os arquivos PDF e XML será encaminhado "
                            f"para **{email.strip()}** assim que o backend concluir "
                            "a extração."
                        )
                        st.caption(f"Empresa: {empresa_solicitante}")
                        st.caption(
                            "Período: "
                            f"{data_inicial.strftime('%d/%m/%Y')} a "
                            f"{data_final.strftime('%d/%m/%Y')}"
                        )
                        st.caption(f"Empresas: {', '.join(empresas)}")
                        if cnpj_digitos:
                            st.caption(f"CNPJ: {cnpj.strip()}")
                        st.caption(
                            f"Enviado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}"
                        )
                        if st.button("OK", use_container_width=True):
                            st.rerun()

                    confirmacao()
                else:
                    st.error(f"O backend respondeu com status {resp.status_code}.")
                    with st.expander("Detalhes da resposta"):
                        st.code(resp.text or "(sem corpo)")
            except requests.exceptions.Timeout:
                st.error("Tempo de resposta excedido. Verifique se o N8N está acessível.")
            except requests.exceptions.ConnectionError:
                st.error("Falha de conexão. Verifique a URL do webhook.")
            except Exception as exc:
                st.error(f"Erro inesperado: {exc}")


# ============================================================
# UI — Detalhamento da Medição (upload → gerar XLSX por cont_id → enviar em lote)
# ============================================================
st.markdown("---")
st.subheader("Detalhamento da Medição")


def _brl_app(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_data(d) -> str:
    try:
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(d) if d else "-"


@st.cache_data(show_spinner=False)
def _processar_arquivo(nome: str, conteudo: bytes) -> dict:
    """Gera os blocos .xlsx (cacheado por nome+conteúdo para não reprocessar a cada rerun)."""
    return sd.gerar_blocos(conteudo)


arquivo = None
if detalhar:
    st.caption(
        "Envie o **Banco de Dados Sillion** (aba **Faturamento**). "
        "O front processa **todos os cont_id**, gera **1 arquivo .xlsx por cont_id** "
        "e envia tudo em **um único lote** ao N8N (cada arquivo separado, em Base64)."
    )
    arquivo = st.file_uploader(
        "Banco de Dados Sillion",
        type=["xlsx"],
        help="Planilha com a aba 'Faturamento' (período em F2/G2, cabeçalho na linha 5, dados a partir da 6).",
    )
else:
    st.info(
        "Ative **Habilitar envio por email** no formulário acima para "
        "subir o Banco de Dados Sillion e enviar o lote de relatórios."
    )

if arquivo is not None:
    try:
        with st.spinner("Lendo a planilha e gerando os relatórios..."):
            resultado = _processar_arquivo(arquivo.name, arquivo.getvalue())

        periodo = resultado["periodo"]
        relatorios = resultado["blocos"]

        if not relatorios:
            st.warning(
                "Nenhum cont_id encontrado. Confirme se o arquivo contém a aba "
                "'Faturamento' com os dados a partir da linha 6."
            )
        else:
            total_geral = sum(r["total"] for r in relatorios)
            st.markdown(
                f"**Período:** {_fmt_data(periodo[0])} a "
                f"{_fmt_data(periodo[1])} · "
                f"**{len(relatorios)} relatórios** · "
                f"**Total geral:** R$ {_brl_app(total_geral)}"
            )
            st.dataframe(
                [
                    {
                        "Cont ID": r["cont_id"],
                        "Cliente": r["cliente_nome"],
                        "Veículos": r["qtd_veiculos"],
                        "Total (R$)": _brl_app(r["total"]),
                        "Arquivo": r["filename"],
                    }
                    for r in relatorios
                ],
                use_container_width=True,
                hide_index=True,
            )

            enviar_lote = st.button(
                "Gerar e enviar lote ao N8N",
                type="primary",
                use_container_width=True,
                key="btn_enviar_lote",
            )

            if enviar_lote:
                erros_lote = []
                if not email.strip():
                    erros_lote.append("Informe o email corporativo no topo do formulário.")
                elif not email_valido(email):
                    erros_lote.append(f"Email inválido. Use um endereço @{DOMINIO_PERMITIDO}.")
                if not empresa_solicitante:
                    erros_lote.append("Selecione a empresa (Sitrack ou Sillion) no topo.")
                if not empresas:
                    erros_lote.append("Selecione o tipo de faturamento (TOT e/ou VALE) no topo.")

                if erros_lote:
                    for e in erros_lote:
                        st.error(e)
                else:
                    with st.spinner(f"Enviando lote com {len(relatorios)} relatórios ao backend..."):
                        try:
                            payload = sd.build_lote_payload(
                                relatorios, periodo, email, empresa_solicitante,
                                OPCAO_DETALHAMENTO, empresas,
                            )
                            resp = enviar_para_n8n(WEBHOOK_LOTE_URL, payload)

                            if 200 <= resp.status_code < 300:
                                @st.dialog("Lote enviado")
                                def confirmacao_lote():
                                    st.success(
                                        f"Lote enviado com sucesso! "
                                        f"{len(relatorios)} relatórios encaminhados ao N8N."
                                    )
                                    st.caption(f"Empresa: {empresa_solicitante}")
                                    st.caption(
                                        "Período: "
                                        f"{_fmt_data(periodo[0])} a "
                                        f"{_fmt_data(periodo[1])}"
                                    )
                                    st.caption(f"Total geral: R$ {_brl_app(total_geral)}")
                                    st.caption(
                                        f"Enviado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}"
                                    )
                                    if st.button("OK", use_container_width=True):
                                        st.rerun()

                                confirmacao_lote()
                            else:
                                st.error(f"O backend respondeu com status {resp.status_code}.")
                                with st.expander("Detalhes da resposta"):
                                    st.code(resp.text or "(sem corpo)")
                        except requests.exceptions.Timeout:
                            st.error("Tempo de resposta excedido. Verifique se o N8N está acessível.")
                        except requests.exceptions.ConnectionError:
                            st.error("Falha de conexão. Verifique a URL do webhook.")
                        except Exception as exc:
                            st.error(f"Erro inesperado ao enviar o lote: {exc}")
    except Exception as exc:
        st.error(f"Não foi possível processar o arquivo: {exc}")


# ============================================================
# UI — Footer (vindo do template HTML)
# ============================================================
inject(render_template("footer", ano=datetime.now().year))
