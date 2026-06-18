"""
Geração dos relatórios de medição (Detalhamento) — camada de tratamento no front.

Fluxo:
1. O usuário faz upload de um arquivo .xlsx já tratado, contendo a aba
   "Detalhamentos PDF" (ou equivalente).
2. Este módulo lê o período do cabeçalho, agrupa as linhas por CONT ID
   e gera 1 PDF por cont_id no layout da "PLANILHA DE MEDIÇÃO".
3. Cada PDF é devolvido em bytes (pronto para ser codificado em Base64 e
   enviado em lote ao N8N — ver build_lote_payload).

Não depende do Streamlit: pode ser testado isoladamente.
"""

from __future__ import annotations

import base64
import io
import re
import unicodedata
from datetime import datetime, date

import openpyxl
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

# ============================================================
# Constantes de layout / colunas do modelo "A Geradora"
# ============================================================
COLUNAS_MODELO = [
    "CONT ID", "CLIENTE ID", "CLIENTE NOME", "NOME", "PLACA",
    "OPERAÇÃO DO MÊS", "DATA DA OPERAÇÃO", "DIAS / QTD.",
    "MENSALIDADE", "PROPORCIONAL MENSAL", "VALOR OPERAÇÃO", "OBSERVAÇÕES",
]
# Índices (0-based) das colunas usadas, na ordem do modelo.
IDX = {name: i for i, name in enumerate(COLUNAS_MODELO)}

AZUL = colors.HexColor("#1a3a5c")
AZUL_CLARO = colors.HexColor("#eef2f6")
ZEBRA = colors.HexColor("#f4f7fa")
TOTAL_BG = colors.HexColor("#dde6ef")

MIME_PDF = "application/pdf"


# ============================================================
# Helpers
# ============================================================
def _num(x) -> float:
    """Converte célula para float; texto/vazio/#N/A viram 0."""
    if x is None:
        return 0.0
    try:
        return float(str(x).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _brl(v: float) -> str:
    """Formata número no padrão brasileiro (sem o símbolo R$)."""
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fdate(d) -> str:
    if isinstance(d, (datetime, date)):
        return d.strftime("%d/%m/%Y")
    if d in (None, "None", "-", ""):
        return "-"
    return str(d)


def _periodo_tag(di) -> str:
    """Sufixo AAAA-MM para o nome do arquivo."""
    if isinstance(di, (datetime, date)):
        return di.strftime("%Y-%m")
    return datetime.now().strftime("%Y-%m")


def _slug(texto: str) -> str:
    """Sanitiza um nome para uso em filename."""
    if not texto:
        return "cliente"
    txt = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    txt = re.sub(r"[^A-Za-z0-9]+", "_", txt).strip("_")
    return (txt[:40] or "cliente").upper()


def _is_subtotal(valor) -> bool:
    """Linhas de subtotal do Excel: '576 Total', 'Total Geral', etc."""
    return bool(re.search(r"total", str(valor), re.IGNORECASE))


# ============================================================
# Leitura da planilha
# ============================================================
def _achar_aba(wb) -> "openpyxl.worksheet.worksheet.Worksheet":
    """Procura a aba 'Detalhamentos PDF'; se não achar, usa a primeira."""
    for ws in wb.worksheets:
        if "detalhamento" in ws.title.strip().lower():
            return ws
    return wb.worksheets[0]


def _ler_periodo(linhas) -> dict:
    """
    Localiza DATA INICIAL / DATA FINAL / QTD. DIAS procurando os rótulos
    nas primeiras linhas e lendo a célula imediatamente abaixo.
    Robusto a pequenas mudanças de coluna.
    """
    di = df = dias = None
    n = min(8, len(linhas))
    for r in range(n - 1):
        for c, cell in enumerate(linhas[r]):
            rotulo = str(cell).strip().upper() if cell is not None else ""
            abaixo = linhas[r + 1][c] if c < len(linhas[r + 1]) else None
            if rotulo == "DATA INICIAL":
                di = abaixo
            elif rotulo == "DATA FINAL":
                df = abaixo
            elif rotulo in ("QTD. DIAS", "QTD DIAS"):
                dias = abaixo
    return {"data_inicial": di, "data_final": df, "qtd_dias": dias}


def _achar_header(linhas) -> int:
    """Índice (0-based) da linha de cabeçalho (a que tem 'CONT ID')."""
    for i, row in enumerate(linhas[:12]):
        if row and str(row[0]).strip().upper() == "CONT ID":
            return i
    # fallback: assume linha 5 (índice 4) como no modelo
    return 4


def ler_planilha(file_bytes: bytes) -> dict:
    """
    Lê o arquivo enviado e devolve:
    {
      "periodo": {data_inicial, data_final, qtd_dias},
      "grupos": { cont_id: [linhas...] }  # cada linha é uma tupla de células
    }
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = _achar_aba(wb)
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()

    periodo = _ler_periodo(linhas)
    h = _achar_header(linhas)
    dados = linhas[h + 1:]

    grupos: dict[str, list] = {}
    for row in dados:
        if not row:
            continue
        cont = row[0]
        if cont in (None, ""):
            continue
        if _is_subtotal(cont):
            continue
        grupos.setdefault(str(cont).strip(), []).append(row)
    return {"periodo": periodo, "grupos": grupos}


# ============================================================
# Geração do PDF (1 cont_id)
# ============================================================
def _gerar_pdf(cont_id: str, rows: list, periodo: dict) -> bytes:
    di = periodo.get("data_inicial")
    df = periodo.get("data_final")
    dias = periodo.get("qtd_dias")

    sum_prop = sum(_num(r[IDX["PROPORCIONAL MENSAL"]]) for r in rows)
    sum_op = sum(_num(r[IDX["VALOR OPERAÇÃO"]]) for r in rows)
    total = sum_prop + sum_op

    styles = getSampleStyleSheet()
    H = ParagraphStyle("H", parent=styles["Normal"], fontName="Helvetica-Bold",
                       fontSize=14, alignment=TA_CENTER, textColor=AZUL)
    cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=6.3, leading=7.6)
    cellc = ParagraphStyle("cellc", parent=cell, alignment=TA_CENTER)
    cellr = ParagraphStyle("cellr", parent=cell, alignment=TA_RIGHT)
    hcell = ParagraphStyle("hcell", parent=styles["Normal"], fontName="Helvetica-Bold",
                           fontSize=6.2, leading=7.4, textColor=colors.white,
                           alignment=TA_CENTER)
    lbl = ParagraphStyle("lbl", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.5)
    lblr = ParagraphStyle("lblr", parent=lbl, alignment=TA_RIGHT)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=8 * mm, rightMargin=8 * mm, topMargin=8 * mm, bottomMargin=10 * mm,
    )
    story = [Paragraph("PLANILHA DE MEDIÇÃO", H), Spacer(1, 5)]

    # --- Bloco período + total (estilo modelo) ---
    hdr = [[
        Paragraph("<b>PERÍODO</b>", lbl),
        Paragraph(f"<b>DATA INICIAL:</b> {_fdate(di)}", cell),
        Paragraph(f"<b>DATA FINAL:</b> {_fdate(df)}", cell),
        Paragraph(f"<b>QTD. DIAS:</b> {dias if dias is not None else '-'}", cell),
        Paragraph("", cell),
        Paragraph("<b>TOTAL (R$)</b>", lblr),
        Paragraph(f"<b>{_brl(sum_prop)}</b>", cellr),
        Paragraph(f"<b>{_brl(sum_op)}</b>", cellr),
        Paragraph(f"<b>TOTAL GERAL: {_brl(total)}</b>", cellr),
    ]]
    ht = Table(hdr, colWidths=[20 * mm, 40 * mm, 40 * mm, 25 * mm, 40 * mm,
                               28 * mm, 26 * mm, 24 * mm, 38 * mm])
    ht.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, AZUL),
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_CLARO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story += [ht, Spacer(1, 5)]

    # --- Tabela com as 12 colunas do modelo ---
    tdata = [[Paragraph(h, hcell) for h in COLUNAS_MODELO]]
    for r in rows:
        def g(name):
            i = IDX[name]
            return r[i] if i < len(r) else None
        tdata.append([
            Paragraph(str(g("CONT ID")), cellc),
            Paragraph(str(g("CLIENTE ID")) if g("CLIENTE ID") else "", cellc),
            Paragraph(str(g("CLIENTE NOME")) if g("CLIENTE NOME") not in (None, "None") else "", cell),
            Paragraph(str(g("NOME")) if g("NOME") not in (None, "None") else "", cell),
            Paragraph(str(g("PLACA")) if g("PLACA") else "", cellc),
            Paragraph(str(g("OPERAÇÃO DO MÊS")) if g("OPERAÇÃO DO MÊS") else "", cellc),
            Paragraph(_fdate(g("DATA DA OPERAÇÃO")), cellc),
            Paragraph(str(g("DIAS / QTD.")) if g("DIAS / QTD.") is not None else "", cellc),
            Paragraph(_brl(_num(g("MENSALIDADE"))) if g("MENSALIDADE") not in (None, "") else "", cellr),
            Paragraph(_brl(_num(g("PROPORCIONAL MENSAL"))) if g("PROPORCIONAL MENSAL") not in (None, "") else "", cellr),
            Paragraph(_brl(_num(g("VALOR OPERAÇÃO"))) if _num(g("VALOR OPERAÇÃO")) != 0 else "", cellr),
            Paragraph(str(g("OBSERVAÇÕES")) if g("OBSERVAÇÕES") not in (None, "None") else "", cell),
        ])
    # Linha de total (igual ao modelo: TOTAL (R$) sob MENSALIDADE,
    # somas sob PROPORCIONAL e VALOR OPERAÇÃO, TOTAL GERAL sob OBSERVAÇÕES)
    tdata.append(
        [Paragraph("", cell)] * 8 + [
            Paragraph("<b>TOTAL (R$)</b>", lblr),
            Paragraph(f"<b>{_brl(sum_prop)}</b>", cellr),
            Paragraph(f"<b>{_brl(sum_op)}</b>", cellr),
            Paragraph(f"<b>{_brl(total)}</b>", cellr),
        ]
    )
    cw = [15 * mm, 17 * mm, 46 * mm, 30 * mm, 18 * mm, 28 * mm,
          22 * mm, 15 * mm, 22 * mm, 26 * mm, 22 * mm, 20 * mm]
    t = Table(tdata, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbbbbb")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, ZEBRA]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("BACKGROUND", (0, -1), (-1, -1), TOTAL_BG),
        ("SPAN", (0, -1), (7, -1)),
        ("LINEABOVE", (0, -1), (-1, -1), 0.7, AZUL),
    ]))
    story.append(t)
    doc.build(story)
    return buf.getvalue()


# ============================================================
# API pública
# ============================================================
def gerar_relatorios(file_bytes: bytes) -> dict:
    """
    Processa o arquivo enviado e gera todos os relatórios.

    Retorna:
    {
      "periodo": {data_inicial, data_final, qtd_dias},
      "relatorios": [
        {
          "cont_id", "cliente_id", "cliente_nome",
          "qtd_veiculos", "total", "soma_proporcional", "soma_operacao",
          "filename", "pdf_bytes"
        }, ...
      ]
    }
    """
    parsed = ler_planilha(file_bytes)
    periodo = parsed["periodo"]
    tag = _periodo_tag(periodo.get("data_inicial"))

    relatorios = []
    for cont_id, rows in parsed["grupos"].items():
        cli_id = rows[0][IDX["CLIENTE ID"]] if len(rows[0]) > 1 else ""
        cli_nome = rows[0][IDX["CLIENTE NOME"]] if len(rows[0]) > 2 else ""
        soma_prop = sum(_num(r[IDX["PROPORCIONAL MENSAL"]]) for r in rows)
        soma_op = sum(_num(r[IDX["VALOR OPERAÇÃO"]]) for r in rows)
        pdf = _gerar_pdf(cont_id, rows, periodo)
        relatorios.append({
            "cont_id": cont_id,
            "cliente_id": str(cli_id) if cli_id else "",
            "cliente_nome": str(cli_nome) if cli_nome else "",
            "qtd_veiculos": len(rows),
            "total": round(soma_prop + soma_op, 2),
            "soma_proporcional": round(soma_prop, 2),
            "soma_operacao": round(soma_op, 2),
            "filename": f"{cont_id}_{_slug(cli_nome)}_medicao_{tag}.pdf",
            "pdf_bytes": pdf,
        })

    relatorios.sort(key=lambda x: int(x["cont_id"]) if x["cont_id"].isdigit() else 0)
    return {"periodo": periodo, "relatorios": relatorios}


def build_lote_payload(
    relatorios: list,
    periodo: dict,
    email: str,
    empresa: str,
    opcao: str = "",
) -> dict:
    """
    Monta o payload do LOTE para o webhook N8N.
    Cada relatório fica separado dentro do array `lote`, já em Base64.
    `opcao` é a string identificadora da opção selecionada no front
    (ex.: "enviar_email"). Ver contrato em DETALHAMENTO_MEDICAO.md.
    """
    def iso(d):
        if isinstance(d, (datetime, date)):
            return d.strftime("%Y-%m-%d")
        return str(d) if d else None

    itens = []
    for r in relatorios:
        itens.append({
            "cont_id": r["cont_id"],
            "cliente_id": r["cliente_id"],
            "cliente_nome": r["cliente_nome"],
            "filename": r["filename"],
            "mime_type": MIME_PDF,
            "total": r["total"],
            "qtd_veiculos": r["qtd_veiculos"],
            "file_base64": base64.b64encode(r["pdf_bytes"]).decode("ascii"),
        })

    return {
        "email": email.strip(),
        "empresa": empresa,
        "tipo": "detalhamento_medicao",
        "opcao": opcao,
        "periodo": {
            "data_inicial": iso(periodo.get("data_inicial")),
            "data_final": iso(periodo.get("data_final")),
            "qtd_dias": periodo.get("qtd_dias"),
        },
        "qtd_relatorios": len(itens),
        "lote": itens,
    }
