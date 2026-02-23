import io
import re
import calendar
import datetime as dt
from dataclasses import dataclass
from typing import Optional, Set, Dict, List, Tuple

import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import holidays
from dateutil import parser as date_parser

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


# =========================================================
# TEMA (Claro / Escuro) - CSS
# =========================================================
CSS_LIGHT = """
<style>
:root { --bg:#f2f5f9; --text:#222; --card:#ffffff; --accent:#0078d7; --muted:#e6eef8; }
html, body, [data-testid="stAppViewContainer"] { background: var(--bg) !important; color: var(--text) !important; }
[data-testid="stHeader"] { background: rgba(0,0,0,0) !important; }
.block-container { padding-top: 2rem; }
.card {
    background: var(--card);
    border-radius: 14px;
    padding: 16px;
    border: 1px solid rgba(0,0,0,0.08);
}
.badge {
    display:inline-block; padding:6px 10px; border-radius:999px;
    background: var(--muted); color: var(--text); font-weight:600; font-size: 0.9rem;
}
.stButton>button { background: var(--accent) !important; border: none !important; }
.stButton>button:hover { filter: brightness(0.92); }
</style>
"""

CSS_DARK = """
<style>
:root { --bg:#121212; --text:#eee; --card:#1f1f1f; --accent:#2196f3; --muted:#2c2c2c; }
html, body, [data-testid="stAppViewContainer"] { background: var(--bg) !important; color: var(--text) !important; }
[data-testid="stHeader"] { background: rgba(0,0,0,0) !important; }
.block-container { padding-top: 2rem; }
.card {
    background: var(--card);
    border-radius: 14px;
    padding: 16px;
    border: 1px solid rgba(255,255,255,0.12);
}
.badge {
    display:inline-block; padding:6px 10px; border-radius:999px;
    background: var(--muted); color: var(--text); font-weight:600; font-size: 0.9rem;
}
.stButton>button { background: var(--accent) !important; border: none !important; }
.stButton>button:hover { filter: brightness(0.92); }
</style>
"""

def apply_theme(theme_name: str):
    st.markdown(CSS_DARK if theme_name == "Escuro" else CSS_LIGHT, unsafe_allow_html=True)


# =========================================================
# ABA 1 - Atualizador de escalas (Excel)
# =========================================================

PT_DOW = {
    "SEG": 0,
    "TER": 1,
    "QUA": 2,
    "QUI": 3,
    "SEX": 4,
    "SAB": 5,
    "SÁB": 5,
    "DOM": 6,
}

DOW_TOKENS_RE = re.compile(r"(SEG|TER|QUA|QUI|SEX|SAB|SÁB|DOM)", re.IGNORECASE)

@dataclass
class RowResult:
    ok: bool
    error: Optional[str] = None


def safe_parse_date(value) -> Optional[dt.date]:
    """Parse Excel cell that may be date/datetime or string."""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    s = str(value).strip()
    try:
        return date_parser.parse(s, dayfirst=True).date()
    except Exception:
        return None


def parse_schedule_days(text: str) -> Optional[Set[int]]:
    """
    Convert schedule description to set of working weekdays (0=SEG ... 6=DOM).
    Supported:
      - "SEG A SEX", "SEG A SAB", "TER A DOM" (wrap supported)
      - list of days: "SEG TER QUA"
      - "FOLGA TER e DOM" => work = all - {TER, DOM}
    """
    if text is None:
        return None
    s = str(text).upper()

    tokens = DOW_TOKENS_RE.findall(s)
    if not tokens:
        return None

    tok = [t.upper() for t in tokens]
    tok = ["SAB" if t == "SÁB" else t for t in tok]

    if "FOLGA" in s:
        off = {PT_DOW[t] for t in tok if t in PT_DOW}
        return set(range(7)) - off

    if " A " in s and len(tok) >= 2:
        a = PT_DOW.get(tok[0])
        b = PT_DOW.get(tok[1])
        if a is None or b is None:
            return None
        if a <= b:
            return set(range(a, b + 1))
        return set(list(range(a, 7)) + list(range(0, b + 1)))

    days = {PT_DOW[t] for t in tok if t in PT_DOW}
    return days if days else None


def month_bounds(year: int, month: int) -> Tuple[dt.date, dt.date]:
    last = calendar.monthrange(year, month)[1]
    return dt.date(year, month, 1), dt.date(year, month, last)


def count_workdays(
    start: dt.date,
    end: dt.date,
    working_days: Set[int],
    holiday_set: Set[dt.date],
) -> int:
    if start > end:
        return 0
    cnt = 0
    d = start
    one = dt.timedelta(days=1)
    while d <= end:
        if (d.weekday() in working_days) and (d not in holiday_set):
            cnt += 1
        d += one
    return cnt


def parse_sheet_month_year(sheet_name: str) -> Optional[Tuple[int, int]]:
    """Expect sheet like '01.2026' -> (2026, 1)."""
    m = re.match(r"^\s*(\d{1,2})\.(\d{4})\s*$", str(sheet_name))
    if not m:
        return None
    month = int(m.group(1))
    year = int(m.group(2))
    if not (1 <= month <= 12):
        return None
    return year, month


def find_header_row_and_map(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Find header row by scanning first ~50 rows and mapping header text -> column index.
    Returns (header_row, header_map).
    """
    for r in range(1, 51):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if not values:
            continue
        normalized = []
        for v in values:
            normalized.append("" if v is None else str(v).strip().upper())

        if "INÍCIO ESCALA NOVA" in normalized and "ESCALA NOVA" in normalized:
            m: Dict[str, int] = {}
            for idx, name in enumerate(normalized, start=1):
                if name:
                    m[name] = idx
            return r, m

    raise ValueError("Não encontrei a linha de cabeçalho (preciso de 'INÍCIO ESCALA NOVA' e 'ESCALA NOVA').")


def ensure_column(ws: Worksheet, header_row: int, header_map: Dict[str, int], header_name: str) -> int:
    """Ensure a column exists; if not, create it at the end."""
    key = header_name.strip().upper()
    if key in header_map:
        return header_map[key]

    new_col = ws.max_column + 1
    ws.cell(row=header_row, column=new_col).value = header_name
    header_map[key] = new_col
    return new_col


def build_brazil_holidays(years: Set[int]) -> Set[dt.date]:
    br = holidays.Brazil(years=years)
    return {d for d in br.keys()}


def parse_extra_holidays(text: str) -> Set[dt.date]:
    out: Set[dt.date] = set()
    if not text:
        return out
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        try:
            out.add(date_parser.parse(s, dayfirst=True).date())
        except Exception:
            pass
    return out


def process_workbook(
    file_bytes: bytes,
    sheet_names: Optional[List[str]],
    extra_holidays: Set[dt.date],
) -> Tuple[bytes, List[str]]:
    logs: List[str] = []
    bio = io.BytesIO(file_bytes)
    wb = load_workbook(bio)
    target_sheets = sheet_names or wb.sheetnames

    years: Set[int] = set()
    sheet_ym: Dict[str, Tuple[int, int]] = {}

    for sname in target_sheets:
        ym = parse_sheet_month_year(sname)
        if ym:
            y, m = ym
            years.add(y)
            sheet_ym[sname] = (y, m)

    for d in extra_holidays:
        years.add(d.year)

    br_holidays = build_brazil_holidays(years) if years else set()
    holiday_set = set(br_holidays) | set(extra_holidays)

    updated_any = False

    for sname in target_sheets:
        if sname not in wb.sheetnames:
            continue
        if sname not in sheet_ym:
            logs.append(f"[IGNORADO] Aba '{sname}' não está no formato MM.AAAA (ex: 01.2026).")
            continue

        year, month_sheet = sheet_ym[sname]
        ws = wb[sname]

        try:
            header_row, header_map = find_header_row_and_map(ws)
        except Exception as e:
            logs.append(f"[ERRO] Aba '{sname}': {e}")
            continue

        old_scale_header = f"ESCALA {month_sheet:02d}.{year}"
        old_key = old_scale_header.upper()

        if old_key not in header_map:
            logs.append(f"[ERRO] Aba '{sname}': não achei a coluna '{old_scale_header}'.")
            continue

        col_old = header_map[old_key]
        col_new = header_map["ESCALA NOVA"]
        col_start = header_map["INÍCIO ESCALA NOVA"]

        m1, y1 = month_sheet, year
        if m1 == 12:
            m2, y2 = 1, year + 1
        else:
            m2, y2 = m1 + 1, year

        h_old_1 = f"DIAS ÚTEIS {m1:02d}.{y1} (ESCALA ANTIGA)"
        h_new_1 = f"DIAS ÚTEIS {m1:02d}.{y1} (ESCALA NOVA)"
        h_due_1 = f"DIAS DEVIDOS {m1:02d}.{y1}"

        h_old_2 = f"DIAS ÚTEIS {m2:02d}.{y2} (ESCALA ANTIGA)"
        h_new_2 = f"DIAS ÚTEIS {m2:02d}.{y2} (ESCALA NOVA)"
        h_due_2 = f"DIAS DEVIDOS {m2:02d}.{y2}"

        h_total = "TOTAL DIAS DEVIDOS"

        c_old_1 = ensure_column(ws, header_row, header_map, h_old_1)
        c_new_1 = ensure_column(ws, header_row, header_map, h_new_1)
        c_due_1 = ensure_column(ws, header_row, header_map, h_due_1)

        c_old_2 = ensure_column(ws, header_row, header_map, h_old_2)
        c_new_2 = ensure_column(ws, header_row, header_map, h_new_2)
        c_due_2 = ensure_column(ws, header_row, header_map, h_due_2)

        c_total = ensure_column(ws, header_row, header_map, h_total)

        processed = 0
        errors = 0

        last_row = ws.max_row
        for r in range(header_row + 1, last_row + 1):
            v_old = ws.cell(row=r, column=col_old).value
            v_new = ws.cell(row=r, column=col_new).value
            v_start = ws.cell(row=r, column=col_start).value

            if v_old is None and v_new is None and v_start is None:
                continue

            start_date = safe_parse_date(v_start)
            if start_date is None:
                errors += 1
                continue

            days_old = parse_schedule_days(v_old)
            days_new = parse_schedule_days(v_new)
            if days_old is None or days_new is None:
                errors += 1
                continue

            m1_start, m1_end = month_bounds(y1, m1)
            if start_date > m1_end:
                old_1 = new_1 = due_1 = 0
            else:
                calc_start = max(start_date, m1_start)
                old_1 = count_workdays(calc_start, m1_end, days_old, holiday_set)
                new_1 = count_workdays(calc_start, m1_end, days_new, holiday_set)
                due_1 = old_1 - new_1

            m2_start, m2_end = month_bounds(y2, m2)
            if start_date > m2_end:
                old_2 = new_2 = due_2 = 0
            else:
                calc_start = max(start_date, m2_start)
                old_2 = count_workdays(calc_start, m2_end, days_old, holiday_set)
                new_2 = count_workdays(calc_start, m2_end, days_new, holiday_set)
                due_2 = old_2 - new_2

            total = due_1 + due_2

            ws.cell(row=r, column=c_old_1).value = old_1
            ws.cell(row=r, column=c_new_1).value = new_1
            ws.cell(row=r, column=c_due_1).value = due_1

            ws.cell(row=r, column=c_old_2).value = old_2
            ws.cell(row=r, column=c_new_2).value = new_2
            ws.cell(row=r, column=c_due_2).value = due_2

            ws.cell(row=r, column=c_total).value = total

            processed += 1

        logs.append(f"[OK] Aba '{sname}': {processed} linhas atualizadas, {errors} linhas com erro (data/escala inválida).")
        updated_any = True

    if not updated_any:
        logs.append("[AVISO] Nenhuma aba foi atualizada. Verifique se as abas estão no formato MM.AAAA (ex: 01.2026) e se os cabeçalhos existem.")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), logs


# =========================================================
# ABA 2 - WorkTime Manager (Streamlit) + Export PDF/Excel
# =========================================================

DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
MESES_PT = [
    "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril",
    "05 - Maio", "06 - Junho", "07 - Julho", "08 - Agosto",
    "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
]

PONTOS_FACULTATIVOS = ["Carnaval", "Cinzas", "Quaresma", "Servidor Público", "Véspera", "observado"]

def is_ponto_facultativo(nome_feriado: str) -> bool:
    return any(p.lower() in nome_feriado.lower() for p in PONTOS_FACULTATIVOS)

def calcular_dias_trabalhados(mes, folgas_semana, ano, data_ini=None, data_fim=None):
    total_dias = calendar.monthrange(ano, mes)[1]
    primeiro_dia = 1
    ultimo_dia = total_dias

    if data_ini and data_ini.month == mes and data_ini.year == ano:
        primeiro_dia = max(1, data_ini.day)
    if data_fim and data_fim.month == mes and data_fim.year == ano:
        ultimo_dia = min(total_dias, data_fim.day)

    feriados = holidays.Brazil(years=ano)
    feriados_do_mes = [
        d.day for d, nome in feriados.items()
        if d.month == mes and not is_ponto_facultativo(nome)
        and primeiro_dia <= d.day <= ultimo_dia
    ]

    folgas_count = sum(
        1 for dia in range(primeiro_dia, ultimo_dia + 1)
        if folgas_semana[dt.datetime(ano, mes, dia).weekday()]
    )

    feriado_em_dia_util = sum(
        1 for dia in feriados_do_mes
        if not folgas_semana[dt.datetime(ano, mes, dia).weekday()]
    )

    dias_trabalhados = (ultimo_dia - primeiro_dia + 1) - folgas_count - feriado_em_dia_util

    feriados_ano = holidays.Brazil(years=ano)
    detalhes = {d.day: nome for d, nome in feriados_ano.items()
                if d.month == mes and not is_ponto_facultativo(nome)}
    if data_fim and data_fim.month == mes and data_fim.year == ano:
        detalhes = {dia: nome for dia, nome in detalhes.items() if dia <= data_fim.day}

    return max(dias_trabalhados, 0), feriados_do_mes, detalhes, (primeiro_dia, ultimo_dia)

def make_pdf_worktime(payload: dict) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    x = 50
    y = h - 60

    c.setFont("Helvetica-Bold", 16)
    c.drawString(x, y, "WorkTime Manager - Relatório")
    y -= 28

    c.setFont("Helvetica", 11)
    c.drawString(x, y, f"Mês/Ano: {payload['mes_nome']} / {payload['ano']}")
    y -= 18
    c.drawString(x, y, f"Período: {payload['periodo_str']}")
    y -= 18
    c.drawString(x, y, f"Dias trabalhados: {payload['dias_trabalhados']}")
    y -= 18
    c.drawString(x, y, f"Folgas semanais: {payload['folgas_str']}")
    y -= 24

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "Feriados considerados (não ponto facultativo):")
    y -= 18

    c.setFont("Helvetica", 10)
    lines = payload["feriados_str"].split("\n") if payload["feriados_str"] else ["Nenhum"]
    for line in lines:
        if y < 60:
            c.showPage()
            y = h - 60
            c.setFont("Helvetica", 10)
        c.drawString(x, y, line)
        y -= 14

    c.showPage()
    c.save()
    return buf.getvalue()

def make_excel_worktime(payload: dict) -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"

    ws["A1"] = "WorkTime Manager - Resultado"
    ws["A3"] = "Mês/Ano"
    ws["B3"] = f"{payload['mes_nome']} / {payload['ano']}"

    ws["A4"] = "Período"
    ws["B4"] = payload["periodo_str"]

    ws["A5"] = "Dias trabalhados"
    ws["B5"] = payload["dias_trabalhados"]

    ws["A6"] = "Folgas semanais"
    ws["B6"] = payload["folgas_str"]

    ws["A8"] = "Feriados considerados"
    r = 9
    if payload["feriados_str"]:
        for line in payload["feriados_str"].split("\n"):
            ws[f"A{r}"] = line
            r += 1
    else:
        ws[f"A{r}"] = "Nenhum"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================================================
# UI - Streamlit
# =========================================================
st.set_page_config(page_title="Sistema de Escalas", layout="centered")

if "tema" not in st.session_state:
    st.session_state["tema"] = "Claro"

st.sidebar.title("⚙️ Configurações")
st.session_state["tema"] = st.sidebar.radio(
    "Tema",
    ["Claro", "Escuro"],
    index=0 if st.session_state["tema"] == "Claro" else 1
)
apply_theme(st.session_state["tema"])

st.title("Sistema de Escalas")
st.caption("Aba 1: Atualiza planilha de escala | Aba 2: WorkTime Manager (dias trabalhados + exportação PDF/Excel)")

aba1, aba2 = st.tabs(["Atualizar Escala (Excel)", "WorkTime Manager"])


# ----------------------------
# ABA 1
# ----------------------------
with aba1:
    st.subheader("Atualizar planilha de alteração de escala")
    st.caption("Faça upload do .xlsx, clique em **ATUALIZAR PLANILHA** e baixe o arquivo atualizado (mesma formatação, só valores).")

    uploaded = st.file_uploader("Envie sua planilha (.xlsx)", type=["xlsx"], key="uploader_excel")

    extra_holidays_text = st.text_area(
        "Feriados extras (opcional) — 1 por linha (ex: 25/01/2026)",
        value="",
        placeholder="Ex:\n25/01/2026\n16/07/2026",
        height=120,
        key="extras_feriados_excel"
    )

    process_all_sheets = st.checkbox("Processar todas as abas no formato MM.AAAA", value=True, key="process_all")

    selected_sheets = None
    sheetnames = None

    if uploaded is not None:
        try:
            wb_preview = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
            sheetnames = wb_preview.sheetnames
            wb_preview.close()
        except Exception as e:
            st.error(f"Não consegui abrir esse arquivo: {e}")
            st.stop()

        if not process_all_sheets and sheetnames:
            selected_sheets = st.multiselect(
                "Selecione as abas para processar",
                options=sheetnames,
                default=[s for s in sheetnames if re.match(r"^\d{1,2}\.\d{4}$", s.strip())],
                key="sheets_select"
            )

    btn = st.button("ATUALIZAR PLANILHA", type="primary", disabled=(uploaded is None), key="btn_update")

    if btn and uploaded is not None:
        with st.spinner("Processando..."):
            extras = parse_extra_holidays(extra_holidays_text)
            updated_bytes, logs = process_workbook(
                file_bytes=uploaded.getvalue(),
                sheet_names=(None if process_all_sheets else selected_sheets),
                extra_holidays=extras,
            )

        st.subheader("Log")
        for line in logs:
            if line.startswith("[ERRO]"):
                st.error(line)
            elif line.startswith("[AVISO]"):
                st.warning(line)
            elif line.startswith("[IGNORADO]"):
                st.info(line)
            else:
                st.success(line)

        filename = uploaded.name
        st.download_button(
            "Baixar planilha atualizada",
            data=updated_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_updated"
        )

    st.divider()
    st.markdown(
        """
**Como funciona**
- Desconta feriados nacionais do Brasil automaticamente (biblioteca `holidays`).
- Você pode adicionar feriados extras no campo acima (ex.: municipal).
- Mantém formatação: o app só escreve valores nas colunas de resultado.
"""
    )


# ----------------------------
# ABA 2
# ----------------------------
with aba2:
    st.subheader("📅 WorkTime Manager - Dias Trabalhados")

    ano = dt.datetime.now().year

    col1, col2 = st.columns([1, 1])
    with col1:
        mes_label = st.selectbox("Selecione o mês:", MESES_PT, index=dt.datetime.now().month - 1, key="mes_wt")
        mes = int(mes_label.split(" - ")[0])
        nome_mes = mes_label.split(" - ")[1]
    with col2:
        st.markdown(f"<span class='badge'>Ano: {ano}</span>", unsafe_allow_html=True)

    usar_periodo = st.checkbox("Data de Entrada e Saída", value=False, key="periodo_wt")

    data_ini = None
    data_fim = None
    if usar_periodo:
        c1, c2 = st.columns(2)
        with c1:
            data_ini = st.date_input("Entrada:", value=dt.date.today(), key="entrada_wt")
        with c2:
            data_fim = st.date_input("Saída:", value=dt.date.today(), key="saida_wt")

    st.markdown("### Dias da semana que são folgas:")
    cols = st.columns(7)
    folgas_semana = [False] * 7
    for i, dia in enumerate(DIAS_SEMANA):
        with cols[i]:
            folgas_semana[i] = st.checkbox(dia, value=False, key=f"folga_wt_{i}")

    if st.button("Calcular Dias Trabalhados", type="primary", key="btn_calc_wt"):
        dias, feriados_dias, detalhes, (p_ini, p_fim) = calcular_dias_trabalhados(
            mes, folgas_semana, ano, data_ini, data_fim
        )

        periodo_str = f"{p_ini:02d}/{mes:02d}/{ano} até {p_fim:02d}/{mes:02d}/{ano}"
        folgas_str = ", ".join([DIAS_SEMANA[i] for i, v in enumerate(folgas_semana) if v]) or "Nenhuma"

        if detalhes:
            feriados_str = "\n".join([f"- {dia:02d}/{mes:02d}/{ano} — {detalhes[dia]}" for dia in sorted(detalhes)])
        else:
            feriados_str = ""

        st.session_state["worktime_payload"] = {
            "ano": ano,
            "mes_nome": nome_mes,
            "dias_trabalhados": dias,
            "periodo_str": periodo_str,
            "folgas_str": folgas_str,
            "feriados_str": feriados_str,
        }

        st.markdown(
            f"""
            <div class='card'>
              <h3>Resultado</h3>
              <p><b>Dias trabalhados em {nome_mes} de {ano}:</b> <span class='badge'>{dias}</span></p>
              <p><b>Período:</b> {periodo_str}</p>
              <p><b>Folgas:</b> {folgas_str}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if detalhes:
            st.info("Feriados considerados:\n\n" + "\n".join([f"• {dia:02d}/{mes:02d}/{ano} ({detalhes[dia]})" for dia in sorted(detalhes)]))
        else:
            st.info("Feriados considerados: Nenhum")

    if "worktime_payload" in st.session_state:
        payload = st.session_state["worktime_payload"]
        colA, colB = st.columns(2)

        with colA:
            pdf_bytes = make_pdf_worktime(payload)
            st.download_button(
                "📄 Baixar PDF",
                data=pdf_bytes,
                file_name=f"WorkTime_{payload['mes_nome']}_{payload['ano']}.pdf",
                mime="application/pdf",
                key="download_pdf_wt"
            )

        with colB:
            xlsx_bytes = make_excel_worktime(payload)
            st.download_button(
                "📊 Baixar Excel",
                data=xlsx_bytes,
                file_name=f"WorkTime_{payload['mes_nome']}_{payload['ano']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_xlsx_wt"
            )
