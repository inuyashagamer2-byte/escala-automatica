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


# ----------------------------
# Helpers: parsing and counting
# ----------------------------

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
        # dayfirst=True to match BR format commonly
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

    # Normalize
    tok = [t.upper() for t in tokens]
    tok = ["SAB" if t == "SÁB" else t for t in tok]

    # If contains "FOLGA" => days listed are off-days
    if "FOLGA" in s:
        off = {PT_DOW[t] for t in tok if t in PT_DOW}
        return set(range(7)) - off

    # Handle ranges like "SEG A SEX"
    # We look for pattern "<DAY> A <DAY>" by using first two tokens if ' A ' present.
    # (Works for most texts in your sheet: "SEG A SEX", etc.)
    if " A " in s and len(tok) >= 2:
        a = PT_DOW.get(tok[0])
        b = PT_DOW.get(tok[1])
        if a is None or b is None:
            return None
        if a <= b:
            return set(range(a, b + 1))
        # wrap around (e.g., "SEX A TER")
        return set(list(range(a, 7)) + list(range(0, b + 1)))

    # Otherwise treat as explicit list of days
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
    """
    Expect sheet like '01.2026' -> (2026, 1).
    """
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
    Header_map keys are normalized (upper, strip).
    """
    for r in range(1, 51):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if not values:
            continue
        normalized = []
        for v in values:
            if v is None:
                normalized.append("")
            else:
                normalized.append(str(v).strip().upper())
        # Heuristic: if contains required headers
        if "INÍCIO ESCALA NOVA" in normalized and "ESCALA NOVA" in normalized:
            # build map
            m: Dict[str, int] = {}
            for idx, name in enumerate(normalized, start=1):
                if name:
                    m[name] = idx
            return r, m
    raise ValueError("Não encontrei a linha de cabeçalho (preciso de 'INÍCIO ESCALA NOVA' e 'ESCALA NOVA').")


def ensure_column(ws: Worksheet, header_row: int, header_map: Dict[str, int], header_name: str) -> int:
    """
    Ensure a column with header header_name exists; if not, create at end.
    Returns column index.
    """
    key = header_name.strip().upper()
    if key in header_map:
        return header_map[key]

    new_col = ws.max_column + 1
    ws.cell(row=header_row, column=new_col).value = header_name
    header_map[key] = new_col
    return new_col


def build_brazil_holidays(years: Set[int]) -> Set[dt.date]:
    """
    National Brazil holidays for given years using python-holidays.
    """
    br = holidays.Brazil(years=years)
    return {d for d in br.keys()}


def parse_extra_holidays(text: str) -> Set[dt.date]:
    """
    Parse user-entered holidays: lines with dates like 25/01/2026, 2026-01-25, etc.
    """
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
            # ignore invalid lines
            pass
    return out


# ----------------------------
# Main processing
# ----------------------------

def process_workbook(
    file_bytes: bytes,
    sheet_names: Optional[List[str]],
    extra_holidays: Set[dt.date],
    backup_mode: bool = False,  # not used in web; kept for future
) -> Tuple[bytes, List[str]]:
    """
    Load xlsx from bytes, update selected sheets, return updated file bytes and logs.
    """
    logs: List[str] = []
    bio = io.BytesIO(file_bytes)
    wb = load_workbook(bio)
    target_sheets = sheet_names or wb.sheetnames

    # collect years for holiday generation from target sheets
    years: Set[int] = set()
    sheet_ym: Dict[str, Tuple[int, int]] = {}

    for sname in target_sheets:
        ym = parse_sheet_month_year(sname)
        if ym:
            y, m = ym
            years.add(y)
            years.add(y)  # explicit
            sheet_ym[sname] = (y, m)

    # also include years from extra holidays
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

        # Determine old scale column for this sheet (e.g. "ESCALA 01.2026")
        old_scale_header = f"ESCALA {month_sheet:02d}.{year}"
        old_key = old_scale_header.upper()

        if old_key not in header_map:
            logs.append(f"[ERRO] Aba '{sname}': não achei a coluna '{old_scale_header}'.")
            continue

        col_old = header_map[old_key]
        col_new = header_map["ESCALA NOVA"]
        col_start = header_map["INÍCIO ESCALA NOVA"]

        # Ensure output columns exist
        # For your use-case we always calculate Jan/Feb 2026 columns, but you can expand later.
        # We'll do: month_sheet and month_sheet+1 as "01" and "02" relative to sheet year.
        # If sheet is 01.2026 -> calc 01.2026 and 02.2026.
        m1 = month_sheet
        y1 = year
        if m1 == 12:
            m2, y2 = 1, year + 1
        else:
            m2, y2 = m1 + 1, year

        # Column headers to fill
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

        # Iterate rows until blank line (or max_row)
        processed = 0
        errors = 0

        last_row = ws.max_row
        for r in range(header_row + 1, last_row + 1):
            # if row seems empty, skip (but don't break aggressively)
            v_old = ws.cell(row=r, column=col_old).value
            v_new = ws.cell(row=r, column=col_new).value
            v_start = ws.cell(row=r, column=col_start).value

            if v_old is None and v_new is None and v_start is None:
                continue

            start_date = safe_parse_date(v_start)
            if start_date is None:
                # If no start date, we can't compute (skip with error)
                errors += 1
                continue

            days_old = parse_schedule_days(v_old)
            days_new = parse_schedule_days(v_new)

            if days_old is None or days_new is None:
                errors += 1
                continue

            # month 1
            m1_start, m1_end = month_bounds(y1, m1)
            if start_date > m1_end:
                old_1 = new_1 = due_1 = 0
            else:
                calc_start = max(start_date, m1_start)
                old_1 = count_workdays(calc_start, m1_end, days_old, holiday_set)
                new_1 = count_workdays(calc_start, m1_end, days_new, holiday_set)
                due_1 = old_1 - new_1

            # month 2
            m2_start, m2_end = month_bounds(y2, m2)
            if start_date > m2_end:
                old_2 = new_2 = due_2 = 0
            else:
                calc_start = max(start_date, m2_start)
                old_2 = count_workdays(calc_start, m2_end, days_old, holiday_set)
                new_2 = count_workdays(calc_start, m2_end, days_new, holiday_set)
                due_2 = old_2 - new_2

            total = due_1 + due_2

            # Write values only (no style changes)
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


# ----------------------------
# Streamlit UI
# ----------------------------

st.set_page_config(page_title="Atualizador de Escalas (Excel)", layout="centered")

st.title("Atualizar planilha de alteração de escala")
st.caption("Faça upload do .xlsx, clique em **ATUALIZAR PLANILHA** e baixe o arquivo atualizado (mesma formatação, só valores).")

uploaded = st.file_uploader("Envie sua planilha (.xlsx)", type=["xlsx"])

extra_holidays_text = st.text_area(
    "Feriados extras (opcional) — 1 por linha (ex: 25/01/2026)",
    value="",
    placeholder="Ex:\n25/01/2026\n16/07/2026",
    height=120,
)

process_all_sheets = st.checkbox("Processar todas as abas no formato MM.AAAA", value=True)

selected_sheets = None
if uploaded is not None:
    # Load workbook just to list sheets (safe, no write)
    try:
        wb_preview = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        sheetnames = wb_preview.sheetnames
        wb_preview.close()
    except Exception as e:
        st.error(f"Não consegui abrir esse arquivo: {e}")
        st.stop()

    if not process_all_sheets:
        selected_sheets = st.multiselect(
            "Selecione as abas para processar",
            options=sheetnames,
            default=[s for s in sheetnames if re.match(r"^\d{1,2}\.\d{4}$", s.strip())],
        )

btn = st.button("ATUALIZAR PLANILHA", type="primary", disabled=(uploaded is None))

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

    # Download updated file with same name (best possible on web)
    filename = uploaded.name
    st.download_button(
        "Baixar planilha atualizada",
        data=updated_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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