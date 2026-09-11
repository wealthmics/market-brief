"""
MICS Weekly Market Snapshot Generator
Run any day of the week — always fetches current week data.
Saturday/Sunday = final (full week). Mon–Fri = partial (week in progress).
"""

import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import os

# ── CONFIG ────────────────────────────────────────────────────────────────────

TICKERS = {
    "Nasdaq 100":     "^NDX",
    "S&P 500":        "^GSPC",
    "Nikkei 225":     "^N225",
    "FTSE 100":       "^FTSE",
    "Nifty 50":       "^NSEI",
    "Shanghai Comp":  "000001.SS",
    "Gold (XAU/USD)": "GC=F",
    "Brent Crude":    "BZ=F",
    "DXY":            "DX-Y.NYB",
    "Bitcoin":        "BTC-USD",
}

SUBTEXTS = {
    "Nasdaq 100":     "US",
    "S&P 500":        "US",
    "Nikkei 225":     "Japan",
    "FTSE 100":       "UK",
    "Nifty 50":       "India",
    "Shanghai Comp":  "China",
    "Gold (XAU/USD)": "USD / oz",
    "Brent Crude":    "USD / bbl",
    "DXY":            "USD Index",
    "Bitcoin":        "Crypto - USD",
}

OUTPUT_DIR = "."  # Change to your preferred folder

# ── COLORS ────────────────────────────────────────────────────────────────────

HEADER_BG  = "1B2A4A"
HEADER_FG  = "FFFFFF"
ROW_ALT    = "F5F7FA"
ROW_WHITE  = "FFFFFF"
GREEN_BG   = "E8F5E9"
RED_BG     = "FFEBEE"
GREEN_FG   = "1B7F34"
RED_FG     = "B71C1C"
BORDER_CLR = "D0D5DD"
LOG_HDR_BG = "2E4057"

# ── WEEK BOUNDS ───────────────────────────────────────────────────────────────

def get_week_bounds():
    """
    Always return the current calendar week.
    as_of = today if Mon–Fri (partial), last Friday if Sat/Sun (final).
    """
    today   = datetime.date.today()
    weekday = today.weekday()          # Mon=0 … Sun=6
    monday  = today - datetime.timedelta(days=weekday)

    if weekday == 5:                   # Saturday
        as_of   = today - datetime.timedelta(days=1)
        partial = False
    elif weekday == 6:                 # Sunday
        as_of   = today - datetime.timedelta(days=2)
        partial = False
    else:                              # Mon–Fri
        as_of   = today
        partial = True

    return monday, as_of, partial

# ── DATA FETCH ────────────────────────────────────────────────────────────────

def fetch_data():
    monday, as_of, partial = get_week_bounds()
    year_start = datetime.date(as_of.year, 1, 1)

    status = "PARTIAL — week in progress" if partial else "FINAL"
    print(f"Fetching week : {monday.strftime('%d %b')} – {as_of.strftime('%d %b %Y')}  [{status}]")

    symbols = list(TICKERS.values())
    raw = yf.download(
        symbols,
        start=year_start - datetime.timedelta(days=5),
        end=as_of + datetime.timedelta(days=1),
        auto_adjust=True,
        progress=False,
        group_by="ticker",
    )

    results = []

    for name, sym in TICKERS.items():
        try:
            df = raw if len(symbols) == 1 else raw[sym]
            df = df.dropna(subset=["Close"])

            # Current price = latest available close
            price      = df["Close"].iloc[-1]
            price_date = df.index[-1].date()

            # Weekly High / Low
            week_mask = (df.index.date >= monday) & (df.index.date <= as_of)
            week_df   = df.loc[week_mask]
            if week_df.empty:
                week_df = df.tail(5)
            high           = week_df["High"].max()
            low            = week_df["Low"].min()
            week_high_date = week_df["High"].idxmax().date()
            week_low_date  = week_df["Low"].idxmin().date()

            # 1W base: last trading day before this week's Monday (prev Friday close)
            prev_df    = df[df.index.date < monday]
            prev_close = prev_df["Close"].iloc[-1] if not prev_df.empty else None
            prev_date  = prev_df.index[-1].date()  if not prev_df.empty else None
            w1         = (price - prev_close) / prev_close if prev_close else None

            # YTD base: last trading day of previous year (close)
            ytd_df        = df[df.index.date < year_start]
            ytd_base      = ytd_df["Close"].iloc[-1] if not ytd_df.empty else None
            ytd_base_date = ytd_df.index[-1].date()  if not ytd_df.empty else None
            ytd           = (price - ytd_base) / ytd_base if ytd_base else None

            results.append({
                "name":           name,
                "sub":            SUBTEXTS[name],
                "price":          price,
                "high":           high,
                "low":            low,
                "w1":             w1,
                "ytd":            ytd,
                # audit fields
                "price_date":     price_date,
                "prev_close":     prev_close,
                "prev_date":      prev_date,
                "ytd_base":       ytd_base,
                "ytd_base_date":  ytd_base_date,
                "week_high_date": week_high_date,
                "week_low_date":  week_low_date,
            })
            print(f"  OK  {name:20s}  {price:>12.2f}  1W {w1*100:+.2f}%  YTD {ytd*100:+.2f}%")

        except Exception as e:
            print(f"  ERR {name}: {e}")
            results.append({
                "name": name, "sub": SUBTEXTS[name],
                "price": None, "high": None, "low": None,
                "w1": None, "ytd": None,
                "price_date": None, "prev_close": None, "prev_date": None,
                "ytd_base": None, "ytd_base_date": None,
                "week_high_date": None, "week_low_date": None,
            })

    return results, monday, as_of, partial

# ── EXCEL HELPERS ─────────────────────────────────────────────────────────────

def thin_border(color=None):
    c = color or BORDER_CLR
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)

def pct_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if value is None:
        cell.value     = "N/A"
        cell.font      = Font(name="Calibri", size=10, color="888888")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = thin_border()
        return
    cell.value         = value
    cell.number_format = '+0.00%;-0.00%;0.00%'
    if value >= 0:
        cell.fill = PatternFill("solid", fgColor=GREEN_BG)
        cell.font = Font(name="Calibri", size=10, bold=True, color=GREEN_FG)
    else:
        cell.fill = PatternFill("solid", fgColor=RED_BG)
        cell.font = Font(name="Calibri", size=10, bold=True, color=RED_FG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = thin_border()

# ── SHEET 1: MARKET SNAPSHOT ──────────────────────────────────────────────────

def build_snapshot_sheet(wb, results, monday, as_of, partial):
    ws = wb.active
    ws.title = "Market Snapshot"
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([28, 14, 14, 14, 12, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.row_dimensions[1].height = 36
    ws.merge_cells("A1:F1")
    t       = ws["A1"]
    t.value = "MICS WEEKLY MARKET SNAPSHOT" + ("  [PARTIAL]" if partial else "")
    t.font  = Font(name="Calibri", size=14, bold=True, color=HEADER_FG)
    t.fill  = PatternFill("solid", fgColor=HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    # Subtitle
    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:F2")
    s = ws["A2"]
    if partial:
        s.value = (
            f"Week: {monday.strftime('%d %b')} – {as_of.strftime('%d %b %Y')}   |   "
            f"Data as of {as_of.strftime('%A, %d %b %Y')}  (week in progress — 1W/High/Low will update)"
        )
    else:
        s.value = (
            f"Week: {monday.strftime('%d %b')} – {as_of.strftime('%d %b %Y')}   |   "
            f"Closing prices as of {as_of.strftime('%A, %d %b %Y')}"
        )
    s.font      = Font(name="Calibri", size=9, color="555555", italic=True)
    s.fill      = PatternFill("solid", fgColor="EEF1F6")
    s.alignment = Alignment(horizontal="center", vertical="center")

    # Headers
    ws.row_dimensions[3].height = 28
    for col, h in enumerate(["MARKET", "PRICE", "HIGH", "LOW", "1W", "YTD"], 1):
        c       = ws.cell(row=3, column=col, value=h)
        c.font  = Font(name="Calibri", size=10, bold=True, color=HEADER_FG)
        c.fill  = PatternFill("solid", fgColor=HEADER_BG)
        c.alignment = Alignment(
            horizontal="left" if col == 1 else "center",
            vertical="center", indent=1 if col == 1 else 0
        )
        c.border = thin_border()

    # Data rows
    for i, r in enumerate(results):
        row = 4 + i
        ws.row_dimensions[row].height = 30
        bg = ROW_ALT if i % 2 else ROW_WHITE

        nc           = ws.cell(row=row, column=1)
        nc.value     = f"{r['name']}\n{r['sub']}"
        nc.font      = Font(name="Calibri", size=10)
        nc.fill      = PatternFill("solid", fgColor=bg)
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        nc.border    = thin_border()

        for col, key in [(2, "price"), (3, "high"), (4, "low")]:
            c               = ws.cell(row=row, column=col)
            c.value         = r[key]
            c.number_format = "#,##0.00"
            c.font          = Font(name="Calibri", size=10)
            c.fill          = PatternFill("solid", fgColor=bg)
            c.alignment     = Alignment(horizontal="center", vertical="center")
            c.border        = thin_border()

        pct_cell(ws, row, 5, r["w1"])
        pct_cell(ws, row, 6, r["ytd"])

    # Footer
    fr = 4 + len(results)
    ws.merge_cells(f"A{fr}:F{fr}")
    f       = ws.cell(row=fr, column=1)
    f.value = (
        f"Weekly data · Closing prices as of {as_of.strftime('%A, %d %b %Y')}"
        + (" · PARTIAL — week not yet complete" if partial
           else " · Prices may vary slightly due to data source timing.")
    )
    f.font      = Font(name="Calibri", size=8, italic=True, color="888888")
    f.fill      = PatternFill("solid", fgColor="EEF1F6")
    f.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── SHEET 2: CALCULATION LOG ──────────────────────────────────────────────────

def build_log_sheet(wb, results, monday, as_of):
    ws = wb.create_sheet("Calculation Log")
    ws.sheet_view.showGridLines = False

    # Column widths
    widths = [22, 12, 14, 14, 14, 14, 14, 14, 12, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:J1")
    t       = ws["A1"]
    t.value = f"CALCULATION LOG — Week {monday.strftime('%d %b')} to {as_of.strftime('%d %b %Y')}"
    t.font  = Font(name="Calibri", size=12, bold=True, color=HEADER_FG)
    t.fill  = PatternFill("solid", fgColor=LOG_HDR_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")

    # Section headers with grouping
    # WTD section
    ws.row_dimensions[2].height = 18
    # ws.merge_cells("A2:A3")
    ws.merge_cells("B2:F2")
    ws.merge_cells("G2:J2")

    for cell, text, bg in [
        ("A2", "MARKET",        LOG_HDR_BG),
        ("B2", "1W CALCULATION",  "1F4E79"),
        ("G2", "YTD CALCULATION",  "1F4E79"),
    ]:
        c       = ws[cell]
        c.value = text
        c.font  = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill  = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border("4472C4")

    # Column headers row 3
    ws.row_dimensions[3].height = 22
    col_headers = [
        "MARKET",
        "AS OF DATE", "CURRENT CLOSE", "1W BASE DATE", "1W BASE CLOSE", "1W RETURN",
        "YTD BASE DATE", "YTD BASE CLOSE", "YTD RETURN",
        "HIGH DATE / LOW DATE",
    ]
    for col, h in enumerate(col_headers, 1):
        c       = ws.cell(row=3, column=col, value=h)
        c.font  = Font(name="Calibri", size=9, bold=True, color=HEADER_FG)
        c.fill  = PatternFill("solid", fgColor="2E4057")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border()

    # Data rows
    def fmt_date(d):
        return d.strftime("%d %b %Y") if d else "N/A"

    def fmt_num(v):
        return round(v, 4) if v is not None else "N/A"

    for i, r in enumerate(results):
        row = 4 + i
        ws.row_dimensions[row].height = 22
        bg = ROW_ALT if i % 2 else ROW_WHITE

        row_vals = [
            r["name"],
            fmt_date(r["price_date"]),
            fmt_num(r["price"]),
            fmt_date(r["prev_date"]),
            fmt_num(r["prev_close"]),
            r["w1"],
            fmt_date(r["ytd_base_date"]),
            fmt_num(r["ytd_base"]),
            r["ytd"],
            f"High: {fmt_date(r['week_high_date'])}  /  Low: {fmt_date(r['week_low_date'])}",
        ]

        for col, val in enumerate(row_vals, 1):
            c       = ws.cell(row=row, column=col, value=val)
            c.fill  = PatternFill("solid", fgColor=bg)
            c.border = thin_border()
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.font  = Font(name="Calibri", size=9)

            # Return columns — colour coded
            if col in (6, 9) and isinstance(val, float):
                c.number_format = '+0.00%;-0.00%;0.00%'
                c.font = Font(name="Calibri", size=9, bold=True,
                              color=GREEN_FG if val >= 0 else RED_FG)
                c.fill = PatternFill("solid", fgColor=GREEN_BG if val >= 0 else RED_BG)

            # Price / base columns — number format
            if col in (3, 5, 8) and isinstance(val, float):
                c.number_format = "#,##0.00"

            # Market name — left aligned
            if col == 1:
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                c.font = Font(name="Calibri", size=9, bold=True)

    # Formula note
    note_row = 4 + len(results) + 1
    ws.merge_cells(f"A{note_row}:J{note_row}")
    n       = ws.cell(row=note_row, column=1)
    n.value = (
        "1W = (Current Close – 1W Base Close) / 1W Base Close     |     "
        "YTD = (Current Close – YTD Base Close) / YTD Base Close     |     "
        "1W Base = last trading day before this week's Monday     |     "
        "YTD Base = last trading day of previous year"
    )
    n.font      = Font(name="Calibri", size=8, italic=True, color="555555")
    n.fill      = PatternFill("solid", fgColor="EEF1F6")
    n.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── EXCEL ─────────────────────────────────────────────────────────────────────

def build_excel(results, monday, as_of, partial):
    wb = openpyxl.Workbook()
    build_snapshot_sheet(wb, results, monday, as_of, partial)
    build_log_sheet(wb, results, monday, as_of)

    suffix = "_PARTIAL" if partial else ""
    fname  = f"MICS_Market_Snapshot_{as_of.strftime('%Y%m%d')}{suffix}.xlsx"
    fpath  = os.path.join(OUTPUT_DIR, fname)
    wb.save(fpath)
    print(f"Saved Excel : {fpath}")
    return fpath

# ── HTML SNIPPET ───────────────────────────────────────────────────────────────

def fmt_price(v):
    """Format price with commas, 2 decimal places."""
    if v is None:
        return "N/A"
    return f"{v:,.2f}"
 
def fmt_pct(v):
    """Format return as +1.68% / -1.27% style."""
    if v is None:
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v*100:.2f}%"
 
FONT_POPPINS = "Poppins, Arial, sans-serif"
FONT_INTER   = "Inter, Arial, sans-serif"
 
def pct_span(v):
    """Return a coloured <span> for a return value."""
    if v is None:
        return f"<span style='font-family:{FONT_INTER};font-size:8px;color:#999;'>N/A</span>"
    text = fmt_pct(v)
    if v >= 0:
        bg, fg = "#dcfce7", "#166534"
    else:
        bg, fg = "#fee2e2", "#991b1b"
    return (
        f"<span style='font-family:{FONT_INTER};background:{bg};color:{fg};"
        f"font-size:8px;font-weight:700;padding:2px 3px;'>{text}</span>"
    )
 
def build_html(results, as_of, partial):
    date_str     = as_of.strftime("%A, %d %b %Y")
    partial_note = " · Data incomplete — week still in progress" if partial else ""
 
    rows_html = ""
    for i, r in enumerate(results):
        bg = "#ffffff" if i % 2 == 0 else "#f7f5f0"
        rows_html += f"""
        <tr style="background:{bg};border-bottom:1px solid #eee;">
          <td style="padding:6px;">
            <div style="font-family:{FONT_POPPINS};font-size:8px;font-weight:700;color:#0b2558;">{r['name']}</div>
            <div style="font-family:{FONT_INTER};font-size:6px;color:#999;">{r['sub']}</div>
          </td>
          <td style="padding:6px;text-align:center;font-family:{FONT_INTER};font-size:8px;">{fmt_price(r['price'])}</td>
          <td style="padding:6px;text-align:center;font-family:{FONT_INTER};font-size:8px;">{fmt_price(r['high'])}</td>
          <td style="padding:6px;text-align:center;font-family:{FONT_INTER};font-size:8px;">{fmt_price(r['low'])}</td>
          <td style="padding:6px;text-align:center;">{pct_span(r['w1'])}</td>
          <td style="padding:6px;text-align:center;">{pct_span(r['ytd'])}</td>
        </tr>"""
 
    html = f"""    <!-- ═══ MARKET SNAPSHOT — auto-generated {date_str} ═══ -->
    <div style="padding:16px 16px;border-bottom:1px solid #e0dcd4;">
      <div style="font-family:{FONT_POPPINS};font-size:10px;letter-spacing:0.5px;text-transform:uppercase;color:#8a8a8a;margin-bottom:6px;font-weight:500;">
        Markets Last Week</div>
      <div style="font-family:{FONT_POPPINS};font-size:17px;font-weight:700;color:#0b2558;margin-bottom:14px;">
        How the markets behaved</div>
 
      <table width="100%" cellpadding="0" cellspacing="0" style="table-layout:fixed;">
        <tr style="background:#0b2558;">
          <td width="20%" align="left" valign="middle" style="padding:6px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">MARKET</span></td>
          <td width="16%" align="center" valign="middle" style="padding:4px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">PRICE</span></td>
          <td width="16%" align="center" valign="middle" style="padding:4px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">HIGH</span></td>
          <td width="16%" align="center" valign="middle" style="padding:4px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">LOW</span></td>
          <td width="16%" align="center" valign="middle" style="padding:4px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">1W</span></td>
          <td width="16%" align="center" valign="middle" style="padding:4px;"><span
              style="font-family:{FONT_POPPINS};font-size:8px;color:#ffffff;font-weight:700;">YTD</span></td>
        </tr>
{rows_html}
      </table>
 
      <div style="font-family:{FONT_INTER};margin-top:8px;font-size:8px;color:#bbb;">
        Weekly data &middot; Closing prices as of {date_str}{partial_note} &middot;
        Closing prices are subject to change, as markets were still active at the time of writing.</div>
    </div>"""
 
    suffix  = "_PARTIAL" if partial else ""
    fname   = f"MICS_Market_Snapshot_{as_of.strftime('%Y%m%d')}{suffix}.html"
    fpath   = os.path.join(OUTPUT_DIR, fname)
    with open(fpath, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Saved HTML  : {fpath}")
    return fpath
 
# ── MAIN ──────────────────────────────────────────────────────────────────────
 
if __name__ == "__main__":
    results, monday, as_of, partial = fetch_data()
    build_excel(results, monday, as_of, partial)
    build_html(results, as_of, partial)
