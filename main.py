#!/usr/bin/env python3
"""Fetch Nord Pool day-ahead prices and save them into a Google Sheet."""

from __future__ import annotations

import json
import urllib.parse
import urllib.request
import traceback
import atexit
import os
import time
from datetime import datetime, date, timedelta
from calendar import monthrange
from typing import Dict, List
from io import StringIO

import requests
import gspread
import pandas as pd

# =====================================================
# Terminal protection logic
# =====================================================
def pause_before_exit():
    # Final safety net so the terminal doesn't snap shut on an unhandled crash
    input("\n[Rendszer] Nyomj Enter-t a kilépéshez a terminálból...")

atexit.register(pause_before_exit)

# =====================================================
# Google Sheets Configuration
# =====================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "running-app-410107-adec34aa4983.json")

# SPREADSHEET IDs
SPREADSHEET_ID = "1QWLwJEc1G7qfj3QB_N6EtHUACtYgbNDad-jfyCFq0as" # EU DAM Spreads
EOD_SPREADSHEET_ID = "1ZrYC2Fbi9drLFgHwEPvoTz9yxambhByf3elrM1v0KUs" # EOD Analysis (Update this monthly!)

API_URL = "https://dataportal-api.nordpoolgroup.com/api/DayAheadPriceIndices"
BASE_URL = "https://labs.hupx.hu/data/v1"

# Added UA to the end of the index names
INDEX_NAMES = [
    "EE", "LV", "AT", "BE", "FR", "GER", "NL", "PL", 
    "DK1", "DK2", "FI", "HU", "BG", "TEL", "UA"
]

# Added the new spreads: TEL (RO) -> UA and HU -> UA
NEIGHBOUR_PAIRS = [
    ("AT", "GER"), ("AT", "HU"), ("BE", "FR"), ("BE", "NL"),
    ("BG", "TEL"), ("DK1", "DK2"), ("DK1", "GER"), ("DK2", "GER"),
    ("GER", "FR"), ("GER", "NL"), ("DK1", "NL"), ("EE", "LV"),
    ("FI", "EE"), ("HU", "TEL"), ("TEL", "HU"), ("PL", "GER"),
    ("PL", "NL"), ("TEL", "UA"), ("HU", "UA")
]

def show_menu() -> str:
    inner_width = 44
    def box_line(content: str = "") -> str:
        return f"║{content:<{inner_width}}║"

    menu_lines = [
        "",
        f"╔{'═' * inner_width}╗",
        box_line("    CBC DAM Spread (Google Sheets) "),
        f"╠{'═' * inner_width}╣",
        box_line(),
        box_line("  [1]  Mai adatok (Today)"),
        box_line("  [2]  Konkrét dátum (pl. 04-01)"),
        box_line("  [3]  Teljes hónap (pl. 2024-05)"),
        box_line("  [0]  Kilépés"),
        box_line(),
        f"╚{'═' * inner_width}╝",
        "",
        "Válassz [1/2/3/0]: ",
    ]
    return input("\n".join(menu_lines)).strip()


def parse_target_dates() -> List[date] | None:
    while True:
        choice = show_menu()
        if choice == "0": 
            return None
        if choice == "1": 
            return [datetime.now().date()]
        if choice == "2":
            raw = input("Add meg a dátumot (MM-DD vagy YYYY-MM-DD): ").strip()
            for fmt in ("%Y-%m-%d", "%m-%d"):
                try:
                    parsed = datetime.strptime(raw, fmt)
                    if fmt == "%m-%d":
                        parsed = parsed.replace(year=datetime.now().year)
                    return [parsed.date()]
                except ValueError:
                    continue
            print("Hibás dátum formátum. Próbáld újra.")
            continue
        if choice == "3":
            raw = input("Add meg a hónapot (YYYY-MM): ").strip()
            try:
                parsed = datetime.strptime(raw, "%Y-%m")
                year, month = parsed.year, parsed.month
                _, last_day = monthrange(year, month)
                return [date(year, month, d) for d in range(1, last_day + 1)]
            except ValueError:
                print("Hibás hónap formátum. Próbáld újra (pl. 2024-05).")
                continue
        print("Érvénytelen választás. Kérlek válassz 1, 2, 3 vagy 0 opciót.")


def fetch_prices(target_date: date) -> Dict:
    # Filter out UA since Nord Pool API doesn't support it directly
    nord_pool_indexes = [idx for idx in INDEX_NAMES if idx != "UA"]
    params = {
        "date": target_date.isoformat(),
        "market": "DayAhead",
        "indexNames": ",".join(nord_pool_indexes),
        "currency": "EUR",
        "resolutionInMinutes": 60,
    }
    headers = {
        "accept": "application/json, text/plain, */*",
        "origin": "https://data.nordpoolgroup.com",
        "referer": "https://data.nordpoolgroup.com/",
        "user-agent": "Mozilla/5.0",
    }
    response = requests.get(API_URL, params=params, headers=headers, timeout=30)
    response.raise_for_status()
    return response.json()


def fetch_json(endpoint: str, filters: List[str], limit: int = 200) -> List[Dict]:
    filter_str = ",".join(filters)
    url = f"{BASE_URL}/{endpoint}?filter={urllib.parse.quote(filter_str)}&limit={limit}"
    all_data: List[Dict] = []
    while url:
        req = urllib.request.Request(url)
        req.add_header("User-Agent", "HUPX-Fetcher/1.0")
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode())
        all_data.extend(body.get("data", []))
        url = body.get("nextPage")
    return all_data


def fetch_dam(date_str: str) -> Dict[int, Dict[str, float | None]]:
    next_date = (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
    rows = fetch_json(
        "dam_aggregated_trading_data",
        [f"DeliveryDay__gte__{date_str}", f"DeliveryDay__lt__{next_date}", "Region__eq__HU"],
    )
    result: Dict[int, Dict[str, float | None]] = {}
    for r in rows:
        hour = int(r["ProductH"])
        result[hour] = {"price": r.get("Price"), "volume": r.get("Volume")}
    return result


def fetch_ua_dam(target_date: date) -> Dict[int, float | None]:
    """Fetches UA DAM prices from the EOD analysis Google Sheet."""
    result: Dict[int, float | None] = {}
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    
    try:
        sh = gc.open_by_key(EOD_SPREADSHEET_ID)
        # The EOD sheet names are just the day of the month (e.g., "1", "2", "31")
        day_str = str(target_date.day)
        ws = sh.worksheet(day_str)
        
        # Column M is 13th letter. Fetch rows 7 to 30.
        # This returns a list of lists, e.g., [['45.2'], ['42.1'], ...]
        cell_values = ws.get('M7:M30')
        
        for i in range(24):
            hour = i + 1
            if i < len(cell_values) and cell_values[i]:
                val_str = cell_values[i][0]
                if val_str is not None and str(val_str).strip() != "":
                    # Handle potential European comma decimal formatting
                    val_clean = str(val_str).replace(',', '.').replace(' ', '')
                    try:
                        result[hour] = float(val_clean)
                    except ValueError:
                        result[hour] = None
                else:
                    result[hour] = None
            else:
                result[hour] = None
                
    except gspread.exceptions.WorksheetNotFound:
        print(f"      -> HIBA: A '{day_str}' nevű lap nem található az EOD sheetben.")
    except Exception as e:
        print(f"      -> UA DAM adatok nem elérhetőek (EOD Sheet): {e}")
        
    return result


def build_rows(payload: Dict, hu_dam: Dict[int, Dict[str, float | None]], ua_dam: Dict[int, float | None]) -> List[List[float | str | None]]:
    rows: List[List[float | str | None]] = []
    entries = payload.get("multiIndexEntries", [])
    if not entries:
        return rows
        
    for idx, item in enumerate(entries, start=1):
        entry_per_area = item.get("entryPerArea", {})
        hu_price = hu_dam.get(idx, {}).get("price")
        ua_price = ua_dam.get(idx)
        
        entry_per_area["HU"] = hu_price
        entry_per_area["UA"] = ua_price
        
        row: List[float | str | None] = [idx]
        for area in INDEX_NAMES:
            val = entry_per_area.get(area)
            row.append(val if val is not None else "")
        rows.append(row)
    return rows


def get_col_letter(col_idx: int) -> str:
    if col_idx <= 26:
        return chr(64 + col_idx)
    return chr(64 + col_idx // 26) + chr(64 + col_idx % 26)


def rgb(r: int, g: int, b: int) -> dict:
    return {"red": r/255.0, "green": g/255.0, "blue": b/255.0}


def save_to_google_sheets(target_date: date, rows: List[List[float | str | None]]) -> None:
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    sh = gc.open_by_key(SPREADSHEET_ID)
    sheet_name = target_date.strftime("%Y.%m.%d.")
    
    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=100, cols=40)

    # Convert "TEL" to "RO" dynamically for headers so there are no duplicates visually
    display_index_names = ["RO" if x == "TEL" else x for x in INDEX_NAMES]

    # 2. Base pricing setup
    headers = ["Hour"] + display_index_names
    full_data = [headers]
    
    for row in rows:
        full_data.append(row)
        
    avg_row_num = len(rows) + 2
    avg_row = ["AVG"]
    for col in range(2, len(headers) + 1):
        col_letter = get_col_letter(col)
        avg_row.append(f'=IFERROR(AVERAGE({col_letter}2:{col_letter}{avg_row_num - 1}); "")')
    full_data.append(avg_row)

    full_data.append([])
    full_data.append([])

    # 3. Spread tables setup
    spread_start_row = avg_row_num + 3
    # Rename TEL to RO in Spread mapping headers
    spread_headers = ["Hour"] + [f"{'RO' if left == 'TEL' else left}-{'RO' if right == 'TEL' else right}" for left, right in NEIGHBOUR_PAIRS]
    full_data.append(spread_headers)

    for hour_idx in range(1, len(rows) + 1):
        spread_row = [hour_idx]
        for left, right in NEIGHBOUR_PAIRS:
            try:
                left_col = INDEX_NAMES.index(left) + 2
                right_col = INDEX_NAMES.index(right) + 2
                col_L = get_col_letter(left_col)
                col_R = get_col_letter(right_col)
                row_n = hour_idx + 1
                
                # JAVÍTOTT SPREAD LOGIKA: Jobb (eladás) - Bal (vétel).
                # Ha kisebb mint 0, akkor 0. Egyébként a különbség.
                formula = f'=IFERROR(IF({col_R}{row_n}-{col_L}{row_n}<0; 0; {col_R}{row_n}-{col_L}{row_n}); "")'
                spread_row.append(formula)
            except ValueError:
                spread_row.append("")
                
        full_data.append(spread_row)

    spread_avg_row_num = spread_start_row + len(rows) + 1
    spread_avg_row = ["AVG"]
    for col_idx in range(2, len(spread_headers) + 1):
        col_letter = get_col_letter(col_idx)
        spread_avg_row.append(f'=IFERROR(AVERAGE({col_letter}{spread_start_row + 1}:{col_letter}{spread_avg_row_num - 1}); "")')
    full_data.append(spread_avg_row)

    clean_full_data = [["" if v is None else v for v in row] for row in full_data]

    try:
        ws.update(values=clean_full_data, range_name='A1', value_input_option='USER_ENTERED')
    except TypeError:
        ws.update('A1', clean_full_data, value_input_option='USER_ENTERED')

    # 4. Styling Block
    try:
        border_style = {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}}
        borders = {"top": border_style, "bottom": border_style, "left": border_style, "right": border_style}
        center_align = "CENTER"

        def create_format(bg_color, bold=False):
            return {
                "backgroundColor": bg_color,
                "horizontalAlignment": center_align,
                "textFormat": {"bold": bold},
                "borders": borders
            }

        header_fmt = create_format(rgb(189, 215, 238), bold=True)
        data_fmt = create_format(rgb(252, 228, 214))
        avg_fmt = create_format(rgb(255, 255, 0), bold=True)
        spread_data_fmt = create_format(rgb(226, 240, 217))

        base_last_col = get_col_letter(len(headers))
        spread_last_col = get_col_letter(len(spread_headers))

        ws.batch_format([
            {"range": f"A1:{base_last_col}1", "format": header_fmt},
            {"range": f"A2:{base_last_col}{avg_row_num - 1}", "format": data_fmt},
            {"range": f"A{avg_row_num}:{base_last_col}{avg_row_num}", "format": avg_fmt},
            {"range": f"A{spread_start_row}:{spread_last_col}{spread_start_row}", "format": header_fmt},
            {"range": f"A{spread_start_row + 1}:{spread_last_col}{spread_avg_row_num - 1}", "format": spread_data_fmt},
            {"range": f"A{spread_avg_row_num}:{spread_last_col}{spread_avg_row_num}", "format": avg_fmt},
        ])
    except Exception as fmt_err:
        print(f"    -> (Formázási hiba elhanyagolva: {fmt_err})")


def main() -> None:
    while True:
        target_dates = parse_target_dates()
        if target_dates is None:
            print("Kilépés kezdeményezve...")
            break

        for target_date in target_dates:
            print(f"\n[1/4] Adatok lekérése ({target_date.isoformat()})...")
            try:
                payload = fetch_prices(target_date)
            except Exception as exc:
                print(f"      -> Nord Pool adatok nem elérhetőek: {exc}")
                payload = {}

            print(f"[2/4] HU DAM lekérése...")
            hu_dam: Dict[int, Dict[str, float | None]] = {}
            try:
                hu_dam = fetch_dam(target_date.isoformat())
            except Exception as exc:
                print(f"      -> HU DAM adatok nem elérhetőek: {exc}")

            print(f"[3/4] UA DAM lekérése (EOD Google Sheet)...")
            ua_dam: Dict[int, float | None] = {}
            try:
                ua_dam = fetch_ua_dam(target_date)
            except Exception as exc:
                print(f"      -> UA DAM adatok nem elérhetőek: {exc}")

            rows = build_rows(payload, hu_dam, ua_dam)

            if not rows:
                print(f"      -> Nincs elérhető piaci adat. Üres struktúra generálása...")
                rows = [[idx] + [""] * len(INDEX_NAMES) for idx in range(1, 25)]
            else:
                print(f"      Sikeresen feldolgozva {len(rows)} órányi adat.")

            print(f"[4/4] Google Sheets feltöltés folyamatban...")
            try:
                save_to_google_sheets(target_date, rows)
                print(f"[✓] SIKER: '{target_date.strftime('%Y.%m.%d.')}' frissítve.")
            except Exception as e:
                print(f"\nHIBA történt a táblázat írásakor ({target_date.isoformat()}):")
                traceback.print_exc()
            
            # End of month check for the terminal notification
            _, last_day = monthrange(target_date.year, target_date.month)
            if target_date.day == last_day:
                print("\n" + "="*70)
                print(f" ⚠️ FIGYELEM: A hónap ({target_date.strftime('%Y-%m')}) véget ért! ⚠️")
                print(f" Kérlek, frissítsd az EOD_SPREADSHEET_ID-t a kódban a következő hónaphoz!")
                print("="*70 + "\n")

            if len(target_dates) > 1:
                time.sleep(1)


if __name__ == "__main__":
    main()
