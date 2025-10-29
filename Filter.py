import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, simpledialog

# === Configuration ===
SHEET_MASSY = 'LAX'
CLEAN_SHEETS = ['Tony单提 (需整理)', '加单', 'CargoVan', '卡车']

COL_MASSY_ADDR = 'processed_address'
COL_MASSY_PHONE = 'Merged Mobiles'

COL_CLEAN_ADDR = 'Pickup Address*'
COL_CLEAN_PHONE = 'Phone Number*'

OUTPUT_FILE = 'new addresses_numbers.xlsx'
OUTPUT_SHEET = 'new addresses'

GEMINI_ENDPOINT = (
    'https://generativelanguage.googleapis.com/v1beta/models/'
    'gemini-2.5-flash:generateContent'
)


def gemini_match(addr1: str, addr2: str, api_key: str) -> bool:
    """Ask Gemini if two addresses refer to the same location."""
    prompt = (
        f'Do these two addresses refer to the same physical location '
        f'(including apt/unit/door number)?\n'
        f'Address A: {addr1}\nAddress B: {addr2}\n'
        f'Answer only "yes" or "no".'
    )
    headers = {
        'Content-Type': 'application/json',
        'x-goog-api-key': api_key
    }
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 2}
    }
    try:
        resp = requests.post(GEMINI_ENDPOINT, json=body, headers=headers, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        text = ''
        if 'candidates' in data and len(data['candidates']) > 0:
            parts = data['candidates'][0].get('content', {}).get('parts', [])
            if parts:
                text = parts[0].get('text', '').strip().lower()
        return text.startswith('y')
    except Exception as e:
        print(f"[Gemini API error]: {e}")
        return False


def select_file(title):
    """Open a clean file dialog (works on all OS)."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.destroy()
    return file_path


def main():
    root = tk.Tk()
    root.withdraw()

    # File selection dialogs
    massy_path = select_file("Select the massy Excel file (e.g., Temu.xlsx)")
    if not massy_path:
        print("❌ No massy file selected. Exiting.")
        return

    clean_path = select_file("Select the clean Excel file (e.g., 半托管长期预约模版.xlsx)")
    if not clean_path:
        print("❌ No clean file selected. Exiting.")
        return

    # AI API key dialog
    api_key = simpledialog.askstring(
        "Gemini API Key",
        "Enter your Gemini API key (required for AI matching):",
        show='*'
    )
    if not api_key:
        print("❌ No API key provided. Exiting.")
        return

    print("📂 Loading Excel files...")
    df_massy = pd.read_excel(massy_path, sheet_name=SHEET_MASSY, dtype=str)

    # Combine all clean sheets
    clean_parts = []
    for sheet in CLEAN_SHEETS:
        df = pd.read_excel(clean_path, sheet_name=sheet, dtype=str)
        df['sheet_source'] = sheet
        clean_parts.append(df[[COL_CLEAN_ADDR, COL_CLEAN_PHONE, 'sheet_source']])
    df_clean = pd.concat(clean_parts, ignore_index=True)

    output_rows = []
    seen_addresses = set()

    print("🔍 Comparing addresses via Gemini AI...")
    for _, row_massy in df_massy.iterrows():
        massy_addr = row_massy[COL_MASSY_ADDR]
        if not massy_addr or massy_addr in seen_addresses:
            continue
        seen_addresses.add(massy_addr)

        matched = False
        matched_phones = []

        # Compare with every clean sheet address using AI
        for _, row_clean in df_clean.iterrows():
            clean_addr = row_clean[COL_CLEAN_ADDR]
            if gemini_match(massy_addr, clean_addr, api_key):
                matched = True
                matched_phones.append(str(row_clean[COL_CLEAN_PHONE]))
                break  # stop after first match

        # Determine phone color
        massy_phone = str(row_massy.get(COL_MASSY_PHONE, '')).strip()
        if matched:
            if massy_phone in matched_phones and len(matched_phones) == 1:
                color = 'black'
            else:
                color = 'pink'
        else:
            color = 'pink'

        if not matched:
            # Only output mismatches per spec
            output_rows.append({
                'address': massy_addr,
                'phone numbers': massy_phone,
                'color': color
            })

    print("💾 Writing output Excel...")
    df_out = pd.DataFrame(output_rows)
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
        workbook = writer.book
        worksheet = writer.sheets[OUTPUT_SHEET]
        fmt_black = workbook.add_format({'font_color': 'black'})
        fmt_pink = workbook.add_format({'font_color': 'pink'})
        for row_num, val in enumerate(df_out['color'], start=1):
            worksheet.set_row(row_num, None, fmt_black if val == 'black' else fmt_pink)

    print(f"\n✅ Done! Output saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
