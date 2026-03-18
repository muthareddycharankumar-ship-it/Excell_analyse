"""
Excel Data Quality Analyser
Usage: python analyse_excel.py <your_file.xlsx>
"""

import sys
import re
import pandas as pd


def fix_garbled(text):
    try:
        return text.encode('latin-1').decode('utf-8')
    except Exception:
        return text


def has_garbled(text):
    if not isinstance(text, str):
        return False
    return len(re.findall(r'[¥À-ÿ½¼¬¨©®]{2,}', text)) > 0


def has_kannada(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u0C80-\u0CFF]', text))


def fix_header(df):
    """Fix sheets where the real header is buried in row 1 or 2."""
    if 'Unnamed: 1' not in df.columns:
        return df
    keywords = ['district', 'k.no', 'name', 'address', 'number', 'beneficiar']
    for i, row in df.iterrows():
        vals = ' '.join(str(v) for v in row.values if pd.notna(v)).lower()
        if any(k in vals for k in keywords):
            df.columns = [str(v) if pd.notna(v) else f'Col_{j}' for j, v in enumerate(row.values)]
            df = df.iloc[i + 1:].reset_index(drop=True)
            # Drop dummy "1 2 3 4 5" numbering row
            if str(df.iloc[0, 0]).strip() == '1' and str(df.iloc[0, 1]).strip() == '2':
                df = df.iloc[1:].reset_index(drop=True)
            break
    return df


def analyse_sheet(name, df):
    df = fix_header(df)
    df_clean = df.dropna(how='all').reset_index(drop=True)

    total_rows = len(df_clean)
    total_cells = total_rows * len(df_clean.columns)

    # --- Empty cells ---
    null_counts = df_clean.isnull().sum()
    total_nulls = int(null_counts.sum())
    rows_with_nulls = int(df_clean.isnull().any(axis=1).sum())

    # --- Language issues ---
    garbled_indices = []
    kannada_indices = []

    for idx, row in df_clean.iterrows():
        combined = ' '.join(str(v) for v in row.values if pd.notna(v))
        if has_garbled(combined):
            garbled_indices.append(idx)
        elif has_kannada(combined):
            kannada_indices.append(idx)

    # --- Print report ---
    sep = '-' * 60
    print(f'\n{sep}')
    print(f'SHEET: {name}')
    print(sep)

    if total_rows == 0:
        print('  STATUS : COMPLETELY EMPTY — no data in this sheet.')
        return

    print(f'  Rows             : {total_rows}')
    print(f'  Columns          : {len(df_clean.columns)}')
    print(f'  Column names     : {list(df_clean.columns)}')

    # Empty cells
    print(f'\n  [EMPTY CELLS]')
    if total_nulls == 0:
        print('  No empty cells found.')
    else:
        print(f'  Total empty cells   : {total_nulls}')
        print(f'  Rows with any empty : {rows_with_nulls}')
        for col in df_clean.columns:
            cnt = int(null_counts[col])
            if cnt > 0:
                pct = round(cnt / total_rows * 100, 1)
                print(f'    Column "{col}": {cnt} empty ({pct}%)')

    # Language issues
    print(f'\n  [LANGUAGE / ENCODING ISSUES]')
    if garbled_indices:
        pct = round(len(garbled_indices) / total_rows * 100, 1)
        print(f'  Garbled encoding rows : {len(garbled_indices)} ({pct}%) — Kannada saved as Latin-1/WinAnsi')
        print()
        print('  GARBLED vs DECODED (all affected rows):')
        print('  ' + '-' * 56)
        for idx in garbled_indices:
            row = df_clean.loc[idx]
            print(f'  Row {idx}:')
            for col, val in row.items():
                if isinstance(val, str) and has_garbled(val):
                    fixed = fix_garbled(val)
                    print(f'    [{col}]')
                    print(f'      Garbled : {val}')
                    print(f'      Decoded : {fixed}')
            print()
    else:
        print('  No garbled encoding detected.')

    if kannada_indices:
        pct = round(len(kannada_indices) / total_rows * 100, 1)
        print(f'  Proper Kannada rows   : {len(kannada_indices)} ({pct}%) — valid Unicode Kannada text')
        print('  Sample Kannada rows:')
        for idx in kannada_indices[:3]:
            row_vals = [str(v) for v in df_clean.loc[idx].values if pd.notna(v)]
            print(f'    Row {idx}: {" | ".join(row_vals[:4])}')
    else:
        print('  No Kannada Unicode text detected.')

    # Summary verdict
    print(f'\n  [VERDICT]')
    issues = []
    if total_nulls > 0:
        issues.append(f'{total_nulls} empty cells')
    if garbled_indices:
        issues.append(f'{len(garbled_indices)} garbled-encoding rows')
    if kannada_indices:
        issues.append(f'{len(kannada_indices)} Kannada-script rows')
    if issues:
        print(f'  Issues found: {", ".join(issues)}')
    else:
        print('  Clean — no issues detected.')


def analyse_sheet_for_report(name, df):
    """Return structured issues for CSV/text report export."""
    df = fix_header(df)
    df_clean = df.dropna(how='all').reset_index(drop=True)
    issues = []

    if len(df_clean) == 0:
        issues.append({
            'sheet': name, 'row': '-', 'column': '-',
            'issue_type': 'EMPTY SHEET', 'value': '-'
        })
        return issues

    # Empty cells
    for idx, row in df_clean[df_clean.isnull().any(axis=1)].iterrows():
        for col in df_clean.columns:
            if pd.isna(row[col]):
                issues.append({
                    'sheet': name,
                    'row': idx,
                    'column': col,
                    'issue_type': 'EMPTY CELL',
                    'value': ''
                })

    # Garbled / Kannada rows
    for idx, row in df_clean.iterrows():
        for col in df_clean.columns:
            val = row[col]
            if isinstance(val, str):
                if has_garbled(val):
                    issues.append({
                        'sheet': name,
                        'row': idx,
                        'column': col,
                        'issue_type': 'GARBLED ENCODING',
                        'value': val
                    })
                elif has_kannada(val):
                    issues.append({
                        'sheet': name,
                        'row': idx,
                        'column': col,
                        'issue_type': 'KANNADA UNICODE',
                        'value': val
                    })
    return issues


def export_csv(all_issues, filepath):
    df = pd.DataFrame(all_issues)
    df.to_csv(filepath, index=False, encoding='utf-8-sig')
    print(f'  CSV report saved  : {filepath}')


def export_txt(all_issues, filepath, source_file):
    lines = []
    lines.append('=' * 70)
    lines.append('EXCEL DATA QUALITY REPORT')
    lines.append(f'Source file : {source_file}')
    lines.append('=' * 70)

    # Group by sheet
    sheets = {}
    for issue in all_issues:
        sheets.setdefault(issue['sheet'], []).append(issue)

    for sheet, issues in sheets.items():
        lines.append(f'\nSHEET: {sheet}')
        lines.append('-' * 70)

        empty_cells = [i for i in issues if i['issue_type'] == 'EMPTY CELL']
        garbled = [i for i in issues if i['issue_type'] == 'GARBLED ENCODING']
        kannada = [i for i in issues if i['issue_type'] == 'KANNADA UNICODE']
        empty_sheet = [i for i in issues if i['issue_type'] == 'EMPTY SHEET']

        if empty_sheet:
            lines.append('  STATUS: COMPLETELY EMPTY SHEET')
            continue

        if empty_cells:
            lines.append(f'\n  EMPTY CELLS ({len(empty_cells)} total):')
            for i in empty_cells:
                lines.append(f'    Row {i["row"]:>4}  |  Column: {i["column"]}')

        if garbled:
            lines.append(f'\n  GARBLED ENCODING ROWS ({len(garbled)} total):')
            for i in garbled:
                lines.append(f'    Row {i["row"]:>4}  |  Column: {i["column"]}')
                lines.append(f'             Value : {i["value"][:80]}')

        if kannada:
            lines.append(f'\n  KANNADA UNICODE ROWS ({len(kannada)} total):')
            for i in kannada:
                lines.append(f'    Row {i["row"]:>4}  |  Column: {i["column"]}')
                lines.append(f'             Value : {i["value"][:80]}')

    lines.append('\n' + '=' * 70)
    lines.append('NOTE: Garbled text uses custom Kannada font encoding (Nudi/Baraha).')
    lines.append('To fix: re-export source data using Unicode (UTF-8) encoding.')
    lines.append('=' * 70)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'  Text report saved : {filepath}')


def main(filepath):
    print('=' * 60)
    print(f'EXCEL DATA QUALITY ANALYSIS')
    print(f'File: {filepath}')
    print('=' * 60)

    try:
        all_sheets = pd.read_excel(filepath, sheet_name=None)
    except Exception as e:
        print(f'ERROR: Could not read file — {e}')
        sys.exit(1)

    print(f'Sheets found: {list(all_sheets.keys())}')

    all_issues = []

    for sheet_name, df in all_sheets.items():
        analyse_sheet(sheet_name, df)
        all_issues.extend(analyse_sheet_for_report(sheet_name, df))

    # Export reports next to the script
    import os
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_name = os.path.basename(filepath).rsplit('.', 1)[0]
    csv_path = os.path.join(script_dir, base_name + '_quality_report.csv')
    txt_path = os.path.join(script_dir, base_name + '_quality_report.txt')

    print('\n' + '=' * 60)
    print('EXPORTING REPORTS')
    print('=' * 60)
    export_csv(all_issues, csv_path)
    export_txt(all_issues, txt_path, filepath)

    print('\n' + '=' * 60)
    print('HOW TO FIX GARBLED TEXT')
    print('=' * 60)
    print('''
  The garbled characters (e.g. "¥ÉæÃgÀuÁ") are Kannada text
  incorrectly saved using Latin-1 / WinAnsi encoding instead
  of Unicode (UTF-8).

  To fix: re-export the source data using Unicode (UTF-8)
  encoding, or contact the data owner to resave the file.
''')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python analyse_excel.py <path_to_excel_file.xlsx>')
        sys.exit(1)
    main(sys.argv[1])