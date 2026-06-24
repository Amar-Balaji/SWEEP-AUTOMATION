# Pure standard-library .xlsx reader for the Sweep Script.
#
# Why this exists: the MaxScript path reads Excel via OLE (a running
# Excel.exe), which is fragile and needs Excel installed. This reads the
# .xlsx directly (it is just a ZIP of XML), using ONLY the Python standard
# library -- no pip install, so it works on a locked-down machine.
#
# How it is called from MaxScript (see readExcelLinesPython in sweep_script.ms):
#   sweep_xlsx_path = r'''<full path to the .xlsx>'''
#   sweep_out_path  = r'''<full path to a temp .txt to write>'''
#   exec(open(r'''<this file>''').read())
# It reads worksheet 1, joins each row's non-empty cells with a space, and
# writes one row per line (UTF-8) to sweep_out_path -- the same shape the
# OLE reader produced.

import zipfile
import re


def _col_index(ref):
    # 'B12' -> 0-based column index (B -> 1).
    m = re.match(r'([A-Za-z]+)', ref)
    letters = (m.group(1) if m else 'A').upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _unescape(s):
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&apos;', "'")
             .replace('&amp;', '&'))


def read_xlsx(path):
    rows_out = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        # Shared strings table (most text cells reference this by index).
        shared = []
        if 'xl/sharedStrings.xml' in names:
            data = z.read('xl/sharedStrings.xml').decode('utf-8', 'ignore')
            for si in re.findall(r'<si\b[^>]*>(.*?)</si>', data, re.S):
                texts = re.findall(r'<t\b[^>]*>(.*?)</t>', si, re.S)
                shared.append(_unescape(''.join(texts)))

        # First worksheet.
        sheet = 'xl/worksheets/sheet1.xml'
        if sheet not in names:
            cands = sorted(n for n in names
                           if n.startswith('xl/worksheets/') and n.endswith('.xml'))
            if not cands:
                return rows_out
            sheet = cands[0]
        data = z.read(sheet).decode('utf-8', 'ignore')

        for row_xml in re.findall(r'<row\b[^>]*>(.*?)</row>', data, re.S):
            cells = {}
            for cm in re.finditer(r'<c\b([^>]*)>(.*?)</c>', row_xml, re.S):
                attrs, body = cm.group(1), cm.group(2)
                rm = re.search(r'r="([A-Za-z]+)\d+"', attrs)
                col = _col_index(rm.group(1)) if rm else len(cells)
                tm = re.search(r't="([^"]+)"', attrs)
                ctype = tm.group(1) if tm else ''

                val = ''
                if ctype == 's':
                    vm = re.search(r'<v>(.*?)</v>', body, re.S)
                    if vm:
                        i = int(vm.group(1))
                        if 0 <= i < len(shared):
                            val = shared[i]
                elif ctype in ('inlineStr', 'str'):
                    val = _unescape(''.join(re.findall(r'<t\b[^>]*>(.*?)</t>', body, re.S)))
                    if not val:
                        vm = re.search(r'<v>(.*?)</v>', body, re.S)
                        if vm:
                            val = _unescape(vm.group(1))
                else:
                    vm = re.search(r'<v>(.*?)</v>', body, re.S)
                    if vm:
                        val = _unescape(vm.group(1))

                val = val.strip()
                if val:
                    cells[col] = val

            if cells:
                line = ' '.join(cells[k] for k in sorted(cells)).strip()
                if line:
                    rows_out.append(line)
    return rows_out


def _main():
    g = globals()
    inp = g.get('sweep_xlsx_path', '')
    outp = g.get('sweep_out_path', '')
    rows = []
    try:
        rows = read_xlsx(inp)
    except Exception:
        rows = []
    try:
        with open(outp, 'w', encoding='utf-8') as f:
            for r in rows:
                f.write(r + '\n')
    except Exception:
        pass


_main()
