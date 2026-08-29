# -*- coding: utf-8 -*-
"""
报关自动化工具 v1.6
- 总表：表头第2行，数据从第3行开始
- 匹配：合同【文件名】<-> 总表【IV&PL】（归一化）
- 关税金额 = 合同 BM
- 总税额   = 合同 PPH（无则 PPN）
"""
import os
import re
import sys
import glob
import traceback
import openpyxl
from openpyxl import load_workbook

try:
    import xlrd
except Exception:
    xlrd = None

# ★★★★★ 修改这里即可 ★★★★★
HEADER_ROW = 2        # 总表表头在第几行
DATA_START_ROW = 3    # 总表数据从第几行开始
# ★★★★★★★★★★★★★★★★★★★★

DEBUG = True

CONTRACT_HEADER_KEYWORDS = [
    "AMOUNT", "DESCRIPTION", "QTY", "QUANTITY", "UNIT", "BM", "PPN", "PPH",
    "TOTAL", "PRICE", "HS CODE", "GROSS", "NET", "TAX", "DUTY", "VAT",
    "发票", "金额", "品名", "件数", "单位", "关税", "增值税",
]


# ==================== 工具函数 ====================
def dbg(box, s):
    if not DEBUG:
        return
    print(s, flush=True)
    if box is not None:
        try:
            box.insert("end", s + "\n")
            box.see("end")
        except Exception:
            pass


def num(x, default=None):
    if x is None or str(x).strip() == "":
        return default
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace(",", "").replace(" ", "").replace("￥", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        return default


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def norm(s):
    """归一化：去空格、换行、括号统一、去 - _，转大写"""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", "", s)
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("-", "").replace("_", "")
    return s.strip().upper()


# ==================== 总表列定位 ====================
def find_header_row(ws, max_check=10):
    best = (0, 1)
    for r in range(1, min(max_check, ws.max_row) + 1):
        row = [cell_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        score = header_score(row)
        if score > best[0]:
            best = (score, r)
    return best[1]


def header_score(row):
    s = " ".join(cell_text(v) for v in row).upper()
    score = 0
    keys = ["IV&PL", "INVOICE", "发票", "关税", "总税", "备注", "件数", "单位",
            "AMOUNT", "DESCRIPTION", "QTY", "BM", "PPN", "PPH"]
    for k in keys:
        if k.replace("&", "") in s or k in s:
            score += 1
    return score


def col_by_names(ws, header_row, name_groups):
    headers = [norm(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    for groups in name_groups:
        for name in groups:
            nn = norm(name)
            if not nn:
                continue
            for i, h in enumerate(headers, start=1):
                if nn in h or h in nn or nn.replace("&", "") in h:
                    return i
    return None


# ==================== 合同读取 ====================
def read_rows_xls(path):
    if xlrd is None:
        raise RuntimeError("未安装 xlrd，无法读取 .xls")
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    return rows


def read_rows_xlsx(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def detect_contract_header(rows):
    """合同表头可能不是第1行（第1行常是公司名），找关键词最多的行"""
    best = (0, 1)
    for i, row in enumerate(rows[:15], start=1):
        score = 0
        for c in [cell_text(v).upper() for v in row]:
            for kw in CONTRACT_HEADER_KEYWORDS:
                if kw in c:
                    score += 1
        if score > best[0]:
            best = (score, i)
    return best[1]


def find_col(headers_norm, *names):
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(headers_norm, start=1):
            if nn in h or h in nn or nn.replace("&", "") in h:
                return i
    return None


def extract_contract(rows):
    hdr = detect_contract_header(rows)
    headers_raw = [cell_text(v) for v in rows[hdr - 1]]
    headers = [norm(v) for v in headers_raw]

    bm_col   = find_col(headers, "BM", "关税", "DUTY", "BM可免")
    ppn_col  = find_col(headers, "PPN", "VAT", "增值税")
    pph_col  = find_col(headers, "PPH", "总税", "总税额", "WHT")
    qty_col  = find_col(headers, "QTY", "件数", "数量")
    unit_col = find_col(headers, "UNIT", "单位", "UOM")

    def collect(col_idx):
        vals = []
        if not col_idx:
            return vals
        for row in rows[hdr:]:
            if col_idx - 1 < len(row):
                v = num(row[col_idx - 1])
                if v is not None:
                    vals.append(v)
        return vals

    bm_vals   = collect(bm_col)
    ppn_vals  = collect(ppn_col)
    pph_vals  = collect(pph_col)
    qty_vals  = collect(qty_col)

    info = {
        "bm":   sum(bm_vals)   if bm_vals else None,
        "ppn":  sum(ppn_vals)  if ppn_vals else None,
        "pph":  sum(pph_vals)  if pph_vals else None,
        "qty":  sum(qty_vals)  if qty_vals else None,
        "unit": "",
        "header_row": hdr,
        "headers": headers_raw,
        "bm_col": bm_col, "ppn_col": ppn_col, "pph_col": pph_col,
        "qty_col": qty_col, "unit_col": unit_col,
    }

    if unit_col:
        for row in rows[hdr:]:
            if unit_col - 1 < len(row):
                u = cell_text(row[unit_col - 1]).strip()
                if u and u not in ("-", "/", ""):
                    info["unit"] = u
                    break

    return info


def read_contract(path, box=None):
    ext = os.path.splitext(path)[1].lower()
    dbg(box, f"  读取合同: {os.path.basename(path)} (ext={ext})")
    try:
        if ext == ".xls":
            rows = read_rows_xls(path)
        elif ext in (".xlsx", ".xlsm"):
            rows = read_rows_xlsx(path)
        else:
            dbg(box, f"  不支持的扩展名: {ext}")
            return None
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None

    dbg(box, f"  行数={len(rows)}")
    info = extract_contract(rows)
    dbg(box,
        f"  表头行={info['header_row']}, "
        f"列: bm={info['bm_col']} ppn={info['ppn_col']} pph={info['pph_col']} "
        f"qty={info['qty_col']} unit={info['unit_col']}")
    dbg(box, f"  → bm={info['bm']} ppn={info['ppn']} pph={info['pph']} qty={info['qty']} unit='{info['unit']}'")
    return info


# ==================== 文件名 <-> IV&PL 匹配 ====================
def contract_name_key(filename):
    """
    975_LCMI20260813_iv FORM E.xls  ->  LCMI20260813
    975_SYMI20260813AU-2_iv (RCEP CHINA).xls -> SYMI20260813AU2
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"(?i)\s*_?iv\b", "", name)          # 去 _iv / iv
    name = re.sub(r"(?i)\s*FORM\s*E", "", name)         # 去 FORM E
    name = re.sub(r"[\(\)（）]", "", name)               # 去括号
    name = re.sub(r"^(975|总表)?\d*[_-]*", "", name, flags=re.I)  # 去 975 前缀
    name = name.strip("_- ")
    return norm(name)


def find_contract_file(folder, ivpl, box=None):
    if not ivpl:
        return None
    key = norm(ivpl)
    if not key:
        return None

    pattern = os.path.join(folder, "*")
    candidates = [
        p for p in glob.glob(pattern)
        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in (".xls", ".xlsx", ".xlsm")
    ]

    exact = None
    contain = None
    for p in candidates:
        k = contract_name_key(p)
        if k == key:
            exact = p
            break
        if contain is None and (key in k or k in key):
            contain = p

    hit = exact or contain
    if hit:
        dbg(box, f"  [匹配] 命中 -> {os.path.basename(hit)}")
        return hit

    # 兜底：原始字符串包含
    raw = re.sub(r"[\s_\-]", "", str(ivpl)).upper()
    for p in candidates:
        base = re.sub(r"[\s_\-()（）]", "", os.path.splitext(os.path.basename(p))[0]).upper()
        if raw and raw in base:
            dbg(box, f"  [匹配] 兜底命中 -> {os.path.basename(p)}")
            return p

    dbg(box, f"  [匹配] 未找到（文件夹内 {len(candidates)} 个文件）")
    return None


# ==================== 主处理 ====================
def process(master_path, folder, out_path, box=None):
    wb = load_workbook(master_path)
    ws = wb.active
    dbg(box, f"\n==== 总表 ====")
    dbg(box, f"Sheet: {ws.title}, 行={ws.max_row}, 列={ws.max_column}")
    dbg(box, f"表头行={HEADER_ROW}, 数据从第{DATA_START_ROW}行开始")

    headers_raw = [cell_text(ws.cell(HEADER_ROW, c).value) for c in range(1, ws.max_column + 1)]
    dbg(box, f"表头: {headers_raw}")

    ivpl_col = col_by_names(ws, HEADER_ROW, [("IV&PL", "INVOICE NO", "发票号", "INVOICE", "合同号")])
    duty_col = col_by_names(ws, HEADER_ROW, [("关税金额", "关税")])
    tax_col  = col_by_names(ws, HEADER_ROW, [("总税额", "总税")])
    remark_col = col_by_names(ws, HEADER_ROW, [("备注",)])
    qty_col  = col_by_names(ws, HEADER_ROW, [("件数", "数量", "QTY")])
    unit_col = col_by_names(ws, HEADER_ROW, [("单位", "UNIT")])

    dbg(box, f"列定位: IV&PL={ivpl_col} 关税={duty_col} 总税={tax_col} "
              f"备注={remark_col} 件数={qty_col} 单位={unit_col}")

    if not ivpl_col:
        dbg(box, "⚠️ 未找到 IV&PL / 发票号 列！请把第2行表头截图发我。")

    processed = skipped = 0
    for r in range(DATA_START_ROW, ws.max_row + 1):
        ivpl = ws.cell(r, ivpl_col).value if ivpl_col else None
        if ivpl is None or str(ivpl).strip() == "":
            continue
        ivpl_s = cell_text(ivpl).strip()
        dbg(box, f"\n--- 第{r}行 发票号='{ivpl_s}' ---")

        cp = find_contract_file(folder, ivpl_s, box)
        if not cp:
            if remark_col:
                ws.cell(r, remark_col).value = "未找到合同"
            skipped += 1
            continue

        info = read_contract(cp, box)
        if info is None:
            if remark_col:
                ws.cell(r, remark_col).value = "合同读取失败"
            skipped += 1
            continue

        # ===== 关税金额 = BM =====
        if duty_col and info["bm"] is not None:
            ws.cell(r, duty_col).value = round(info["bm"], 2)

        # ===== 总税额：默认用 PPH；无 PPH 则用 PPN =====
        # ★ 若你要 BM+PPN+PPH，把下面改成：total = (info["bm"] or 0)+(info["ppn"] or 0)+(info["pph"] or 0)
        total = info["pph"]
        if total is None:
            total = info["ppn"]
        if tax_col and total is not None:
            ws.cell(r, tax_col).value = round(total, 2)

        # ===== 备注：单位/件数不一致 =====
        remarks = []
        if qty_col and info["qty"] is not None:
            master_qty = num(ws.cell(r, qty_col).value)
            if master_qty is not None and abs(master_qty - info["qty"]) > 1e-6:
                remarks.append("件数不一致")
        if unit_col and info["unit"]:
            master_unit = cell_text(ws.cell(r, unit_col).value).strip().lower()
            if master_unit and master_unit != info["unit"].strip().lower():
                remarks.append("单位不一致")
        if remark_col:
            old = cell_text(ws.cell(r, remark_col).value)
            new = "; ".join(remarks) if remarks else old
            ws.cell(r, remark_col).value = new

        processed += 1
        dbg(box, f"  ✅ 关税={info['bm']} 总税额={total} {('备注:' + ';'.join(remarks)) if remarks else ''}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v1.6")
    root.geometry("780x640")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=58).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]) or mv.get())
    ).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=58).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(
        filedialog.askdirectory() or fv.get())).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=28, font=("Consolas", 9))
    box.grid(row=3, column=0, columnspan=3, padx=12, pady=10, sticky="nsew")

    def run():
        if not mv.get() or not fv.get():
            return messagebox.showwarning("提示", "先选总表，再选合同文件夹")
        box.delete("1.0", "end")
        try:
            process(mv.get(), fv.get(), os.path.splitext(mv.get())[0] + "_已填写.xlsx", box)
            messagebox.showinfo("完成", "已生成 _已填写.xlsx")
        except Exception:
            box.insert("end", traceback.format_exc())
            messagebox.showerror("报错", "请看日志")

    tk.Button(root, text="③ 开始处理", command=run, bg="#1f6feb", fg="white",
              font=("Microsoft YaHei", 10, "bold"), width=18).grid(row=2, column=1, pady=6)
    root.grid_rowconfigure(3, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        process(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "out.xlsx")
    else:
        gui()
