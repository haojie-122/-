# -*- coding: utf-8 -*-
"""
报关自动化工具 v1.9（终版 - 税额=金额×税率）
规则：
1. 匹配：合同文件名 <-> 总表 IV&PL（归一化）
2. 关税金额 = 合同「BM」税额
3. 总税额   = BM税额 + PPN税额（去掉 PPH）
4. 件数/单位：不校验
税额计算：先找"税额"列；没有则用 AMOUNT × 税率(BM/PPN/PPH)
"""
import os
import re
import glob
import traceback
import openpyxl
from openpyxl import load_workbook

try:
    import xlrd
except Exception:
    xlrd = None

HEADER_ROW = 2
DATA_START_ROW = 3
DEBUG = True

# ★ 如果你的"税额"是 BM+PPN+PPH（不去PPH），改成 False
REMOVE_PPH = True


# ==================== 工具 ====================
def dbg(box, s):
    if not DEBUG:
        return
    print(s, flush=True)
    if box is not None:
        try:
            box.insert("end", s + "\n"); box.see("end")
        except Exception:
            pass


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[_\-]", "", s)
    return s.strip().upper()


def to_num(x):
    """转数字；识别 '免/免征/none' 为 None；识别百分比"""
    if x is None or str(x).strip() == "":
        return None, False
    s = str(x).replace(",", "").replace(" ", "").replace("％", "%").strip()
    low = s.lower()
    is_pct = "%" in s
    if any(k in low for k in ["免", "none", "na", "n/a"]):
        return 0.0, False   # 免征=0
    try:
        v = float(s.replace("%", ""))
    except Exception:
        return None, False
    if is_pct:
        return v, True        # (数值, 是百分比)
    return v, False


# ==================== 总表列定位 ====================
def header_score(row):
    s = " ".join(cell_text(v) for v in row).upper()
    keys = ["IV&PL", "INVOICE", "发票", "关税", "总税", "备注",
            "AMOUNT", "DESCRIPTION", "BM", "PPN", "PPH", "金额"]
    return sum(1 for k in keys if k.replace("&", "") in s or k in s)


def find_header_row(ws, max_check=10):
    best = (0, 1)
    for r in range(1, min(max_check, ws.max_row) + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        sc = header_score(row)
        if sc > best[0]:
            best = (sc, r)
    return best[1]


def col_by_names(ws, header_row, name_groups):
    headers = [norm(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    # 去重：同一关键字只保留第一个匹配列
    used_cols = set()
    for groups in name_groups:
        for name in groups:
            nn = norm(name)
            if not nn:
                continue
            for i, h in enumerate(headers, start=1):
                if i in used_cols:
                    continue
                if nn in h or h in nn or nn.replace("&", "") in h:
                    used_cols.add(i)
                    return i
    return None


# ==================== 合同读取 ====================
def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        if xlrd is None:
            raise RuntimeError("需要 xlrd 读取 .xls")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def find_table_header(rows, max_check=15):
    """找含 BM/PPN/PPH 税率表的表头行"""
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        # 优先：同时含 BM 和 (PPN 或 PPH) 的行 = 税率表头
        has_bm = "BM" in texts
        has_ppn_pph = ("PPN" in texts) or ("PPH" in texts)
        if has_bm and has_ppn_pph:
            return i
    # 兜底：含任一关键字
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["BM", "PPN", "PPH", "合计税率", "税率"]):
            return i
    return 1


def find_col(rows, header_row, *names, used=None):
    """在 header_row 按名字找列(1-based)，used=已占用的列集合"""
    if header_row < 1 or header_row > len(rows):
        return None
    hdr = [norm(v) for v in rows[header_row - 1]]
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if used and i in used:
                continue
            if nn in h or h in nn or nn.replace("&", "") in h:
                return i
    return None


def get_amount(rows):
    """找合计金额（CIF/美元/合计 行的最大数字）"""
    best = None
    for row in rows:
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["合计", "TOTAL", "CIF", "美元", "GRAND", "小计", "SUBTOTAL"]):
            for v in row:
                n, _ = to_num(v)
                if n and n > 0 and (best is None or n > best):
                    best = n
    if best:
        return best
    for row in rows:
        for v in row:
            n, _ = to_num(v)
            if n and n > 100 and (best is None or n > best):
                best = n
    return best


def get_tax_rate_and_amount(rows, header_row, keyword, used):
    """
    返回 (税额数值, 是否成功)。
    策略：优先找同行右侧的「税额」列；否则取该列的税率% × AMOUNT。
    """
    col = find_col(rows, header_row, keyword, used=used)
    if not col:
        return None
    used.add(col)

    # 1) 先看 header_row 这一行 keyword 列的值（可能是税率数字）
    rate_val, is_pct = to_num(rows[header_row - 1][col - 1])
    if rate_val is None:
        rate_val = 0.0

    # 2) 找「税额」列：优先紧邻右侧，或在表头里含"税额/TAX/金额"且非 BM/PPN/PPH 列
    tax_col = find_col(rows, header_row, "税额", "TAX", "税额金额", used=used)
    if tax_col:
        # 取该列合计行（最靠后）的数字
        for r in range(len(rows), header_row, -1):
            n, _ = to_num(rows[r - 1][tax_col - 1])
            if n and n > 0:
                return n
        # 若合计行无，则求和
        total = 0.0
        for r in range(header_row + 1, len(rows) + 1):
            n, _ = to_num(rows[r - 1][tax_col - 1])
            if n:
                total += n
        if total > 0:
            return total

    # 3) 兜底：税率 × AMOUNT
    amount = get_amount(rows)
    if amount and is_pct and rate_val > 0:
        return round(amount * rate_val / 100.0, 2)
    if amount and not is_pct and rate_val > 0:
        # 像 0.05 这种小数（非百分比形式）也按比率算
        return round(amount * rate_val, 2)
    return None


def read_contract(path, box=None):
    dbg(box, f"\n  读取合同: {os.path.basename(path)}")
    try:
        rows = read_rows(path)
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None
    dbg(box, f"  行数={len(rows)}")

    hr = find_table_header(rows, max_check=15)
    dbg(box, f"  税率表头行 = 第{hr}行")
    dbg(box, f"  表头: {[cell_text(v) for v in rows[hr-1]]}")

    used = set()
    amount = get_amount(rows)

    bm   = get_tax_rate_and_amount(rows, hr, "BM可免",   used) if False else get_tax_rate_and_amount(rows, hr, "BM", used)
    # 重新建 used（上面可能误加），简单起见每个独立计算：
    used_bm = {find_col(rows, hr, "BM可免", "BM", used=None)}
    used_ppn = set()
    used_pph = set()

    # 重新干净地算（避免交叉占用）
    bm  = _calc_tax(rows, hr, ["BM可免", "BM"], used_bm)
    ppn = _calc_tax(rows, hr, ["PPN", "VAT", "增值税"], used_ppn)
    pph = _calc_tax(rows, hr, ["PPH", "WHT"], used_pph)

    info = {
        "amount": amount,
        "bm":  bm,
        "ppn": ppn,
        "pph": pph,
    }
    dbg(box, f"  → amount={amount}  BM={bm}  PPN={ppn}  PPH={pph}")
    return info


def _calc_tax(rows, header_row, keywords, used):
    """独立计算单个税种税额（避免列占用冲突）"""
    col = None
    hdr = [norm(v) for v in rows[header_row - 1]]
    for n in keywords:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if used and i in used:
                continue
            if nn in h or h in nn or nn.replace("&", "") in h:
                col = i
                if used is not None:
                    used.add(i)
                break
        if col:
            break
    if not col:
        return None

    # 找税额列（同行右侧第一个含"税额/TAX/金额"且非税率列）
    tax_col = None
    for i, h in enumerate(hdr, start=1):
        if i <= col or (used and i in used):
            continue
        hu = h
        if any(k in hu for k in ["税额", "TAX", "金额", "合计税额"]):
            tax_col = i
            if used is not None:
                used.add(i)
            break

    if tax_col:
        # 取合计（从后往前第一个正数）
        for r in range(len(rows), header_row, -1):
            n, _ = to_num(rows[r - 1][tax_col - 1])
            if n and n > 0:
                return round(n, 2)
        total = 0.0
        for r in range(header_row + 1, len(rows) + 1):
            n, _ = to_num(rows[r - 1][tax_col - 1])
            if n:
                total += n
        if total > 0:
            return round(total, 2)

    # 兜底：税率 × AMOUNT
    rate, is_pct = to_num(rows[header_row - 1][col - 1])
    amount = None
    # 需要 AMOUNT，从全局取
    amount = _get_amount(rows)
    if rate is None:
        return None
    if amount:
        if is_pct:
            return round(amount * rate / 100.0, 2)
        return round(amount * rate, 2)
    return None


def _get_amount(rows):
    best = None
    for row in rows:
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["合计", "TOTAL", "CIF", "美元", "GRAND", "小计", "SUBTOTAL"]):
            for v in row:
                n, _ = to_num(v)
                if n and n > 0 and (best is None or n > best):
                    best = n
    if best:
        return best
    for row in rows:
        for v in row:
            n, _ = to_num(v)
            if n and n > 100 and (best is None or n > best):
                best = n
    return best


# ==================== 文件名 <-> IV&PL ====================
def contract_name_key(filename):
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"(?i)\s*_?iv\b", "", name)
    name = re.sub(r"(?i)\s*FORM\s*E", "", name)
    name = re.sub(r"[\(\)（）]", "", name)
    name = re.sub(r"^(975|总表)?\d*[_-]*", "", name, flags=re.I)
    name = name.strip("_- ")
    return norm(name)


def find_contract_file(folder, ivpl, box=None):
    if not ivpl:
        return None
    key = norm(ivpl)
    if not key:
        return None
    candidates = [
        p for p in glob.glob(os.path.join(folder, "*"))
        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in (".xls", ".xlsx", ".xlsm")
    ]
    exact = contain = None
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

    ivpl_col   = col_by_names(ws, HEADER_ROW, [("IV&PL", "INVOICE NO", "发票号", "INVOICE", "合同号")])
    amt_col    = col_by_names(ws, HEADER_ROW, [("发票金额", "AMOUNT", "金额")])
    duty_col   = col_by_names(ws, HEADER_ROW, [("关税金额", "关税")])
    tax_col    = col_by_names(ws, HEADER_ROW, [("总税额", "总税")])
    remark_col = col_by_names(ws, HEADER_ROW, [("备注",)])

    dbg(box, f"列: IV&PL={ivpl_col} 发票金额={amt_col} 关税={duty_col} 总税={tax_col} 备注={remark_col}")

    if not ivpl_col:
        dbg(box, "⚠️ 未找到 IV&PL 列！")

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

        # ① 关税金额 = BM 税额
        if duty_col and info["bm"] is not None:
            ws.cell(r, duty_col).value = round(info["bm"], 2)

        # ② 总税额 = 去掉 PPH
        if REMOVE_PPH:
            total = (info["bm"] or 0) + (info["ppn"] or 0)       # 不含 pph
        else:
            total = (info["bm"] or 0) + (info["ppn"] or 0) + (info["pph"] or 0)
        if tax_col and total:
            ws.cell(r, tax_col).value = round(total, 2)

        # ③ 件数/单位：不校验

        processed += 1
        dbg(box, f"  ✅ 关税(BM)={info['bm']}  总税额(去PPH)={round(total, 2)}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v1.9（税额=金额×税率）")
    root.geometry("820x680")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=62).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]) or mv.get())
    ).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=62).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())
    ).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=32, font=("Consolas", 9))
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
