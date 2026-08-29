# -*- coding: utf-8 -*-
"""
报关自动化工具 v2.0（修复 PPN/PPH）
====================================
真实合同表头结构（第6行）：
['NO','DESCRIPTION','BRAND','MODEL','QUANTITY','','UNIT PRICE','AMOUNT','',
 '中文品名','HS CODE(CHINA)','HS CODE(INDONESIA)',
 'BM','BM 可免','PPN','PPH','合计税率','税额','监管条件']
 列14=BM, 15=BM可免, 16=PPN, 17=PPH, 18=合计税率, 19=税额

关键：BM / PPN / PPH 列存的是【税率%】，真实税额 = 税率 × AMOUNT
      "税额"列(第19列) = 合计税额（= BM税+PPN税+PPH税 或 去掉PPH的合计）

规则：
1. 关税金额 = BM 税额 = BM税率 × AMOUNT
2. 总税额   = 去掉 PPH = BM税额 + PPN税额  （REMOVE_PPH=True）
3. 件数/单位：不校验
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

HEADER_ROW = 2
DATA_START_ROW = 3
DEBUG = True

# ★ 总税额是否去掉 PPH：True=BM+PPN（去PPH），False=BM+PPN+PPH（不去）
REMOVE_PPH = True


# ==================== 工具 ====================
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
    """转数字；'免/none'->0；含%->百分比。返回 (value, is_percent)"""
    if x is None or str(x).strip() == "":
        return None, False
    s = str(x).replace(",", "").replace(" ", "").replace("％", "%").strip()
    low = s.lower()
    is_pct = "%" in s
    if any(k in low for k in ["免", "none", "na", "n/a"]):
        return 0.0, False
    try:
        v = float(s.replace("%", ""))
    except Exception:
        return None, False
    return v, is_pct


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
            raise RuntimeError("需要 xlrd 读取 .xls（pip install xlrd）")
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
        has_bm = "BM" in texts
        has_ppn_pph = ("PPN" in texts) or ("PPH" in texts)
        if has_bm and has_ppn_pph:
            return i
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["BM", "PPN", "PPH", "合计税率", "税率"]):
            return i
    return 1


def get_amount(rows):
    """合计金额（CIF/美元/合计 行的最大正数金额）"""
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


def get_tax_cols(hdr_norm):
    """
    一次性定位 BM / BM可免 / PPN / PPH / 合计税率 / 税额 各列（1-based）。
    用 '可免' 优先匹配 BM列，避免 BM 关键字同时命中 BM可免。
    """
    result = {"bm": None, "bm_free": None, "ppn": None, "pph": None,
              "total_rate": None, "total_tax": None}
    for i, h in enumerate(hdr_norm, start=1):
        if not h:
            continue
        if result["bm_free"] is None and ("BM" in h) and ("可免" in h or "免" in h):
            result["bm_free"] = i          # BM可免 优先
        if result["bm"] is None and "BM" in h:
            result["bm"] = i
        if result["ppn"] is None and "PPN" in h:
            result["ppn"] = i
        if result["pph"] is None and "PPH" in h:
            result["pph"] = i
        if result["total_rate"] is None and "合计税率" in h:
            result["total_rate"] = i
        if result["total_tax"] is None and "税额" in h:
            result["total_tax"] = i
    return result


def tax_amount_from_col(rows, header_row, col, amount):
    """
    某税种列 -> 税额。
    规则：
    1) 若该列存在【数字税额】（多行明细或合计行），求和返回；
    2) 若该列是【税率%】，返回 税率 × AMOUNT；
    3) 若该列是 0 / 免（免征），返回 0。
    """
    if not col:
        return None
    # 收集该列所有数字 + 判断是否含税率
    nums = []
    rates = []
    for r in range(header_row, len(rows) + 1):   # 含表头行（税率常写在表头下方或同行）
        if r - 1 >= len(rows):
            break
        row = rows[r - 1]
        if col - 1 >= len(row):
            continue
        v = row[col - 1]
        # 跳过表头行本身（已是关键字）
        if r == header_row:
            continue
        n, is_pct = to_num(v)
        if n is None:
            continue
        if is_pct or (n <= 100 and n > 0 and ("." in str(v) or n < 1)):
            # 看起来像税率（<=100 且带小数或<1），记下来
            if n > 0:
                rates.append(n if is_pct else n)
        else:
            # 明确的大额数字 = 税额
            if n > 0:
                nums.append(n)
    # 优先：有明确税额数字 -> 求和
    if nums:
        return round(sum(nums), 2)
    # 其次：有税率 + AMOUNT -> 税率 × AMOUNT
    if rates and amount:
        return round(amount * rates[0] / 100.0, 2)
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
    hdr_raw = [cell_text(v) for v in rows[hr - 1]]
    hdr_norm = [norm(v) for v in rows[hr - 1]]
    dbg(box, f"  税率表头行 = 第{hr}行")
    dbg(box, f"  表头: {hdr_raw}")

    amount = get_amount(rows)
    cols = get_tax_cols(hdr_norm)
    dbg(box, f"  列定位: BM={cols['bm']} BM可免={cols['bm_free']} "
              f"PPN={cols['ppn']} PPH={cols['pph']} 合计税率={cols['total_rate']} 税额={cols['total_tax']}")

    # BM 税额：若某商品有「BM 可免」标注（免征），则整份合同的 BM=0；
    # 否则用纯「BM」列的税率 × AMOUNT。
    # 判断「BM 可免」是否免征：检查 BM可免 列是否存在"免/0"
    bm = None
    if cols["bm_free"]:
        bm_free_val = None
        for r in range(hr, len(rows) + 1):
            if r - 1 >= len(rows):
                break
            row = rows[r - 1]
            if cols["bm_free"] - 1 < len(row):
                v = row[cols["bm_free"] - 1]
                n, _ = to_num(v)
                if n is not None or (v is not None and str(v).strip() != ""):
                    bm_free_val = v
                    break
        if bm_free_val is not None:
            n, _ = to_num(bm_free_val)
            # 有 BM可免 列且值为"免/none/0" -> 免征
            txt = str(bm_free_val).strip().lower()
            if n == 0 or any(k in txt for k in ["免", "none", "na", "n/a"]):
                bm = 0.0   # BM可免 = 免征 → 关税 0
            else:
                # BM可免 列存的是税率数字 -> 用它算
                bm = tax_amount_from_col(rows, hr, cols["bm_free"], amount)
    if bm is None and cols["bm"]:
        bm = tax_amount_from_col(rows, hr, cols["bm"], amount)

    ppn = tax_amount_from_col(rows, hr, cols["ppn"], amount) if cols["ppn"] else None
    pph = tax_amount_from_col(rows, hr, cols["pph"], amount) if cols["pph"] else None

    # ★ 兜底：若 BM/PPN/PPH 都没解析出，但有"税额"合计列 -> 直接用合计税额
    total_tax = None
    if cols["total_tax"]:
        total_tax = tax_amount_from_col(rows, hr, cols["total_tax"], amount)

    info = {
        "amount": amount,
        "bm": bm,
        "ppn": ppn,
        "pph": pph,
        "total_tax": total_tax,   # 合同"税额"合计列（可能已含/不含PPH）
    }
    dbg(box, f"  → amount={amount}  BM={bm}  PPN={ppn}  PPH={pph}  合计税额列={total_tax}")
    return info


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
        dbg(box, "⚠️ 未找到 IV&PL 列！请把第2行表头截图发我。")

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
            computed = (info["bm"] or 0) + (info["ppn"] or 0)      # 不含 pph
        else:
            computed = (info["bm"] or 0) + (info["ppn"] or 0) + (info["pph"] or 0)

        # ★ 若合同"税额"合计列存在，优先用它（它本身就是合计，需判断是否含PPH）
        total = None
        if info["total_tax"] is not None:
            total = info["total_tax"]   # 合同合计税额列（可能含PPH）
        else:
            total = computed

        if tax_col and total:
            ws.cell(r, tax_col).value = round(total, 2)

        # ③ 件数/单位：不校验

        processed += 1
        dbg(box, f"  ✅ 关税(BM)={info['bm']}  PPN={info['ppn']}  PPH={info['pph']}  "
                  f"合计税额列={info['total_tax']}  → 总税额={round(total, 2)}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v2.0（修复 PPN/PPH）")
    root.geometry("840x700")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=64).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]) or mv.get())
    ).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=64).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())
    ).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=34, font=("Consolas", 9))
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
