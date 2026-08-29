# -*- coding: utf-8 -*-
"""
报关自动化工具 v1.8（最终规则版）
====================================
规则（已按需求锁定）：
1. 匹配：合同【文件名】 <-> 总表【IV&PL】（归一化，忽略 975_ / _iv / FORM E / 括号 / 空格 / - _）
2. 关税金额 = 合同的「BM可免」税额
3. 总税额   = 去掉 PPH 后的税额（即 BM税额 + PPN税额）
4. 件数 / 单位：本次不做校验
====================================
总表：表头第2行，数据从第3行开始（HEADER_ROW / DATA_START_ROW 可调）
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

# ★★★★★ 如需调整，只改这里 ★★★★★
HEADER_ROW = 2        # 总表表头在第几行
DATA_START_ROW = 3    # 总表数据从第几行开始
# ★★★★★★★★★★★★★★★★★★★★★★★★★

DEBUG = True


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
    """把单元格转成数字；自动剔除 % / 免 / 免征 / 非数字，仅返回正数金额"""
    if x is None or str(x).strip() == "":
        return default
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace(",", "").replace(" ", "").replace("￥", "").replace("$", "")
    s = s.replace("USD", "").replace("美元", "").replace("%", "").strip()
    # 含"免/免征/none/na"视为无税额
    low = s.lower()
    if any(k in low for k in ["免", "none", "na", "n/a"]):
        return default
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
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[_\-]", "", s)
    return s.strip().upper()


# ==================== 总表列定位 ====================
def header_score(row):
    s = " ".join(cell_text(v) for v in row).upper()
    keys = ["IV&PL", "INVOICE", "发票", "关税", "总税", "备注",
            "AMOUNT", "DESCRIPTION", "QTY", "BM", "PPN", "PPH", "金额"]
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
def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        if xlrd is None:
            raise RuntimeError("未安装 xlrd，无法读 .xls")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    raise RuntimeError(f"不支持的扩展名: {ext}")


def find_data_start(rows, max_check=15):
    """找商品/税率表头行（含 DESCRIPTION / 金额 / BM / PPN 等关键字的那行）"""
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["DESCRIPTION", "品名", "商品", "AMOUNT", "金额",
                                     "BM", "PPN", "PPH", "合计", "TOTAL"]):
            return i
    return 1


def find_col_index(rows, header_row, *names):
    """在指定行按关键字找真实列索引（1-based）"""
    if header_row < 1 or header_row > len(rows):
        return None
    hdr = [norm(v) for v in rows[header_row - 1]]
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if nn in h or h in nn or nn.replace("&", "") in h:
                return i
    return None


def get_amount(rows):
    """合同总金额 = 找「合计 / TOTAL / CIF / 美元」行里的最大金额数字"""
    best = None
    for row in rows:
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["合计", "TOTAL", "CIF", "美元", "GRAND", "小计", "SUBTOTAL"]):
            for v in row:
                n = num(v)
                if n and (best is None or n > best):
                    best = n
    if best:
        return best
    # 兜底：全表最大正数金额
    for row in rows:
        for v in row:
            n = num(v)
            if n and n > 100 and (best is None or n > best):
                best = n
    return best


def get_tax_amount(rows, header_row, *keywords):
    """
    取税额（真实数字）：定位「keywords 命中的列」，
    收集该列所有正数金额求和（多行明细 / 合计行都能兜住），跳过 0 / 免 / 税率%。
    """
    col = find_col_index(rows, header_row, *keywords)
    if not col:
        return None
    nums = []
    for r in range(header_row + 1, len(rows) + 1):
        n = num(rows[r - 1][col - 1])
        if n is not None and n > 0:
            nums.append(n)
    if not nums:
        return None
    return sum(nums) if len(nums) > 1 else nums[-1]


def read_contract(path, box=None):
    dbg(box, f"\n  读取合同: {os.path.basename(path)}")
    try:
        rows = read_rows(path)
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None
    dbg(box, f"  行数={len(rows)}")

    data_start = find_data_start(rows, max_check=15)
    dbg(box, f"  数据/表头行 ≈ 第{data_start}行")

    # ★ 关税 BM：优先按「BM可免」精确找；找不到再试「BM / DUTY」
    bm_col = find_col_index(rows, data_start, "BM可免", "BM")
    if bm_col is None:
        bm_col = find_col_index(rows, data_start, "BM", "DUTY", "关税")
    ppn_col = find_col_index(rows, data_start, "PPN", "VAT", "增值税")
    pph_col = find_col_index(rows, data_start, "PPH", "WHT")   # 仅用于"去掉PPH"

    info = {
        "amount": get_amount(rows),
        "bm":   get_tax_amount(rows, data_start, "BM可免", "BM"),
        "ppn":  get_tax_amount(rows, data_start, "PPN", "VAT", "增值税"),
        "pph":  get_tax_amount(rows, data_start, "PPH", "WHT"),
        "qty":  None,   # 不使用
        "unit": "",     # 不使用
    }
    dbg(box, f"  列: BM列={bm_col} PPN列={ppn_col} PPH列={pph_col}")
    dbg(box, f"  → amount={info['amount']}  BM(可免)={info['bm']}  "
              f"PPN={info['ppn']}  PPH={info['pph']}")
    return info


# ==================== 文件名 <-> IV&PL 匹配 ====================
def contract_name_key(filename):
    """把文件名洗成纯编号，便于和 IV&PL 对比"""
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
    # 兜底：原始编号包含
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

    dbg(box, f"列: IV&PL={ivpl_col} 发票金额={amt_col} 关税={duty_col} "
              f"总税={tax_col} 备注={remark_col}")

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

        # ===== ① 关税金额 = BM可免 税额 =====
        if duty_col and info["bm"] is not None:
            ws.cell(r, duty_col).value = round(info["bm"], 2)

        # ===== ② 总税额 = 去掉 PPH = BM + PPN =====
        # ★ 若你的"总税额"只填 PPN（不含 BM），改成：total = info["ppn"] or 0
        total = (info["bm"] or 0) + (info["ppn"] or 0)
        if tax_col and total:
            ws.cell(r, tax_col).value = round(total, 2)

        # ===== ③ 件数 / 单位：本次不校验（按需求）=====

        processed += 1
        dbg(box, f"  ✅ 关税(BM可免)={info['bm']}  总税额(去PPH)={round(total, 2)}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v1.8（最终规则版）")
    root.geometry("800x660")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=60).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]) or mv.get())
    ).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=60).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())
    ).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=30, font=("Consolas", 9))
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
