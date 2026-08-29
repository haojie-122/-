# -*- coding: utf-8 -*-
"""
报关自动化工具 v2.2（基于11个真实合同的精确解析）
====================================================================
【已彻底搞清的合同结构】（所有合同统一）
--------------------------------------------------------------------
第6行是税率表头：
  NO | DESC | BRAND | MODEL | QTY | _ | UNIT_PRICE | AMOUNT | _
  | 中文品名 | HS(CHINA) | HS(INDONESIA)
  | BM(税率) | BM可免(BM税额) | PPN(税率) | PPH(税率) | 合计税率 | 税额 | 监管条件
  (列13)      (列14)           (列15)     (列16)     (列17)    (列18)  (列19)
  索引:         12      13       14       15         16        17

关键规律（逐行，差=0 精确匹配所有合同）：
  税额(col18) = AMOUNT × [ BM + PPN + PPH + BM×(PPN+PPH) ]
             = BM税额 + PPN税额 + PPH税额 + 交叉项(BM×PPN + BM×PPH)
  其中 BM可免(col14) = AMOUNT × BM（若该商品 BM=0 则 BM可免=0）

【填写规则（用户确认）】
  总表「关税金额」 = Σ(BM税额) = 合计行 col14
  总表「总税额」   = 去掉 PPH（去掉 PPH税额 + BM×PPH 交叉项）
                   = Σ[ BM税额 + PPN税额 + 交叉项(BM×PPN) ]
                   = 逐行精确计算后减去所有 PPH 相关部分
  件数/单位：不校验
====================================================================
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
REMOVE_PPH = True   # True: 总税去掉PPH（用户规则）；False: 总税=合计税额


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
    """转成 float；含'免/免征/none'返回 0.0；识别百分比符号"""
    if x is None or str(x).strip() == "":
        return None
    s = str(x).replace(",", "").replace(" ", "").replace("％", "%").strip()
    low = s.lower()
    if any(k in low for k in ["免", "none", "na", "n/a"]):
        return 0.0
    try:
        return float(s.replace("%", ""))
    except Exception:
        return None


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
    """
    在总表表头行定位列。★ 跳过空表头单元格，
    否则 '' in '关税' == True 会把空列误当成命中（同 find_col 的坑）。
    """
    headers = [norm(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    used = set()
    for groups in name_groups:
        for name in groups:
            nn = norm(name)
            if not nn:
                continue
            for i, h in enumerate(headers, start=1):
                if i in used:
                    continue
                if not h:        # ← 跳过空表头单元格
                    continue
                if nn in h or h in nn or nn.replace("&", "") in h:
                    used.add(i)
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
    """找税率表头行：含 BM 且含 (PPN 或 PPH)"""
    for i, row in enumerate(rows[:max_check], start=1):
        t = " ".join(cell_text(v).upper() for v in row)
        if "BM" in t and ("PPN" in t or "PPH" in t):
            return i
    return 1


def find_total_row(rows, start):
    """找合计行（GRAND TOTAL / TOTAL CIF AMOUNT）"""
    for i in range(start, len(rows)):
        t = " ".join(cell_text(v).upper() for v in rows[i])
        if "GRAND" in t or "TOTAL CIF AMOUNT" in t or "TOTAL AMOUNT" in t:
            return i
    return None


def find_col(hdr, *names, used=None):
    """
    在表头行按名字找列(1-based)；used=已占用列。
    ★ 关键：跳过空表头单元格——否则 '' in 'AMOUNT' == True，
    会把空列(如第6、9列的间隔列)误当成命中，导致所有列都返回同一列。
    """
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if used and i in used:
                continue
            hn = norm(h)
            if not hn:           # ← 修复：跳过空表头单元格
                continue
            if nn in hn or hn in nn or nn.replace("&", "") in hn:
                return i
    return None


def read_contract(path, box=None):
    """
    精确解析一个合同，返回：
      amount    : 合同总金额（GRAND TOTAL AMOUNT，校验用）
      bm_total  : Σ(BM税额)               → 总表「关税金额」
      total     : 去掉PPH后的总税额         → 总表「总税额」
      tax_total : 合同「税额」列合计（不去掉PPH，校验用）
    """
    dbg(box, f"\n  读取合同: {os.path.basename(path)}")
    try:
        rows = read_rows(path)
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None
    dbg(box, f"  行数={len(rows)}")

    hr = find_table_header(rows, max_check=15)
    hdr = rows[hr - 1]
    hdr_norm = [norm(v) for v in hdr]
    dbg(box, f"  税率表头行=第{hr}行  表头: {[cell_text(v) for v in hdr]}")

    # 定位列（1-based）
    used = set()
    col_amt    = find_col(hdr_norm, "AMOUNT", used=used);  used.add(col_amt) if col_amt else None
    col_bm     = find_col(hdr_norm, "BM", "BM可免", used=used);  used.add(col_bm) if col_bm else None
    col_bmfree = find_col(hdr_norm, "BM可免", "BM", used=used);  used.add(col_bmfree) if col_bmfree else None
    col_ppn    = find_col(hdr_norm, "PPN", "VAT", "增值税", used=used);  used.add(col_ppn) if col_ppn else None
    col_pph    = find_col(hdr_norm, "PPH", "WHT", used=used);  used.add(col_pph) if col_pph else None
    col_tax    = find_col(hdr_norm, "税额", "TAX", used=used);  used.add(col_tax) if col_tax else None

    dbg(box, f"  列: AMOUNT={col_amt} BM={col_bm} BM可免={col_bmfree} PPN={col_ppn} PPH={col_pph} 税额={col_tax}")

    total_row = find_total_row(rows, hr)
    if total_row is not None:
        tr = rows[total_row]
        grand_amount = to_num(tr[col_amt - 1]) if col_amt else None
        tax_col_total = to_num(tr[col_tax - 1]) if col_tax else None
        bmfree_total  = to_num(tr[col_bmfree - 1]) if col_bmfree else None
        dbg(box, f"  合计行=第{total_row+1}行  GRAND_AMOUNT={grand_amount}  col14(BM可免)={bmfree_total}  col18(税额)={tax_col_total}")

    # ===== 逐行精确计算 =====
    bm_total = 0.0     # Σ(BM税额)
    ppn_total = 0.0    # Σ(PPN税额)
    pph_total = 0.0    # Σ(PPH税额)
    cross_bm_ppn = 0.0 # Σ(AMOUNT × BM × PPN)  交叉项
    cross_bm_pph = 0.0 # Σ(AMOUNT × BM × PPH)  交叉项

    end = total_row if total_row is not None else len(rows)
    for i in range(hr, end):
        row = rows[i]
        amt = to_num(row[col_amt - 1]) if col_amt else None
        if not amt:
            continue
        bm  = to_num(row[col_bm - 1])     if col_bm     else 0.0
        ppn = to_num(row[col_ppn - 1])    if col_ppn    else 0.0
        pph = to_num(row[col_pph - 1])    if col_pph    else 0.0
        # 处理"免/免征"→0
        bm  = bm  if bm  is not None else 0.0
        ppn = ppn if ppn is not None else 0.0
        pph = pph if pph is not None else 0.0

        bm_total  += amt * bm
        ppn_total += amt * ppn
        pph_total += amt * pph
        cross_bm_ppn += amt * bm * ppn
        cross_bm_pph += amt * bm * pph

    # 完整税额（校验，应=合计行 col18）
    full_tax = bm_total + ppn_total + pph_total + cross_bm_ppn + cross_bm_pph

    # ★ 关税金额 = BM税额合计
    duty = round(bm_total, 2)

    # ★ 总税额 = 去掉 PPH = BM税额 + PPN税额 + 交叉项(BM×PPN)
    if REMOVE_PPH:
        total = round(bm_total + ppn_total + cross_bm_ppn, 2)
    else:
        total = round(full_tax, 2)

    info = {
        "amount": (grand_amount if (total_row is not None and grand_amount) else amt),
        "bm": duty,
        "ppn": round(ppn_total, 2),
        "pph": round(pph_total, 2),
        "total": total,
        "full_tax": round(full_tax, 2),
    }
    dbg(box, f"  → BM税额={duty}  PPN税额={round(ppn_total,2)}  PPH税额={round(pph_total,2)}  "
              f"交叉项(BM×PPN)={round(cross_bm_ppn,2)}  BM×PPH={round(cross_bm_pph,2)}")
    dbg(box, f"  → 完整税额(验算)={round(full_tax,2)}  {'✓' if (total_row is None or abs((tax_col_total or 0) - full_tax) < 1) else '⚠ 与合计行差异'}")
    dbg(box, f"  → ★ 关税={duty}  总税(去PPH)={total}")
    return info


# ==================== 文件名 <-> IV&PL ====================
def contract_name_key(filename):
    """
    把文件名洗成纯 IV&PL 编号，便于和总表 IV&PL 比对。
    例：975_IFMI20260813KZH_iv(FORM E).xls -> IFMI20260813KZH
        975_SYMI20260813AU-2_iv (RCEP CHINA).xls -> SYMI20260813AU-2
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    # 1) 去括号及括号内所有内容：(FORM E) / (RCEP CHINA) 等
    name = re.sub(r"[\(（][^)）]*[\)）]", "", name)
    # 2) 去 _iv / iv 后缀
    name = re.sub(r"(?i)\s*_?iv\b", "", name)
    # 3) 去 "FORM E" 之类残留
    name = re.sub(r"(?i)\s*FORM\s*E", "", name)
    # 4) 去开头的 975 / 船 等编号前缀
    name = re.sub(r"^(975|总表|船)\d*[_-]*", "", name, flags=re.I)
    # 5) 清理首尾的 _ - 空格
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
    vat_col    = col_by_names(ws, HEADER_ROW, [("增值税金额", "增值税")])
    remark_col = col_by_names(ws, HEADER_ROW, [("备注",)])

    dbg(box, f"列: IV&PL={ivpl_col} 发票金额={amt_col} 关税={duty_col} 增值税={vat_col} 总税={tax_col} 备注={remark_col}")

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

        # ① 关税金额 = BM税额合计
        if duty_col:
            ws.cell(r, duty_col).value = info["bm"]

        # ② 增值税金额 = PPN税额（额外赠送，列存在就填）
        if vat_col and info["ppn"]:
            ws.cell(r, vat_col).value = info["ppn"]

        # ③ 总税额 = 去掉 PPH
        if tax_col and info["total"]:
            ws.cell(r, tax_col).value = info["total"]

        # ④ 金额校验：合同总金额 vs 总表发票金额
        if amt_col:
            master_amt = to_num(ws.cell(r, amt_col).value)
            if master_amt and info["amount"] and abs(master_amt - info["amount"]) > 1:
                if remark_col:
                    old = cell_text(ws.cell(r, remark_col).value)
                    ws.cell(r, remark_col).value = (old + "; " if old else "") + "金额不一致"

        processed += 1
        dbg(box, f"  ✅ 关税(BM)={info['bm']}  增值税(PPN)={info['ppn']}  "
                  f"总税(去PPH)={info['total']}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v2.2（精确税额解析）")
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
        try:
            gui()
        except Exception as e:
            print("无法启动图形界面（可能是当前环境无 tkinter）：", e)
            print("用法: python single.py <总表.xlsx> <合同文件夹> [输出.xlsx]")
