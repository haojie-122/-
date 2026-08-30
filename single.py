# -*- coding: utf-8 -*-
"""
报关自动化工具 v2.4（终版 - 复合税公式）
====================================
真实税额公式（印尼进口税，已由真实合同反向验证，差值=0）：
    每行:
      BM税额  = AMOUNT × BM%
      PPN税额 = (AMOUNT + BM税额) × PPN%      # 基于完税价格 CIF+BM
      PPH税额 = (AMOUNT + BM税额) × PPH%      # 基于完税价格
      该行税额 = BM + PPN + PPH
    汇总 = 所有行求和

总表填值规则：
    关税金额   = Σ(BM税额)   = 合同「BM可免」列(col14)合计
    增值税金额 = Σ(PPN税额)  = 合同「PPN」列(col15)合计
    总税额     = Σ(税额) − Σ(PPH税额)  = 去掉 PPH（BM+PPN+交叉项）
    件数/单位  = 不校验
====================================
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

# ★ 总税额是否去掉 PPH：True = BM+PPN（去PPH），False = BM+PPN+PPH
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
    """转数字；'免/免征/none'→0；识别%符号"""
    if x is None or str(x).strip() == "":
        return 0.0, False
    s = str(x).replace(",", "").replace(" ", "").replace("％", "%").strip()
    low = s.lower()
    is_pct = "%" in s
    if any(k in low for k in ["免", "none", "na", "n/a"]):
        return 0.0, False
    try:
        v = float(s.replace("%", ""))
    except Exception:
        return 0.0, False
    return v, is_pct


def to_rate(x):
    """税率单元格 → 小数比率。值>1视为百分比数值(需/100)，否则已是小数"""
    val, is_pct = to_num(x)
    if val is None:
        return 0.0
    if is_pct or (val > 1 and val < 100):
        return val / 100.0
    return val


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
                if h and (nn in h or h in nn or nn.replace("&", "") in h):
                    used_cols.add(i)
                    return i
    return None


# ==================== 合同读取 ====================
def read_rows(path):
    """按扩展名读取，.xls 用 xlrd，.xlsx 用 openpyxl；自动兜底格式误判"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xls":
            if xlrd is None:
                raise RuntimeError("需要 xlrd 读取 .xls（pip install xlrd==1.2.0）")
            wb = xlrd.open_workbook(path)
            sh = wb.sheet_by_index(0)
            return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    except Exception as e:
        # .xls 实为 xlsx 的兜底
        if ext == ".xls" and "not supported" in str(e).lower():
            try:
                wb = load_workbook(path, data_only=True, read_only=True)
                ws = wb.active
                rows = [list(row) for row in ws.iter_rows(values_only=True)]
                wb.close()
                return rows
            except Exception:
                pass
        raise


def find_table_header(rows, max_check=20):
    """找税率表头行：含 BM 且 (PPN 或 PPH) 的行；找不到则返回靠后的候选行"""
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if "BM" in texts and ("PPN" in texts or "PPH" in texts):
            return i
    # 兜底：找含任一关键字的最靠后行
    last = 1
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["BM", "PPN", "PPH", "税额", "合计税率", "税率"]):
            last = i
    return last


def find_col_in_row(hdr, *names, used=None):
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if used and i in used:
                continue
            if h and (nn in h or h in nn or nn.replace("&", "") in h):
                return i
    return None


def get_grand_amount(rows, header_row, amt_col):
    """合计行(GRAND/TOTAL/合计)的 AMOUNT 值"""
    for r in range(header_row + 1, len(rows) + 1):
        row = rows[r - 1]
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["GRAND", "合计", "TOTAL", "小计"]):
            if amt_col:
                n, _ = to_num(row[amt_col - 1])
                if n and n > 0:
                    return n, r
            best = None
            for v in row:
                n, _ = to_num(v)
                if n and n > 100 and (best is None or n > best):
                    best = n
            if best:
                return best, r
    return None, None


def get_amount_fallback(rows):
    """全表最大金额数字"""
    best = None
    for row in rows:
        for v in row:
            n, _ = to_num(v)
            if n and n > 100 and (best is None or n > best):
                best = n
    return best


def read_contract(path, box=None):
    """读取单个合同，返回 {bm, ppn, pph, total_ex_pph, grand_amount}"""
    dbg(box, f"\n  读取合同: {os.path.basename(path)}")
    try:
        rows = read_rows(path)
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None
    dbg(box, f"  行数={len(rows)}")

    # 定位表头行（提前初始化，任何分支都有值）
    hr = 6
    try:
        hr = find_table_header(rows, max_check=20)
    except Exception:
        hr = 6
    if hr < 1 or hr > len(rows):
        hr = 1

    hdr = [norm(v) for v in rows[hr - 1]]
    dbg(box, f"  税率表头行=第{hr}行")
    dbg(box, f"  表头: {[cell_text(v) for v in rows[hr-1]]}")

    # 列定位（空表头跳过）
    used = set()
    amt_col  = find_col_in_row(hdr, "AMOUNT", "金额", "CIF", used=used)
    bm_col   = find_col_in_row(hdr, "BM可免", "BM", used=used) or find_col_in_row(hdr, "BM", used=used)
    ppn_col  = find_col_in_row(hdr, "PPN", "VAT", "增值税", used=used)
    pph_col  = find_col_in_row(hdr, "PPH", "WHT", used=used)
    tax_col  = find_col_in_row(hdr, "税额", "TAX", "合计税额", used=used)
    dbg(box, f"  列: AMOUNT={amt_col} BM={bm_col} PPN={ppn_col} PPH={pph_col} 税额={tax_col}")

    # 合计金额（优先合计行）
    grand, grand_row = get_grand_amount(rows, hr, amt_col)
    if not grand:
        grand = get_amount_fallback(rows)
    dbg(box, f"  合计行=第{grand_row}行  GRAND_AMOUNT={grand}")

    # ===== 逐行用复合税公式计算 =====
    #   BM  = AMOUNT × BM%
    #   PPN = (AMOUNT + BM) × PPN%
    #   PPH = (AMOUNT + BM) × PPH%
    #   税额 = BM + PPN + PPH
    # 若税额列(tax_col)存在且可读到合计，优先用合同自身的税额列（最权威）
    sum_bm = sum_ppn = sum_pph = sum_tax = 0.0
    data_rows = range(hr + 1, len(rows) + 1)

    # 优先：税额列合计（直接取合同算好的值）
    if tax_col:
        for r in data_rows:
            n, _ = to_num(rows[r - 1][tax_col - 1])
            if n and n > 0:
                sum_tax += n

    # 逐行算 BM/PPN/PPH（用于填 关税/增值税/去PPH）
    for r in data_rows:
        # 跳过合计行
        if r == grand_row:
            continue
        row = rows[r - 1]
        amt, _ = to_num(row[amt_col - 1]) if amt_col else (0.0, False)
        if not amt or amt <= 0:
            continue
        bm_r  = to_rate(row[bm_col - 1])  if bm_col  else 0.0
        ppn_r = to_rate(row[ppn_col - 1]) if ppn_col else 0.0
        pph_r = to_rate(row[pph_col - 1]) if pph_col else 0.0

        # BM可免=免 → 该行为0
        bm_raw = row[bm_col - 1] if bm_col else None
        if bm_raw is not None and str(bm_raw).strip() in ["免", "免征"]:
            bm_r = 0.0

        bm  = amt * bm_r
        ppn = (amt + bm) * ppn_r
        pph = (amt + bm) * pph_r
        sum_bm  += bm
        sum_ppn += ppn
        sum_pph += pph

    # 若税额列无合计值，用复合税求和
    if sum_tax == 0:
        sum_tax = sum_bm + sum_ppn + sum_pph

    # 去 PPH：总税额 = 税额合计 − PPH合计
    if REMOVE_PPH:
        total_ex_pph = sum_tax - sum_pph
    else:
        total_ex_pph = sum_tax

    info = {
        "amount": grand,
        "bm":  round(sum_bm, 2),
        "ppn": round(sum_ppn, 2),
        "pph": round(sum_pph, 2),
        "total_ex_pph": round(total_ex_pph, 2),
    }
    dbg(box, f"  → BM税额={info['bm']}  PPN税额={info['ppn']}  PPH税额={info['pph']}")
    dbg(box, f"  → 税额合计={round(sum_tax,2)}  去PPH={info['total_ex_pph']}")
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
    duty_col   = col_by_names(ws, HEADER_ROW, [("关税金额", "关税")])
    vat_col    = col_by_names(ws, HEADER_ROW, [("增值税金额", "增值税")])
    tax_col    = col_by_names(ws, HEADER_ROW, [("总税额", "总税")])
    remark_col = col_by_names(ws, HEADER_ROW, [("备注",)])

    dbg(box, f"列: IV&PL={ivpl_col} 关税={duty_col} 增值税={vat_col} 总税={tax_col} 备注={remark_col}")

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

        try:
            info = read_contract(cp, box)
        except Exception as e:
            dbg(box, f"  ⚠️ 读取异常: {e}")
            if remark_col:
                ws.cell(r, remark_col).value = "合同读取失败"
            skipped += 1
            continue

        if info is None:
            if remark_col:
                ws.cell(r, remark_col).value = "合同读取失败"
            skipped += 1
            continue

        # ① 关税金额 = BM 税额
        if duty_col and info["bm"] is not None:
            ws.cell(r, duty_col).value = info["bm"]

        # ② 增值税金额 = PPN 税额
        if vat_col and info["ppn"] is not None:
            ws.cell(r, vat_col).value = info["ppn"]

        # ③ 总税额 = 去掉 PPH
        if tax_col and info["total_ex_pph"]:
            ws.cell(r, tax_col).value = info["total_ex_pph"]

        # ④ 件数/单位：不校验

        processed += 1
        dbg(box, f"  ✅ 关税(BM)={info['bm']}  增值税(PPN)={info['ppn']}  "
                  f"总税(去PPH)={info['total_ex_pph']}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v2.4（复合税公式终版）")
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
