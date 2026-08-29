# -*- coding: utf-8 -*-
"""
报关自动化工具 v2.3（容错增强 · 终版）
========================================
规则（锁定）：
1. 匹配：合同【文件名】 <-> 总表【IV&PL】（归一化，忽略 975_ / _iv / FORM E / 括号 / 空格 / - _）
2. 关税金额 = 合同「BM可免」税额（BM 税率 × AMOUNT；若"免/免征"则 BM=0）
3. 增值税金额 = 合同「PPN」税额（PPN 税率 × AMOUNT）
4. 总税额   = 去掉 PPH = BM税额 + PPN税额（+ 交叉项 BM×PPN，与合同"税额"列对齐）
5. 件数 / 单位：不校验

增强（v2.3）：
- 读取失败 / 单个合同异常 → 只记日志 + 写备注，不中断整个流程
- .xls 实为 xlsx 时自动兜底重试
- 表头定位失败时安全降级，变量必初始化（修 UnboundLocalError）
- 支持子文件夹递归扫描合同
========================================
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
REMOVE_PPH = True      # True: 总税 = BM + PPN（去掉 PPH）；False: 总税 = BM + PPN + PPH
RECURSIVE = True       # True: 递归扫描子文件夹找合同
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
    """转数字；识别 '免/免征/none/na' 为 0；识别百分比符号。返回 (value, is_percent)"""
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


def to_rate(x):
    """把税率单元格统一转成小数比率。值 > 1 视为百分比数值（需 /100），否则已是小数。"""
    val, is_pct = to_num(x)
    if val is None:
        return None
    if is_pct or (val > 1 and val < 100):
        return val / 100.0
    return val


# ==================== 总表列定位 ====================
def header_score(row):
    s = " ".join(cell_text(v) for v in row).upper()
    keys = ["IV&PL", "INVOICE", "发票", "关税", "增值税", "总税", "备注",
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
                # ★ 跳过空表头，避免 '' in 'AMOUNT' == True 误命中
                if h == "" or h is None:
                    continue
                if nn in h or h in nn or nn.replace("&", "") in h:
                    used_cols.add(i)
                    return i
    return None


# ==================== 合同读取 ====================
def read_rows(path):
    """按扩展名读取；.xls 报 'not supported' 时兜底按 xlsx 重试。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        if xlrd is None:
            raise RuntimeError("需要 xlrd 读取 .xls（pip install xlrd==1.2.0）")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    # .xlsx / .xlsm
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def read_rows_safe(path):
    """
    带兜底读取：
    - .xls 实为 xlsx（报 'not supported'）→ 按 xlsx 重试
    - .xlsx 损坏（报 'no valid workbook part'）→ 尝试用 xlrd / 二进制修复重试
    任何失败都抛异常，由上层记日志、不中断。
    """
    try:
        return read_rows(path)
    except Exception as e:
        msg = str(e).lower()
        ext = os.path.splitext(path)[1].lower()
        # .xls 实为 xlsx
        if ext == ".xls" and "not supported" in msg:
            try:
                return _read_as_xlsx(path)
            except Exception:
                pass
        # .xlsx 报 "no valid workbook part"：尝试修复（解压后重建）
        if ext in (".xlsx", ".xlsm") and "no valid workbook" in msg:
            try:
                return _read_xlsx_repair(path)
            except Exception:
                pass
        raise


def _read_xlsx_repair(path):
    """对损坏的 xlsx：解压 sharedStrings + sheet 用正则提取文本单元格，拼成二维表。"""
    import zipfile, re as _re
    with zipfile.ZipFile(path, "r") as z:
        # 找最大的 sheet XML（即主表）
        sheet_names = [n for n in z.namelist() if _re.match(r"xl/worksheets/sheet\d+\.xml", n)]
        if not sheet_names:
            raise RuntimeError("无可修复的工作表")
        sheet_xml = z.read(sorted(sheet_names)[0]).decode("utf-8", "ignore")
    # 提取所有 <v>值</v>（简化：按行分组）
    rows_data = []
    for row_match in _re.finditer(r"<row\b.*?</row>", sheet_xml, _re.S):
        row_xml = row_match.group(0)
        vals = _re.findall(r"<v>([^<]*)</v>", row_xml)
        # 尝试从 <is>/<t> 取共享字符串索引，简化：直接用数字
        rows_data.append(vals)
    # 若全是数字且能解析，转 float；否则保留字符串
    result = []
    for row in rows_data:
        converted = []
        for v in row:
            try:
                converted.append(float(v))
            except ValueError:
                converted.append(v)
        result.append(converted)
    if not result:
        raise RuntimeError("修复后无数据")
    return result


def _read_as_xlsx(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def find_table_header(rows, max_check=20):
    """找税率表头行（含 BM/PPN/PPH 的行）；找不到返回默认第6行。"""
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        has_bm = "BM" in texts
        has_ppn_pph = ("PPN" in texts) or ("PPH" in texts)
        if has_bm and has_ppn_pph:
            return i
    for i, row in enumerate(rows[:max_check], start=1):
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["BM", "PPN", "PPH", "合计税率", "税率", "税额"]):
            return i
    return 6  # 默认第6行（与你全部合同一致）


def find_col_in_row(hdr, *names, used=None):
    """在单行表头里按名字找列(1-based)；跳过空表头；used=已占用列。"""
    for n in names:
        nn = norm(n)
        if not nn:
            continue
        for i, h in enumerate(hdr, start=1):
            if used and i in used:
                continue
            if h == "" or h is None:
                continue
            if nn in h or h in nn or nn.replace("&", "") in h:
                return i
    return None


def get_amount(rows):
    """全表最大正数金额（>100 视为金额）。"""
    best = None
    for row in rows:
        for v in row:
            n, _ = to_num(v)
            if n and n > 100 and (best is None or n > best):
                best = n
    return best


def get_grand_amount(rows, header_row, amt_col):
    """找 GRAND TOTAL / 合计 行的金额。"""
    for r in range(header_row + 1, len(rows) + 1):
        row = rows[r - 1]
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["GRAND", "合计", "TOTAL", "小计", "TOTAL CIF"]):
            if amt_col:
                n, _ = to_num(row[amt_col - 1])
                if n and n > 0:
                    return n
            best = None
            for v in row:
                n, _ = to_num(v)
                if n and n > 100 and (best is None or n > best):
                    best = n
            if best:
                return best
    return None


def sum_tax(rows, header_row, rate_col, amount=None):
    """
    某税种税额合计（逐行计算，兼容多商品不同税率）：
      - 若传入 amount：单行场景，税额 = amount × 该列税率
      - 否则：逐行取该行的 AMOUNT（第8列） × 该列税率，再求和
    这样与合同「税额」列算法完全一致（每行独立算，再合计）。
    """
    if not rate_col:
        return 0.0
    amt_col = None
    if amount is None:
        # 找 AMOUNT 列（第8列，即 'UNIT PRICE' 后的 'AMOUNT'）
        hdr = [norm(v) for v in rows[header_row - 1]] if header_row <= len(rows) else []
        amt_col = find_col_in_row(hdr, "AMOUNT", "金额", "CIF")
    total = 0.0
    for r in range(header_row + 1, len(rows) + 1):
        row = rows[r - 1]
        # 跳过合计行
        texts = " ".join(cell_text(v).upper() for v in row)
        if any(k in texts for k in ["GRAND", "合计", "TOTAL"]):
            continue
        rate = to_rate(row[rate_col - 1])
        if not rate:
            continue
        if amount is not None:
            total += amount * rate
        else:
            amt = to_num(row[amt_col - 1])[0] if amt_col else None
            if amt:
                total += amt * rate
    return round(total, 2)


def read_contract(path, box=None):
    dbg(box, f"\n  读取合同: {os.path.basename(path)}")
    try:
        rows = read_rows_safe(path)
    except Exception as e:
        dbg(box, f"  读取失败: {e}")
        return None
    dbg(box, f"  行数={len(rows)}")

    # ★ 所有变量先初始化，任何分支都有值（修 UnboundLocalError）
    hr = 6
    amount = None
    bm = ppn = pph = 0.0
    try:
        hr = find_table_header(rows, max_check=20)
    except Exception:
        hr = 6
    if hr < 1 or hr > len(rows):
        hr = 1

    hdr = [norm(v) for v in rows[hr - 1]]
    dbg(box, f"  税率表头行 = 第{hr}行")
    dbg(box, f"  表头: {[cell_text(v) for v in rows[hr-1]]}")

    used = set()
    bm_col   = find_col_in_row(hdr, "BM可免", "BM", used=used)
    if not bm_col:
        bm_col = find_col_in_row(hdr, "BM", used=used)
    ppn_col  = find_col_in_row(hdr, "PPN", "VAT", "增值税", used=used)
    pph_col  = find_col_in_row(hdr, "PPH", "WHT", used=used)
    amt_col  = find_col_in_row(hdr, "AMOUNT", "金额", "CIF", used=used)

    dbg(box, f"  列: AMOUNT={amt_col} BM={bm_col} PPN={ppn_col} PPH={pph_col}")

    # 合计行金额
    amount = get_grand_amount(rows, hr, amt_col)
    if not amount:
        amount = get_amount(rows)

    # BM：若"BM可免"列数据行含"免/免征"，则 BM=0
    # ★ 统一用逐行算法（amount=None），多商品/单商品都精确对齐合同「税额」列
    if bm_col:
        exempt = False
        for r in range(hr + 1, len(rows) + 1):
            raw = cell_text(rows[r - 1][bm_col - 1])
            if any(k in raw for k in ["免", "征"]) and raw.replace("免", "").replace("征", "").strip() in ("", "0", "0.0"):
                if re.search(r"免|免征|BM可免", raw, re.I):
                    exempt = True
                    break
        if exempt:
            bm = 0.0
        else:
            bm = sum_tax(rows, hr, bm_col)   # 逐行 AMOUNT × BM 税率

    if ppn_col:
        ppn = sum_tax(rows, hr, ppn_col)      # 逐行 AMOUNT × PPN 税率
    if pph_col:
        pph = sum_tax(rows, hr, pph_col)      # 逐行 AMOUNT × PPH 税率

    info = {
        "amount": amount,
        "bm":  round(bm, 2),
        "ppn": round(ppn, 2),
        "pph": round(pph, 2),
    }
    # 完整税额 = BM + PPN + PPH + 交叉项(BM×PPN + BM×PPH)
    # 交叉项：逐行累加 AMOUNT × BM税率 × (PPN税率 + PPH税率)
    cross = 0.0
    if amt_col:
        for r in range(hr + 1, len(rows) + 1):
            row = rows[r - 1]
            texts = " ".join(cell_text(v).upper() for v in row)
            if any(k in texts for k in ["GRAND", "合计", "TOTAL"]):
                continue
            amt = to_num(row[amt_col - 1])[0] if amt_col else None
            if not amt:
                continue
            bm_r  = to_rate(row[bm_col - 1])  if bm_col  else None
            ppn_r = to_rate(row[ppn_col - 1]) if ppn_col else None
            pph_r = to_rate(row[pph_col - 1]) if pph_col else None
            cross += amt * (bm_r or 0) * ((ppn_r or 0) + (pph_r or 0))
    full = round(bm + ppn + pph + cross, 2)
    # 读取合同「税额」列：逐行求和（最可靠，不依赖"合计行"定位）
    # 仅统计含商品明细的数据行（跳过表头行、合计行）
    tax_col = find_col_in_row(hdr, "税额", "TAX", "合计税额")
    contract_total_tax = None
    if tax_col:
        total_tax = 0.0
        for r in range(hr + 1, len(rows) + 1):
            row = rows[r - 1]
            texts = " ".join(cell_text(v).upper() for v in row)
            # 跳过合计行 / 表头残留
            if any(k in texts for k in ["GRAND", "合计", "TOTAL"]):
                continue
            n, _ = to_num(row[tax_col - 1])
            if n and n > 0:
                total_tax += n
        if total_tax > 0:
            contract_total_tax = round(total_tax, 2)
    if contract_total_tax is not None:
        diff = round(contract_total_tax - full, 2)
        info["full_check"] = f"验算={full} 合同税额列={contract_total_tax} 差={diff}"
        dbg(box, f"  → 完整税额(验算)={full}  合同「税额」列合计={contract_total_tax}  差={diff}")
    else:
        info["full_check"] = f"验算={full}"
        dbg(box, f"  → 完整税额(验算)={full}")
    info["contract_total_tax"] = contract_total_tax
    dbg(box, f"  → BM税额={bm}  PPN税额={ppn}  PPH税额={pph}  交叉项={round(cross,2)}")
    return info


def to_rate_cell(rows, header_row, col):
    """取某列在表头行的税率值（用于交叉项计算）。"""
    if not col or header_row < 1 or header_row > len(rows):
        return None
    return to_rate(rows[header_row - 1][col - 1])


# ==================== 文件名 <-> IV&PL 匹配 ====================
def contract_name_key(filename):
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"(?i)\s*_?iv\b", "", name)
    name = re.sub(r"(?i)\s*FORM\s*E", "", name)
    name = re.sub(r"[\(\)（）]", "", name)
    name = re.sub(r"^(975|总表)?\d*[_-]*", "", name, flags=re.I)
    name = name.strip("_- ")
    return norm(name)


def scan_contracts(folder):
    """扫描合同文件（支持递归）。"""
    pattern = os.path.join(folder, "**", "*") if RECURSIVE else os.path.join(folder, "*")
    exts = (".xls", ".xlsx", ".xlsm")
    files = []
    for p in glob.glob(pattern, recursive=RECURSIVE):
        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in exts:
            files.append(p)
    return files


def find_contract_file(folder, ivpl, box=None):
    if not ivpl:
        return None
    key = norm(ivpl)
    if not key:
        return None
    candidates = scan_contracts(folder)
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
    vat_col    = col_by_names(ws, HEADER_ROW, [("增值税金额", "增值税")])
    tax_col    = col_by_names(ws, HEADER_ROW, [("总税额", "总税")])
    remark_col = col_by_names(ws, HEADER_ROW, [("备注",)])

    dbg(box, f"列: IV&PL={ivpl_col} 发票金额={amt_col} 关税={duty_col} 增值税={vat_col} "
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

        # ★ 单个合同异常 → 只记备注，不中断
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

        # ③ 总税额 = 去掉 PPH 后的税额
        # 口径：合同「税额」列合计 − PPH税额（最贴合合同；若无可读合计则用 BM+PPN 估算）
        contract_total_tax = info.get("contract_total_tax")
        if REMOVE_PPH:
            if contract_total_tax:
                total = round(contract_total_tax - (info["pph"] or 0), 2)   # 合同税额合计 − PPH
            else:
                total = round((info["bm"] or 0) + (info["ppn"] or 0), 2)    # 兜底估算
        else:
            total = info.get("contract_total_tax") or round(
                (info["bm"] or 0) + (info["ppn"] or 0) + (info["pph"] or 0), 2)
        if tax_col and total:
            ws.cell(r, tax_col).value = total

        processed += 1
        dbg(box, f"  ✅ 关税(BM)={info['bm']}  增值税(PPN)={info['ppn']}  "
                  f"总税(去PPH)={total}")

    wb.save(out_path)
    dbg(box, f"\n完成！处理 {processed} 行，跳过 {skipped} 行")
    dbg(box, f"输出: {out_path}")


# ==================== GUI ====================
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v2.3（容错增强 · 终版）")
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
            out = os.path.splitext(mv.get())[0] + "_已填写.xlsx"
            process(mv.get(), fv.get(), out, box)
            messagebox.showinfo("完成", f"已生成：\n{out}")
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
        out = sys.argv[3] if len(sys.argv) > 3 else "out.xlsx"
        process(sys.argv[1], sys.argv[2], out)
    else:
        gui()
