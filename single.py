# -*- coding: utf-8 -*-
"""
报关自动化工具 v3.3 终版
=========================
计税口径（用户逐条确认，不可再改）：
  每行 关税(BM可免) = H × M                        H=AMOUNT, M=BM税率
  每行 总税         = H×M + (H + H×M) × O          O=PPN税率，不含 PPH
  代码只填 2 列：关税金额、总税额（逐行求和后再填）
  增值税金额 由总表公式 =(发票金额 + 关税金额) × 0.11 自动算，代码不填
  PPH 全程不参与任何一列

作者：自用工具
"""
import os
import re
import sys
import time
import traceback
from datetime import datetime

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

# ============================ 配置区 ============================
CONTRACT_DIR = r"C:\Users\31966\Desktop\报关自动化工具\合同"   # 合同文件夹，按需改
OUTPUT_DIR   = r"C:\Users\31966\Desktop\报关自动化工具\输出"     # 输出目录，按需改

# 总表目标列（按表头文字匹配，不区分大小写/空格/换行）
COL_DUTY = "关税金额"     # 代码写入
COL_TAX  = "总税额"       # 代码写入
COL_VAT  = "增值税金额"   # 代码不写，总表自带公式 =(H+关税)*0.11
COL_AMT  = "发票金额"     # 若总表该列为空，代码顺手填 AMOUNT 合计；已有则跳过

# 总表定位
MASTER_SHEET = "Sheet2"      # 总表工作表名
HEADER_ROW   = 2             # 表头在第 2 行
FIRST_DATA_ROW = 3           # 数据从第 3 行开始

# 合同表定位
CONTRACT_HEADER_ROW = 6      # 合同税率表头在第 6 行（按日志）
CONTRACT_AMOUNT_COL = "AMOUNT"   # 合同里 H 列
CONTRACT_BM_COL      = "BM"       # 合同里 M 列
CONTRACT_PPN_COL     = "PPN"      # 合同里 O 列
# 注：合同里的 PPH 列即使存在，也完全不读取、不参与计算

VAT_RATE = 0.11   # 总表增值税公式固定税率，若合同有差异改这里

LOG_BOX = None


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    if LOG_BOX:
        try:
            LOG_BOX.configure(state="normal")
            LOG_BOX.insert("end", line + "\n")
            LOG_BOX.see("end")
            LOG_BOX.configure(state="disabled")
        except Exception:
            pass


def norm(s):
    """表头归一化：去空格/换行/全角，转小写"""
    return re.sub(r"[\s\u3000\n\r\t]+", "", str(s or "")).lower()


def to_num(v):
    """安全转数字"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "").replace("$", "").replace("¥", "").replace("%", "")
    if t in ("", "-", "--", "None", "nan", "N/A"):
        return 0.0
    try:
        return float(t)
    except Exception:
        return 0.0


# ============================ 读合同 ============================
def read_contract(path):
    """
    返回 dict:
      bm   = Σ(H×M)                      关税金额
      tax  = Σ(H×M + (H+H×M)×O)          总税额（不含 PPH）
      rows = 商品行数
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    rows = [r for r in rows]

    # 找表头行
    hr = CONTRACT_HEADER_ROW - 1
    header = rows[hr] if len(rows) > hr else []
    idx = {}
    for i, cell in enumerate(header):
        k = norm(cell)
        if not k:
            continue
        if "amount" in k or k in ("amt", "金额"):
            idx.setdefault("AMOUNT", i)
        elif k == "bm" or k.endswith("可免") or "bm可免" in k:
            idx.setdefault("BM", i)
        elif k == "ppn" or "ppn" in k:
            idx.setdefault("PPN", i)

    if "AMOUNT" not in idx:
        wb.close()
        raise ValueError(f"合同表头找不到 AMOUNT 列：{os.path.basename(path)}")

    c_h = idx["AMOUNT"]
    c_m = idx.get("BM", -1)
    c_o = idx.get("PPN", -1)

    # 找合计行（含 GRAND / TOTAL 字样的行，跳过）
    grand = None
    for i in range(hr + 1, len(rows)):
        first = str(rows[i][0] or "") if rows[i] else ""
        if re.search(r"grand|total|合计|总计", first, re.I):
            grand = i
            break

    sum_bm = 0.0
    sum_tax = 0.0
    cnt = 0

    for r in range(hr + 1, len(rows)):
        if grand is not None and r == grand:
            continue
        row = rows[r]
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue

        h = to_num(row[c_h] if c_h < len(row) else None)
        if h <= 0:
            continue
        m = to_num(row[c_m]) if c_m >= 0 and c_m < len(row) else 0.0
        o = to_num(row[c_o]) if c_o >= 0 and c_o < len(row) else 0.0

        bm_row = h * m                                  # H×M
        tax_row = bm_row + (h + bm_row) * o             # H×M + (H+H×M)×O  ← 不含 PPH

        sum_bm += bm_row
        sum_tax += tax_row
        cnt += 1

    wb.close()

    return {
        "bm": round(sum_bm, 2),
        "tax": round(sum_tax, 2),
        "rows": cnt,
        "file": os.path.basename(path),
    }


# ============================ 主流程 ============================
def run(master_path):
    log("=" * 58)
    log(f"开始处理 → {os.path.basename(master_path)}")

    if not os.path.isdir(CONTRACT_DIR):
        raise RuntimeError(f"合同文件夹不存在：{CONTRACT_DIR}")
    if not os.path.isdir(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1) 扫描合同（只吃 xls/xlsx，跳过损坏文件）
    contracts = []
    for fn in sorted(os.listdir(CONTRACT_DIR)):
        if not fn.lower().endswith((".xlsx", ".xls")):
            continue
        contracts.append(os.path.join(CONTRACT_DIR, fn))
    log(f"合同文件夹内有效文件 = {len(contracts)} 个")

    # 2) 打开总表
    wb = load_workbook(master_path, data_only=False)
    if MASTER_SHEET not in wb.sheetnames:
        raise RuntimeError(f"总表找不到工作表「{MASTER_SHEET}」，现有：{wb.sheetnames}")
    ws = wb[MASTER_SHEET]

    # 定位目标列
    hdr = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(HEADER_ROW, c).value
        hdr[norm(v)] = c
    log(f"总表表头(第{HEADER_ROW}行)列映射完成")

    col_duty = hdr.get(norm(COL_DUTY))
    col_tax  = hdr.get(norm(COL_TAX))
    col_vat  = hdr.get(norm(COL_VAT))
    col_amt  = hdr.get(norm(COL_AMT))

    if not col_duty:
        raise RuntimeError(f"总表第{HEADER_ROW}行找不到表头「{COL_DUTY}」")
    if not col_tax:
        raise RuntimeError(f"总表第{HEADER_ROW}行找不到表头「{COL_TAX}」")
    log(f"写入列 → 关税=第{col_duty}列  总税=第{col_tax}列  增值税=第{col_vat}列(仅公式,不写值)")

    # 3) 发票号列：自动找含 发票/INVOICE/invoice 的表头
    inv_col = None
    for k, c in hdr.items():
        if "发票" in k or "invoice" in k or "inv" == k:
            inv_col = c
            break
    if not inv_col:
        # 兜底：第 A 列
        inv_col = 1
        log("⚠ 未识别到发票号表头，默认取第 1 列作为发票号列")
    else:
        log(f"发票号列 = 第{inv_col}列")

    # 4) 逐行匹配
    hit = miss = 0
    stamp = datetime.now().strftime("%m%d_%H%M%S")

    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        inv = ws.cell(r, inv_col).value
        if inv is None or str(inv).strip() == "":
            continue
        inv_s = str(inv).strip()

        # 合同名里含发票号即命中（如 975_IWIP20260813KZH_iv.xls）
        matched = None
        for cp in contracts:
            if inv_s.upper() in os.path.basename(cp).upper():
                matched = cp
                break

        if not matched:
            miss += 1
            log(f"  第{r}行 发票号={inv_s} → ❌ 未找到合同")
            continue

        try:
            info = read_contract(matched)
        except Exception as e:
            miss += 1
            log(f"  第{r}行 发票号={inv_s} → ⚠ 读合同失败：{e}")
            continue

        # ★ 只写两列：关税、总税
        ws.cell(r, col_duty).value = info["bm"]
        ws.cell(r, col_tax).value  = info["tax"]
        if col_amt:
            # 发票金额列若为空，补 AMOUNT 合计（仅兜底，已有值不动）
            if ws.cell(r, col_amt).value in (None, ""):
                ws.cell(r, col_amt).value = round(info["bm"] + (info["tax"] - info["bm"]), 2)

        hit += 1
        log(f"  第{r}行 {inv_s} → ✅ 命中 {info['file']} | 商品{info['rows']}行 | "
            f"关税={info['bm']:,.2f} | 总税={info['tax']:,.2f}")

    # 5) 保存（时间戳命名，避开 Excel 占用导致的 PermissionError）
    out_name = f"总表_已填写_{stamp}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    wb.save(out_path)
    wb.close()

    log("-" * 58)
    log(f"✅ 完成：命中 {hit} 行，未匹配 {miss} 行")
    log(f"✅ 输出文件 → {out_path}")
    log(f"⚠ 提示：增值税列请确认总表已设公式 =(发票金额列 + 关税列) * {VAT_RATE}")
    log("=" * 58)
    return out_path


# ============================ GUI ============================
class App:
    def __init__(self, root):
        global LOG_BOX
        self.root = root
        root.title("报关自动化工具 v3.3")
        root.geometry("900x560")

        tk.Label(root, text="报关自动化 · 关税/总税自动填写", font=("Microsoft YaHei", 13, "bold")).pack(pady=8)

        frm = tk.Frame(root)
        frm.pack(fill="x", padx=12)
        self.path_var = tk.StringVar(value="（选总表 xlsx）")
        tk.Entry(frm, textvariable=self.path_var, width=80, font=("Microsoft YaHei", 9)).pack(side="left", padx=(0, 8))
        tk.Button(frm, text="选择总表", command=self.pick).pack(side="left")
        tk.Button(frm, text="▶ 开始运行", command=self.start, bg="#1677ff", fg="white",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left", padx=8)

        tk.Label(root, text="运行日志：", font=("Microsoft YaHei", 9), anchor="w").pack(fill="x", padx=14, pady=(10, 0))

        LOG_BOX = tk.Text(root, state="disabled", font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4")
        LOG_BOX.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        self.master = None

    def pick(self):
        p = filedialog.askopenfilename(title="选总表", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            self.path_var.set(p)
            self.master = p

    def start(self):
        if not self.master:
            messagebox.showwarning("提示", "请先选择总表文件")
            return
        threading.Thread(target=self._run, args=(self.master,), daemon=True).start()

    def _run(self, p):
        try:
            out = run(p)
            messagebox.showinfo("完成", f"处理完成！\n\n{out}")
        except Exception as e:
            log(traceback.format_exc())
            messagebox.showerror("出错", f"{type(e).__name__}：{e}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    # 命令行模式：拖一个总表路径进来也能跑
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        run(sys.argv[1])
    else:
        main()
