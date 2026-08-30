import os
import sys
import re
import glob
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ===================== 路径配置 =====================
BASE_DIR = r"C:\Users\31966\Desktop\报关自动化工具"
CONTRACT_DIR = os.path.join(BASE_DIR, "合同")
OUTPUT_DIR = os.path.join(BASE_DIR, "输出")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# 合同表头行（你这个例子是第6行）
CONTRACT_HEADER_ROW = 6

# 总表表头行，按你实际；如果不是第2行改这里
MASTER_HEADER_ROW = 2

# 关键字匹配
INVOICE_COL_KEYWORDS = ["发票号", "invoice", "inv"]
AMOUNT_COL_KEYWORDS = ["发票金额", "金额", "amount", "cif", "grand total cif"]
DUTY_COL_KEYWORDS = ["关税金额", "关税", "bm金额", "bm可免合计", "关税(bm)"]
TAX_COL_KEYWORDS = ["总税额", "税额合计", "总税", "合计税额"]
VAT_COL_KEYWORDS = ["增值税金额", "增值税", "vat", "ppn金额"]


def log(msg, log_widget=None):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    if log_widget is not None:
        log_widget.insert(tk.END, line + "\n")
        log_widget.see(tk.END)
        log_widget.update_idletasks()


def normalize_header(x):
    if x is None:
        return ""
    return re.sub(r"\s+", "", str(x)).replace("\n", "").replace("\r", "").replace("\\n", "").lower()


def to_num(x):
    if x is None or x == "":
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "").replace("¥", "").replace("￥", "")
    s = s.replace("USD", "").replace("CNY", "").replace("RMB", "")
    try:
        return float(s)
    except:
        return 0.0


def percent_to_rate(x):
    v = to_num(x)
    s = str(x).strip() if x is not None else ""
    if "%" in s:
        return v / 100.0
    if v > 1:
        return v / 100.0
    return v


def find_file_by_invoice(contract_dir, invoice_no):
    if not invoice_no:
        return None
    key = str(invoice_no).strip()
    patterns = [
        f"*{key}*.xls",
        f"*{key}*.xlsx",
    ]
    found = []
    for root, dirs, files in os.walk(contract_dir):
        for name in files:
            if name.startswith("~$"):
                continue
            low = name.lower()
            if low.endswith((".xls", ".xlsx")):
                if key.lower() in name.lower():
                    found.append(os.path.join(root, name))
    # 优先精确一点：含 iv / invoice 的优先，但不强求
    found.sort(key=lambda p: ("iv" in os.path.basename(p).lower(), p), reverse=True)
    return found[0] if found else None


def read_contract_amount_and_tax(contract_path):
    """
    返回:
    bm_sum, tax_sum, grand_amount, rows_count
    公式：
    每行:
        bm_row = H_amount * M_bm_rate
        tax_row = bm_row + (H_amount + bm_row) * O_ppn_rate
    合计:
        bm_sum = sum(bm_row)
        tax_sum = sum(tax_row)
    不含 PPH。
    """
    wb = load_workbook(contract_path, data_only=True, read_only=True, keep_vba=False)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < CONTRACT_HEADER_ROW:
        raise ValueError(f"合同行数不足，表头行={CONTRACT_HEADER_ROW}，文件={os.path.basename(contract_path)}")

    header = rows[CONTRACT_HEADER_ROW - 1]
    head = [normalize_header(x) for x in header]

    # 你给的表头顺序参考:
    # NO, DESCRIPTION, BRAND, MODEL, QUANTITY, '', UNIT PRICE, AMOUNT,
    # '', 中文品名, HS CODE(CHINA), HS CODE(INDONESIA), BM, BM可免, PPN, PPH, 合计税率, 税额, 监管条件, ''
    # 索引: AMOUNT=7, BM=12, BM可免=13, PPN=14, PPH=15, 税额=17
    # 但用关键字找更稳。
    def idx(keywords):
        for k in keywords:
            kk = normalize_header(k)
            for i, h in enumerate(head):
                if kk and h and kk in h:
                    return i
        return None

    i_amount = idx(["AMOUNT", "金额", "UNIT PRICE"] + AMOUNT_COL_KEYWORDS)
    i_bm = idx(["BM", "关税", "BM税率"] + DUTY_COL_KEYWORDS)
    i_ppn = idx(["PPN", "增值税", "PPN税率"] + VAT_COL_KEYWORDS)
    i_pph = idx(["PPH"])
    i_grand_text = idx(["GRAND TOTAL CIF", "GRAND TOTAL", "TOTAL CIF", "总金额"])

    # 兜底：按你截图固定列
    if i_amount is None:
        i_amount = 7
    if i_bm is None:
        i_bm = 12
    if i_ppn is None:
        i_ppn = 14
    if i_pph is None:
        i_pph = 15

    sum_bm = 0.0
    sum_tax = 0.0
    grand_amount = None

    data_start = CONTRACT_HEADER_ROW
    data_end = len(rows)

    # 找合计行：包含 GRAND TOTAL CIF 或 AMOUNT列附近有合计
    grand_row_idx = None
    for r_idx in range(data_start, len(rows)):
        row = rows[r_idx]
        if not row:
            continue
        txt = " ".join(str(x) for x in row if x is not None)
        if "GRAND TOTAL" in txt.upper() or "TOTAL CIF" in txt.upper():
            grand_row_idx = r_idx
            if i_amount is not None and row[i_amount] is not None:
                grand_amount = abs(to_num(row[i_amount]))
            break

    end_scan = grand_row_idx if grand_row_idx is not None else len(rows)

    item_rows = 0
    for r_idx in range(data_start, end_scan):
        row = rows[r_idx]
        if not row or len(row) == 0:
            continue

        # 跳过合计/空行
        txt = " ".join(str(x) for x in row if x is not None).upper()
        if "GRAND TOTAL" in txt or "TOTAL CIF" in txt:
            continue

        h = to_num(row[i_amount]) if i_amount is not None and i_amount < len(row) else 0.0
        if h <= 0:
            continue

        m = percent_to_rate(row[i_bm]) if i_bm is not None and i_bm < len(row) else 0.0
        o = percent_to_rate(row[i_ppn]) if i_ppn is not None and i_ppn < len(row) else 0.0
        # pph 不用于计算
        _pph = percent_to_rate(row[i_pph]) if i_pph is not None and i_pph < len(row) else 0.0

        bm_row = h * m
        tax_row = bm_row + (h + bm_row) * o

        sum_bm += bm_row
        sum_tax += tax_row
        item_rows += 1

    if grand_amount is None:
        # 没找到合计行就从商品金额求和，或最后一行amount尝试
        last = rows[-1]
        if i_amount is not None and i_amount < len(last):
            grand_amount = abs(to_num(last[i_amount]))
        if grand_amount is None or grand_amount == 0:
            grand_amount = sum(to_num(rows[r][i_amount]) for r in range(data_start, end_scan)
                               if i_amount is not None and i_amount < len(rows[r]))

    return round(sum_bm, 2), round(sum_tax, 2), grand_amount, item_rows


def find_master_columns(ws):
    """
    返回列号（1-based）
    """
    header_row = MASTER_HEADER_ROW
    row = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]

    head = [normalize_header(x) for x in row]

    def find(keywords):
        for k in keywords:
            kk = normalize_header(k)
            for i, h in enumerate(head):
                if kk and h and kk in h:
                    return i + 1
        return None

    col_inv = find(INVOICE_COL_KEYWORDS + ["发票号码"])
    col_amount = find(AMOUNT_COL_KEYWORDS + ["发票金额", "cif金额"])
    col_duty = find(DUTY_COL_KEYWORDS)
    col_tax = find(TAX_COL_KEYWORDS)
    col_vat = find(VAT_COL_KEYWORDS)

    # 兜底：如果表头识别不到，按常见位置；你截图逻辑是 H发票金额, I关税, J增值税, K总税
    if col_amount is None:
        col_amount = 8
    if col_duty is None:
        col_duty = 9
    if col_vat is None:
        col_vat = 10
    if col_tax is None:
        col_tax = 11

    return col_inv, col_amount, col_duty, col_vat, col_tax


def run_master(master_path, contract_dir, log_widget=None):
    log(f"开始处理 → {os.path.basename(master_path)}", log_widget)
    log(f"合同目录 → {contract_dir}", log_widget)

    if not os.path.isdir(contract_dir):
        raise FileNotFoundError(f"合同目录不存在: {contract_dir}")

    files = [f for f in os.listdir(contract_dir) if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$")]
    log(f"合同目录内有效文件 = {len(files)} 个", log_widget)

    wb = load_workbook(master_path)
    ws = wb.active

    col_inv, col_amount, col_duty, col_vat, col_tax = find_master_columns(ws)
    log(f"写入列 → 发票金额=H/{col_amount}, 关税=I/{col_duty}, 增值税=J/{col_vat}(不写值,总表公式自算), 总税=K/{col_tax}", log_widget)
    log(f"发票号列 = {get_column_letter(col_inv) if col_inv else '未识别'}", log_widget)

    hit = 0
    miss = 0
    miss_list = []

    # 从数据行开始，假设第3行起是数据；也可以扫描到最大行
    start_row = MASTER_HEADER_ROW + 1
    max_row = ws.max_row

    for r in range(start_row, max_row + 1):
        inv = ws.cell(r, col_inv).value if col_inv else None
        if inv is None or str(inv).strip() == "":
            continue
        inv = str(inv).strip()

        contract_path = find_file_by_invoice(contract_dir, inv)
        if not contract_path:
            log(f"   第{r}行 {inv} → ❌ 未找到合同", log_widget)
            miss += 1
            miss_list.append(inv)
            continue

        try:
            bm, tax, grand_amount, item_rows = read_contract_amount_and_tax(contract_path)
        except Exception as e:
            log(f"   第{r}行 {inv} → 读取合同失败: {e}", log_widget)
            miss += 1
            miss_list.append(inv)
            continue

        ws.cell(r, col_duty).value = bm
        ws.cell(r, col_tax).value = tax

        # 增值税列：不写值，保留/由总表公式 =(发票金额列+关税列)*0.11 计算
        # 如果原单元格不是公式且你希望清掉，可取消下一行注释：
        # if ws.cell(r, col_vat).value is not None and not str(ws.cell(r, col_vat).value).startswith("="):
        #     ws.cell(r, col_vat).value = None

        log(f"   第{r}行 {inv} → ✅ 命中 {os.path.basename(contract_path)} | 商品行={item_rows} | 关税={bm:.2f} | 总税={tax:.2f}", log_widget)
        hit += 1

    ts = datetime.now().strftime("%m%d_%H%M%S")
    out_path = os.path.join(OUTPUT_DIR, f"总表_已填写_{ts}.xlsx")
    wb.save(out_path)

    log(f"------------------------------------------------------------", log_widget)
    log(f"✅ 完成：命中 {hit} 行，未匹配 {miss} 行", log_widget)
    log(f"✅ 输出文件 → {out_path}", log_widget)
    if miss_list:
        miss_txt = os.path.join(OUTPUT_DIR, f"未匹配发票号_{ts}.txt")
        with open(miss_txt, "w", encoding="utf-8") as f:
            f.write("\n".join(miss_list))
        log(f"⚠ 未匹配清单 → {miss_txt}", log_widget)
    log(f"⚠ 提示：增值税列请保持总表公式 = (发票金额列 + 关税列) * 0.11", log_widget)

    return out_path


# ===================== GUI =====================
class App:
    def __init__(self, root):
        self.root = root
        root.title("报关自动化工具 v3.3")
        root.geometry("900x560")

        self.master_var = tk.StringVar()
        self.contract_var = tk.StringVar(value=CONTRACT_DIR)

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill=tk.X)

        ttk.Label(frm, text="总表 xlsx:").pack(side=tk.LEFT)
        ttk.Entry(frm, textvariable=self.master_var, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm, text="选择总表", command=self.choose_master).pack(side=tk.LEFT, padx=2)
        ttk.Button(frm, text="开始运行", command=self.run).pack(side=tk.LEFT, padx=5)

        frm2 = ttk.Frame(root, padding=(10, 0))
        frm2.pack(fill=tk.X)
        ttk.Label(frm2, text="合同目录:").pack(side=tk.LEFT)
        ttk.Entry(frm2, textvariable=self.contract_var, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm2, text="选择合同目录", command=self.choose_contract_dir).pack(side=tk.LEFT, padx=2)

        ttk.Label(root, text="运行日志:", padding=(10, 8, 0, 0)).pack(anchor=tk.W)
        self.log = tk.Text(root, bg="#1e1e1e", fg="#dcdcdc", font=("Consolas", 10))
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def choose_master(self):
        p = filedialog.askopenfilename(title="选择总表 xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            self.master_var.set(p)

    def choose_contract_dir(self):
        d = filedialog.askdirectory(title="选择合同目录", initialdir=self.contract_var.get() or BASE_DIR)
        if d:
            self.contract_var.set(d)

    def run(self):
        p = self.master_var.get().strip()
        cdir = self.contract_var.get().strip()
        if not p or not os.path.isfile(p):
            messagebox.showwarning("提示", "请先选择总表 xlsx")
            return
        if not cdir or not os.path.isdir(cdir):
            messagebox.showwarning("提示", "请选择有效合同目录")
            return

        self.log.delete("1.0", tk.END)
        t = tk.Thread(target=self._run, args=(p, cdir), daemon=True)
        t.start()

    def _run(self, p, cdir):
        try:
            out = run_master(p, cdir, self.log)
            messagebox.showinfo("完成", f"处理完成！\n\n{out}")
        except Exception as e:
            log(traceback.format_exc(), self.log)
            messagebox.showerror("出错", f"{type(e).__name__}：{e}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        # 命令行：python xxx.py 总表.xlsx
        # 合同目录仍用 GUI/默认 CONTRACT_DIR，或环境变量 CONTRACT_DIR
        cdir = os.environ.get("CONTRACT_DIR", CONTRACT_DIR)
        run_master(sys.argv[1], cdir)
    else:
        main()
