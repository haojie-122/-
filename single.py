# 报关自动化工具 v1.0 —— 免命令行，双击即用
import os, sys, glob, re, traceback
import openpyxl
from openpyxl import load_workbook

# ---------- 小工具 ----------
def num(x, default=0.0):
    if x is None: return default
    s = str(x).strip().replace(",", "")
    if not s: return default
    try: return float(s)
    except: return default

def txt(x):
    return "" if x is None else str(x).strip()

def norm_unit(s):
    s = txt(s).lower()
    s = (s.replace("sets", "set").replace("set", "set")
          .replace("pieces", "pc").replace("piece", "pc").replace("pcs", "pc")
          .replace("kgs", "kg").replace("cartons", "ctn").replace("carton", "ctn")
          .replace("units", "pc").replace("unit", "pc"))
    return s.strip()

# ---------- 读合同（.xls/.xlsx，优先 attachment 页） ----------
def read_contract(path):
    wb = load_workbook(path, data_only=True)
    ws = None
    for n in wb.sheetnames:
        if "attach" in n.lower():
            ws = wb[n]; break
    if ws is None:
        for n in wb.sheetnames:
            if n.strip().upper() in ("CI", "ATTACHMENT 1", "ATTACHMENT"):
                ws = wb[n]; break
    if ws is None:
        ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    info = dict(amount=None, bm=None, ppn_rate=11.0, ppn_amt=None,
                qty=None, unit="", has_pph=False, has_bm=False)

    # 找表头行
    hi = 0
    for i, row in enumerate(rows[:25]):
        if any("AMOUNT" in txt(c).upper() or "TOTAL" in txt(c).upper() for c in row if c is not None):
            hi = i; break

    head = [txt(c).upper() for c in rows[hi]]
    def col(*keys):
        for k in keys:
            for i, h in enumerate(head):
                if k in h and "PRICE" not in h:   # 别把 UNIT PRICE 当 UNIT
                    return i
        return -1

    i_amt = col("AMOUNT", "TOTAL", "SUBTOTAL")
    i_qty = col("QTY", "QUANTITY", "PCS")
    i_unit = col("UNIT", "UOM")
    i_bm   = col("BM", "DUTY")
    i_ppn  = col("PPN", "VAT")
    i_pph  = col("PPH", "WITHHOLDING")

    for row in rows[hi+1:hi+150]:
        a = num(row[i_amt]) if i_amt >= 0 else 0
        if a: info["amount"] = max(info["amount"] or 0, a)
        q = num(row[i_qty]) if i_qty >= 0 else 0
        if q: info["qty"] = q
        u = txt(row[i_unit]) if i_unit >= 0 else ""
        if u: info["unit"] = u
        b = num(row[i_bm]) if i_bm >= 0 else 0
        if b: info["bm"] = b; info["has_bm"] = True
        p = num(row[i_ppn]) if i_ppn >= 0 else 0
        if p: info["ppn_amt"] = p
        w = num(row[i_pph]) if i_pph >= 0 else 0
        if w: info["has_pph"] = True

    wb.close()
    return info

# ---------- 发票号模糊匹配 ----------
def norm(s):
    return re.sub(r"[^A-Z0-9]", "", txt(s).upper())

def find_contract(folder, inv_key):
    k = norm(inv_key)
    if not k: return None
    best = None
    for p in glob.glob(os.path.join(folder, "*")):
        if not p.lower().endswith((".xls", ".xlsx")): continue
        n = norm(os.path.basename(p))
        if k in n or n in k:
            return p
    return best

# ---------- 主处理 ----------
def process(master_path, folder, out_path):
    wb = load_workbook(master_path)
    ws = wb.active

    head = [txt(c).upper() for c in ws[1]]
    def C(*keys):
        for k in keys:
            for i, h in enumerate(head):
                if k in h: return i + 1
        return None

    c_inv  = C("IV&PL", "INVOICE NO", "发票号")
    c_tax  = C("关税金额", "关税")
    c_tot  = C("总税额", "总税")
    c_note = C("备注")
    c_qty  = C("件数", "数量", "QTY")
    c_unit = C("单位", "UNIT")

    n_ok = n_skip = 0
    for r in range(2, ws.max_row + 1):
        inv = txt(ws.cell(r, c_inv).value) if c_inv else ""
        if not inv: continue
        cp = find_contract(folder, inv)
        if not cp:
            if c_note: ws.cell(r, c_note).value = "未找到合同"
            n_skip += 1; continue
        try:
            info = read_contract(cp)
        except Exception:
            if c_note: ws.cell(r, c_note).value = "合同读取失败"
            n_skip += 1; continue

        note = ""
        # 关税金额 = BM可免
        if c_tax and info["bm"]:
            ws.cell(r, c_tax).value = round(info["bm"], 2)

        # 总税额 = 关税 + PPN（已删 PPH）
        ppn = info["ppn_amt"]
        if not ppn and info["amount"]:
            ppn = round(info["amount"] * info["ppn_rate"] / 100.0, 2)
        if c_tot and (info["bm"] or ppn):
            ws.cell(r, c_tot).value = round((info["bm"] or 0) + (ppn or 0), 2)

        # 单位不一致
        if c_qty and c_unit and info["qty"]:
            mq = num(ws.cell(r, c_qty).value)
            mu = norm_unit(ws.cell(r, c_unit).value)
            cu = norm_unit(info["unit"])
            if (mq and mq != info["qty"]) or (mu and cu and mu != cu):
                note = "单位不一致"
        if note and c_note:
            ws.cell(r, c_note).value = note

        n_ok += 1
        print(f"[OK] {inv} -> 关税={info['bm']} 总税额={(info['bm'] or 0)+(ppn or 0)} {note}")

    wb.save(out_path)
    print(f"\n完成！处理 {n_ok} 行，跳过 {n_skip} 行")
    print(f"输出文件：{out_path}")

# ---------- 界面（tkinter，Windows 自带，不用装） ----------
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("报关自动化工具 v1.0")
    root.geometry("640x440")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=58).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]) or mv.get())).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=58).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=16)
    box.grid(row=3, column=0, columnspan=3, padx=12, pady=10, sticky="nsew")

    def run():
        if not mv.get() or not fv.get():
            return messagebox.showwarning("提示", "先选总表，再选合同文件夹")
        box.delete("1.0", "end")
        import io, contextlib
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                process(mv.get(), fv.get(), os.path.splitext(mv.get())[0] + "_已填写.xlsx")
            box.insert("end", buf.getvalue())
            messagebox.showinfo("完成", "已生成 _已填写.xlsx")
        except Exception:
            box.insert("end", traceback.format_exc())
            messagebox.showerror("报错了", "看下面红字/白字日志，把内容发我")

    tk.Button(root, text="③ 开始处理", command=run, bg="#1f6feb", fg="white",
              font=("Microsoft YaHei", 10, "bold"), width=18).grid(row=2, column=1, pady=6)
    root.grid_rowconfigure(3, weight=1); root.grid_columnconfigure(1, weight=1)
    root.mainloop()

if __name__ == "__main__":
    if len(sys.argv) >= 3:          # 高级：命令行也能跑
        process(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "out.xlsx")
    else:
        gui()
