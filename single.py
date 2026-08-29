# 报关自动化工具 v1.1 —— 诊断增强版
import os, sys, glob, re, traceback
import openpyxl
from openpyxl import load_workbook

DEBUG = True   # 打印诊断信息到日志框

def dbg(box, s):
    if DEBUG:
        print(s)
        if box: box.insert("end", s + "\n"); box.see("end")

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
    s = (s.replace("sets","set").replace("set","set")
          .replace("pieces","pc").replace("piece","pc").replace("pcs","pc")
          .replace("kgs","kg").replace("cartons","ctn").replace("carton","ctn")
          .replace("units","pc").replace("unit","pc"))
    return s.strip()

# ---------- 读合同 ----------
def read_contract(path):
    ext = path.lower()
    # .xls 转成临时 .xlsx 再读，避免 openpyxl 读不了旧格式
    converted = None
    if ext.endswith(".xls"):
        try:
            from win32com.client import Dispatch   # 需要本机装 pywin32；云端用下面兜底
        except Exception:
            Dispatch = None
        if Dispatch is None:
            # 云端/无Excel环境：用 xlrd 读
            try:
                import xlrd
                return _read_xls_xlrd(path)
            except Exception:
                pass
        # 有Excel则转换
        xl = Dispatch("Excel.Application"); xl.Visible = False
        wb = xl.Workbooks.Open(os.path.abspath(path))
        tmp = os.path.join(os.path.dirname(path), "_tmp_convert.xlsx")
        wb.SaveAs(os.path.abspath(tmp), 51); wb.Close(); xl.Quit()
        converted = tmp
        path = tmp

    wb = load_workbook(path, data_only=True)
    ws = None
    for n in wb.sheetnames:
        if "attach" in n.lower(): ws = wb[n]; break
    if ws is None:
        for n in wb.sheetnames:
            if n.strip().upper() in ("CI","ATTACHMENT 1","ATTACHMENT","IV"):
                ws = wb[n]; break
    if ws is None: ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    info = dict(amount=None, bm=None, ppn_rate=11.0, ppn_amt=None,
                qty=None, unit="", has_pph=False, has_bm=False)

    hi = 0
    for i, row in enumerate(rows[:25]):
        if any("AMOUNT" in txt(c).upper() or "TOTAL" in txt(c).upper() for c in row if c is not None):
            hi = i; break

    head = [txt(c).upper() for c in rows[hi]]
    def col(*keys):
        for k in keys:
            for i, h in enumerate(head):
                if k in h and "PRICE" not in h:
                    return i
        return -1

    i_amt  = col("AMOUNT","TOTAL","SUBTOTAL")
    i_qty  = col("QTY","QUANTITY","PCS")
    i_unit = col("UNIT","UOM")
    i_bm   = col("BM","DUTY")
    i_ppn  = col("PPN","VAT")
    i_pph  = col("PPH","WITHHOLDING")

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
    if converted and os.path.exists(converted):
        try: os.remove(converted)
        except Exception: pass
    return info

def _read_xls_xlrd(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    # 简化：只取前150行，用于匹配金额
    rows = [[sh.cell_value(r,c) for c in range(sh.ncols)] for r in range(min(sh.nrows,150))]
    info = dict(amount=None, bm=None, ppn_rate=11.0, ppn_amt=None, qty=None, unit="", has_pph=False, has_bm=False)
    hi = 0
    for i,row in enumerate(rows[:25]):
        if any("AMOUNT" in txt(c).upper() for c in row): hi=i; break
    head=[txt(c).upper() for c in rows[hi]]
    def col(*keys):
        for k in keys:
            for i,h in enumerate(head):
                if k in h and "PRICE" not in h: return i
        return -1
    i_amt=col("AMOUNT","TOTAL"); i_bm=col("BM","DUTY"); i_ppn=col("PPN"); i_pph=col("PPH")
    i_qty=col("QTY"); i_unit=col("UNIT")
    for row in rows[hi+1:]:
        a=num(row[i_amt]) if i_amt>=0 else 0
        if a: info["amount"]=max(info["amount"] or 0,a)
        b=num(row[i_bm]) if i_bm>=0 else 0
        if b: info["bm"]=b; info["has_bm"]=True
        p=num(row[i_ppn]) if i_ppn>=0 else 0
        if p: info["ppn_amt"]=p
        w=num(row[i_pph]) if i_pph>=0 else 0
        if w: info["has_pph"]=True
        q=num(row[i_qty]) if i_qty>=0 else 0
        if q: info["qty"]=q
        u=txt(row[i_unit]) if i_unit>=0 else ""
        if u: info["unit"]=u
    return info

# ---------- 匹配 ----------
def norm(s):
    return re.sub(r"[^A-Z0-9]", "", txt(s).upper())

def find_contract(folder, inv_key, box=None):
    k = norm(inv_key)
    dbg(box, f"  [匹配] 发票号归一化 = '{k}'")
    if not k: return None
    files = [p for p in glob.glob(os.path.join(folder,"*"))
             if p.lower().endswith((".xls",".xlsx"))]
    dbg(box, f"  [匹配] 文件夹内候选合同数 = {len(files)}")
    # 1) 精确包含
    for p in files:
        n = norm(os.path.basename(p))
        if k in n or n in k:
            dbg(box, f"  [匹配] 命中 -> {os.path.basename(p)}")
            return p
    # 2) 放宽：只要发票号的核心数字串在文件名里（如 20260813AU2）
    core = re.sub(r"[^0-9A-Z]", "", k)
    for p in files:
        n = norm(os.path.basename(p))
        if core and core in n:
            dbg(box, f"  [匹配] 模糊命中 -> {os.path.basename(p)}")
            return p
    dbg(box, "  [匹配] 未找到任何匹配合同")
    return None

# ---------- 主处理 ----------
def process(master_path, folder, out_path, box=None):
    wb = load_workbook(master_path)
    ws = wb.active
    dbg(box, f"总表活动Sheet: {ws.title}, 最大行={ws.max_row}, 最大列={ws.max_column}")

    head = [txt(c) for c in ws[1]]
    dbg(box, f"表头第1行: {head}")
    def C(*keys):
        for k in keys:
            for i, h in enumerate(head):
                hu = h.upper().replace(" ", "")
                if k.upper().replace(" ", "") in hu:
                    dbg(box, f"  列匹配 '{k}' -> 第{i+1}列 (表头='{h}')")
                    return i + 1
        dbg(box, f"  列匹配 '{keys}' -> 未找到")
        return None

    c_inv  = C("IV&PL","INVOICE NO","发票号","INVOICE","合同号")
    c_tax  = C("关税金额","关税")
    c_tot  = C("总税额","总税")
    c_note = C("备注")
    c_qty  = C("件数","数量","QTY")
    c_unit = C("单位","UNIT")

    n_ok = n_skip = 0
    for r in range(2, ws.max_row + 1):
        inv = txt(ws.cell(r, c_inv).value) if c_inv else ""
        if not inv:
            continue
        dbg(box, f"\n--- 第{r}行 发票号='{inv}' ---")
        cp = find_contract(folder, inv, box)
        if not cp:
            if c_note: ws.cell(r, c_note).value = "未找到合同"
            n_skip += 1; continue
        try:
            info = read_contract(cp)
        except Exception as e:
            dbg(box, f"  读合同失败: {e}")
            if c_note: ws.cell(r, c_note).value = "合同读取失败"
            n_skip += 1; continue

        note = ""
        if c_tax and info["bm"]:
            ws.cell(r, c_tax).value = round(info["bm"], 2)
        ppn = info["ppn_amt"]
        if not ppn and info["amount"]:
            ppn = round(info["amount"] * info["ppn_rate"] / 100.0, 2)
        if c_tot and (info["bm"] or ppn):
            ws.cell(r, c_tot).value = round((info["bm"] or 0) + (ppn or 0), 2)
        if c_qty and c_unit and info["qty"]:
            mq = num(ws.cell(r, c_qty).value)
            mu = norm_unit(ws.cell(r, c_unit).value)
            cu = norm_unit(info["unit"])
            if (mq and mq != info["qty"]) or (mu and cu and mu != cu):
                note = "单位不一致"
        if note and c_note:
            ws.cell(r, c_note).value = note
        n_ok += 1
        dbg(box, f"  [OK] 关税={info['bm']} 总税额={(info['bm'] or 0)+(ppn or 0)} {note}")

    wb.save(out_path)
    print(f"\n完成！处理 {n_ok} 行，跳过 {n_skip} 行")
    print(f"输出: {out_path}")

# ---------- GUI ----------
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
    root = tk.Tk()
    root.title("报关自动化工具 v1.1 (诊断版)")
    root.geometry("680x520")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=58).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(filedialog.askopenfilename(
        filetypes=[("Excel","*.xlsx *.xlsm *.xls")]) or mv.get())).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=58).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=22, font=("Consolas", 9))
    box.grid(row=3, column=0, columnspan=3, padx=12, pady=10, sticky="nsew")

    def run():
        if not mv.get() or not fv.get():
            return messagebox.showwarning("提示","先选总表，再选合同文件夹")
        box.delete("1.0","end")
        import io, contextlib
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                process(mv.get(), fv.get(), os.path.splitext(mv.get())[0]+"_已填写.xlsx", box)
            box.insert("end", buf.getvalue())
            messagebox.showinfo("完成","已生成 _已填写.xlsx")
        except Exception:
            box.insert("end", traceback.format_exc())
            messagebox.showerror("报错","看日志")

    tk.Button(root, text="③ 开始处理", command=run, bg="#1f6feb", fg="white",
              font=("Microsoft YaHei",10,"bold"), width=18).grid(row=2, column=1, pady=6)
    root.grid_rowconfigure(3, weight=1); root.grid_columnconfigure(1, weight=1)
    root.mainloop()

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        process(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv)>3 else "out.xlsx")
    else:
        gui()
