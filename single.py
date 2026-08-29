# 报关自动化工具 v1.4 —— 明确表头=第2行，数据从第3行开始
import os, sys, glob, re, traceback
import openpyxl
from openpyxl import load_workbook

DEBUG = True

# ★★★ 你的总表：表头在第几行？改这里即可（1-based） ★★★
HEADER_ROW = 2
# ★★★ 数据从第几行开始？默认 = HEADER_ROW + 1 = 3 ★★★
DATA_START_ROW = HEADER_ROW + 1

HEADER_KEYWORDS = [
    "IV&PL", "INVOICE", "发票", "合同号", "采购订单", "PO",
    "关税", "增值税", "总税", "税额", "TAX", "DUTY",
    "件数", "数量", "QTY", "QUANTITY",
    "单位", "UNIT", "UOM",
    "备注", "REMARK",
    "AMOUNT", "TOTAL", "DESCRIPTION", "品名",
    "毛重", "净重", "GROSS", "NET", "WEIGHT",
    "体积", "VOLUME", "CBM",
    "BL", "OBL", "LOADING", "PORT",
]

def dbg(box, s):
    if DEBUG:
        print(s)
        if box:
            try: box.insert("end", s + "\n"); box.see("end")
            except Exception: pass

def num(x, default=0.0):
    if x is None: return default
    if isinstance(x, str):
        s = x.strip().replace(",", "")
        if not s: return default
        try: return float(s)
        except: return default
    try: return float(x)
    except: return default

def txt(x):
    if x is None: return ""
    if hasattr(x, "value"):
        return txt(x.value)
    return str(x).strip()

def norm_unit(s):
    s = txt(s).lower()
    s = (s.replace("sets","set").replace("set","set")
          .replace("pieces","pc").replace("piece","pc").replace("pcs","pc")
          .replace("kgs","kg").replace("cartons","ctn").replace("carton","ctn")
          .replace("units","pc").replace("unit","pc"))
    return s.strip()

# ---------- 读合同 ----------
def _pick_sheet(wb):
    ws = None
    for n in wb.sheetnames:
        if "attach" in n.lower(): ws = wb[n]; break
    if ws is None:
        for n in wb.sheetnames:
            if n.strip().upper() in ("CI","ATTACHMENT 1","ATTACHMENT","IV"):
                ws = wb[n]; break
    if ws is None: ws = wb.active
    return ws

def _make_info():
    return dict(amount=None, bm=None, ppn_rate=11.0, ppn_amt=None,
                qty=None, unit="", has_pph=False, has_bm=False)

def _col(head, *keys):
    for k in keys:
        for i, h in enumerate(head):
            hu = h.upper().replace(" ", "")
            ku = k.upper().replace(" ", "")
            if ku in hu and "PRICE" not in hu:
                return i
    return -1

def read_contract(path, box=None):
    ext = path.lower()
    converted = None
    if ext.endswith(".xls"):
        try:
            import xlrd
            return _read_xls_xlrd(path, box)
        except Exception as e:
            dbg(box, f"  xlrd失败: {e}")
            try:
                from win32com.client import Dispatch
                xl = Dispatch("Excel.Application"); xl.Visible = False
                wb = xl.Workbooks.Open(os.path.abspath(path))
                tmp = os.path.join(os.path.dirname(path), "_tmp_convert.xlsx")
                wb.SaveAs(os.path.abspath(tmp), 51); wb.Close(); xl.Quit()
                converted = tmp; path = tmp
            except Exception as e2:
                dbg(box, f"  无法读xls: {e2}")

    wb = load_workbook(path, data_only=True)
    ws = _pick_sheet(wb)

    # 合同表也自动找表头（通常第1行就是）
    head = [txt(ws.cell(1, c).value).upper() for c in range(1, ws.max_column + 1)]
    dbg(box, f"  合同表头: {head}")

    i_amt  = _col(head, "AMOUNT","TOTAL","SUBTOTAL")
    i_qty  = _col(head, "QTY","QUANTITY","PCS")
    i_unit = _col(head, "UNIT","UOM")
    i_bm   = _col(head, "BM","DUTY")
    i_ppn  = _col(head, "PPN","VAT")
    i_pph  = _col(head, "PPH","WITHHOLDING")

    info = _make_info()
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        a = num(vals[i_amt]) if i_amt >= 0 else 0
        if a: info["amount"] = max(info["amount"] or 0, a)
        q = num(vals[i_qty]) if i_qty >= 0 else 0
        if q: info["qty"] = q
        u = txt(vals[i_unit]) if i_unit >= 0 else ""
        if u: info["unit"] = u
        b = num(vals[i_bm]) if i_bm >= 0 else 0
        if b: info["bm"] = b; info["has_bm"] = True
        p = num(vals[i_ppn]) if i_ppn >= 0 else 0
        if p: info["ppn_amt"] = p
        w = num(vals[i_pph]) if i_pph >= 0 else 0
        if w: info["has_pph"] = True

    wb.close()
    if converted and os.path.exists(converted):
        try: os.remove(converted)
        except Exception: pass
    return info

def _read_xls_xlrd(path, box=None):
    import xlrd
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    raw = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(min(sh.nrows, 200))]
    head = [txt(c).upper() for c in raw[0]]
    dbg(box, f"  [xls] 表头: {head}")

    i_amt=_col(head,"AMOUNT","TOTAL"); i_bm=_col(head,"BM","DUTY")
    i_ppn=_col(head,"PPN"); i_pph=_col(head,"PPH")
    i_qty=_col(head,"QTY"); i_unit=_col(head,"UNIT")
    info = _make_info()
    for row in raw[1:]:
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
    files = [p for p in glob.glob(os.path.join(folder, "*"))
             if p.lower().endswith((".xls", ".xlsx"))]
    dbg(box, f"  [匹配] 文件夹内候选合同数 = {len(files)}")
    cores = re.sub(r"[^0-9A-Z]", "", k)
    for p in files:
        n = norm(os.path.basename(p))
        if k in n or n in k:
            dbg(box, f"  [匹配] 精确命中 -> {os.path.basename(p)}"); return p
    for p in files:
        n = norm(os.path.basename(p))
        if cores and cores in n:
            dbg(box, f"  [匹配] 模糊命中 -> {os.path.basename(p)}"); return p
    dbg(box, "  [匹配] 未找到任何匹配合同")
    return None

# ---------- 主处理 ----------
def process(master_path, folder, out_path, box=None):
    wb = load_workbook(master_path, data_only=True)
    ws = wb.active
    dbg(box, f"\n==== 总表 ====")
    dbg(box, f"Sheet: {ws.title}, 行={ws.max_row}, 列={ws.max_column}")
    dbg(box, f"表头行=第{HEADER_ROW}行, 数据从第{DATA_START_ROW}行开始")

    head = [txt(ws.cell(HEADER_ROW, c).value) for c in range(1, ws.max_column + 1)]
    dbg(box, f"表头(第{HEADER_ROW}行): {head}")

    def C(*keys):
        for k in keys:
            for i, h in enumerate(head):
                hu = h.upper().replace(" ", "")
                ku = k.upper().replace(" ", "")
                if ku in hu:
                    dbg(box, f"  列匹配 '{k}' -> 第{i+1}列 ('{h}')")
                    return i + 1
        dbg(box, f"  列匹配 {keys} -> 未找到")
        return None

    c_inv  = C("IV&PL", "INVOICE NO", "发票号", "INVOICE", "合同号")
    c_tax  = C("关税金额", "关税")
    c_tot  = C("总税额", "总税")
    c_note = C("备注")
    c_qty  = C("件数", "数量", "QTY")
    c_unit = C("单位", "UNIT")

    if not c_inv:
        print("⚠️ 仍未找到发票号列！请把表头行(第2行)截图发我。", flush=True)

    n_ok = n_skip = 0
    for r in range(DATA_START_ROW, ws.max_row + 1):
        inv = txt(ws.cell(r, c_inv).value) if c_inv else ""
        if not inv:
            continue
        dbg(box, f"\n--- 第{r}行 发票号='{inv}' ---")
        cp = find_contract(folder, inv, box)
        if not cp:
            if c_note: ws.cell(r, c_note).value = "未找到合同"
            n_skip += 1; continue
        try:
            info = read_contract(cp, box)
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
    print(f"\n完成！处理 {n_ok} 行，跳过 {n_skip} 行", flush=True)
    print(f"输出: {out_path}", flush=True)

# ---------- GUI ----------
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
    root = tk.Tk()
    root.title("报关自动化工具 v1.4")
    root.geometry("720x580")
    mv, fv = tk.StringVar(), tk.StringVar()

    tk.Label(root, text="① 选总表：").grid(row=0, column=0, sticky="e", padx=8, pady=12)
    tk.Entry(root, textvariable=mv, width=58).grid(row=0, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: mv.set(filedialog.askopenfilename(
        filetypes=[("Excel","*.xlsx *.xlsm *.xls")]) or mv.get())).grid(row=0, column=2, padx=6)

    tk.Label(root, text="② 选合同文件夹：").grid(row=1, column=0, sticky="e", padx=8)
    tk.Entry(root, textvariable=fv, width=58).grid(row=1, column=1, padx=4)
    tk.Button(root, text="浏览…", command=lambda: fv.set(filedialog.askdirectory() or fv.get())).grid(row=1, column=2, padx=6)

    box = scrolledtext.ScrolledText(root, height=26, font=("Consolas", 9))
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
