#!/usr/bin/env python3
# HDFC SME Bulk Pay - Prototype App
# Lightweight Tkinter app for beneficiaries & payments, exports matching HDFC templates.
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd, os, datetime, sys, json
from pathlib import Path

APP_DIR = Path(__file__).parent
MASTER_FILE = APP_DIR / "master.xlsx"
PAYMENTS_FILE = APP_DIR / "payments.xlsx"
LOGS_FILE = APP_DIR / "logs.xlsx"
DAILY_FILE = APP_DIR / "daily_summary.xlsx"
TEMPLATE_BENE = APP_DIR / "bulk-bene-file-upload.xlsx"
TEMPLATE_PAYOUT = APP_DIR / "bulk-payout-template.xlsx"
DAILY_LIMIT = 5000000
BATCH_SIZE = 25

def ensure_files():
    if not MASTER_FILE.exists():
        pd.DataFrame(columns=pd.read_excel(TEMPLATE_BENE).columns if TEMPLATE_BENE.exists() else ["Beneficiary Name","Account Number","IFSC","Bank Name","Branch","Beneficiary Type","Address1","Address2","City","State","Pincode","Phone","Email","IndividualLimit"]).to_excel(MASTER_FILE, index=False)
    if not PAYMENTS_FILE.exists():
        pd.DataFrame(columns=["EntryID","Date","Party Name","Account Number","IFSC","Bank Name","Mode","Amount","Narration","Status","Exported_File"]).to_excel(PAYMENTS_FILE, index=False)
    if not LOGS_FILE.exists():
        pd.DataFrame(columns=["Timestamp","Action","Details","FileName","RecordCount","User"]).to_excel(LOGS_FILE, index=False)
    if not DAILY_FILE.exists():
        pd.DataFrame(columns=["Date","TotalPayments","LimitExceeded","Notes"]).to_excel(DAILY_FILE, index=False)

def read_master():
    try:
        return pd.read_excel(MASTER_FILE)
    except:
        return pd.DataFrame()

def save_master(df):
    df.to_excel(MASTER_FILE, index=False)

def read_payments():
    try:
        return pd.read_excel(PAYMENTS_FILE)
    except:
        return pd.DataFrame()

def save_payments(df):
    df.to_excel(PAYMENTS_FILE, index=False)

def append_log(action, details, fname="", count=0):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df = read_logs()
    df = df.append({"Timestamp":ts,"Action":action,"Details":details,"FileName":fname,"RecordCount":count,"User":os.getlogin()}, ignore_index=True)
    df.to_excel(LOGS_FILE, index=False)

def read_logs():
    try:
        return pd.read_excel(LOGS_FILE)
    except:
        return pd.DataFrame(columns=["Timestamp","Action","Details","FileName","RecordCount","User"])

def is_valid_ifsc(ifsc):
    if not isinstance(ifsc, str): return False
    s = ifsc.strip().upper()
    return len(s)==11 and s[:4].isalpha() and s[4:].isalnum()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HDFC SME Bulk Pay")
        self.geometry("1000x650")
        ensure_files()
        self.create_widgets()
        self.refresh_master()
        self.refresh_payments()
        self.show_logs()

    def create_widgets(self):
        tabControl = ttk.Notebook(self)
        tab1 = ttk.Frame(tabControl); tab2 = ttk.Frame(tabControl); tab3 = ttk.Frame(tabControl); tab4 = ttk.Frame(tabControl)
        tabControl.add(tab1, text='Beneficiaries'); tabControl.add(tab2, text='Payments'); tabControl.add(tab3, text='Logs'); tabControl.add(tab4, text='Dashboard')
        tabControl.pack(expand=1, fill="both")

        # Beneficiary Tab
        frm = ttk.Frame(tab1, padding=10); frm.pack(fill="both", expand=True)
        labels = list(pd.read_excel(TEMPLATE_BENE).columns) if TEMPLATE_BENE.exists() else ["Beneficiary Name","Account Number","IFSC","Bank Name","Branch","Beneficiary Type","Address1","Address2","City","State","Pincode","Phone","Email","IndividualLimit"]
        self.bene_vars = {}
        for i,lab in enumerate(labels):
            ttk.Label(frm, text=lab).grid(row=i, column=0, sticky="w", pady=2)
            v = tk.StringVar(); self.bene_vars[lab]=v
            ttk.Entry(frm, textvariable=v, width=50).grid(row=i, column=1, sticky="w")
        ttk.Button(frm, text="Add to Master", command=self.add_beneficiary).grid(row=0, column=2, padx=10)
        ttk.Button(frm, text="Export Beneficiaries (HDFC file)", command=self.export_beneficiaries).grid(row=1, column=2, padx=10)

        # Master list display
        cols = ("acc","ifsc","limit","name")
        self.master_tree = ttk.Treeview(frm, columns=cols, show="headings", height=10)
        for c,h in [("acc","Account"),("ifsc","IFSC"),("limit","Limit"),("name","Name")]:
            self.master_tree.heading(c, text=h); self.master_tree.column(c, width=200)
        self.master_tree.grid(row=len(labels)+1, column=0, columnspan=3, pady=10)

        # Payments Tab
        frm2 = ttk.Frame(tab2, padding=10); frm2.pack(fill="both", expand=True)
        p_labels = ["Party Name","Account Number","IFSC","Bank Name","Mode","Amount","Narration"]
        self.pay_vars = {}
        for i,lab in enumerate(p_labels):
            ttk.Label(frm2, text=lab).grid(row=i, column=0, sticky="w", pady=2)
            v = tk.StringVar(); self.pay_vars[lab]=v
            ttk.Entry(frm2, textvariable=v, width=50).grid(row=i, column=1, sticky="w")
        ttk.Button(frm2, text="Add Payment", command=self.add_payment).grid(row=0, column=2, padx=10)
        ttk.Button(frm2, text="Export Payments", command=self.export_payments).grid(row=1, column=2, padx=10)

        self.pay_tree = ttk.Treeview(frm2, columns=("party","acc","ifsc","mode","amt","status"), show="headings", height=10)
        for c,h in [("party","Party"),("acc","Account"),("ifsc","IFSC"),("mode","Mode"),("amt","Amount"),("status","Status")]:
            self.pay_tree.heading(c, text=h); self.pay_tree.column(c, width=130)
        self.pay_tree.grid(row=10, column=0, columnspan=3, pady=10)

        # Logs Tab
        self.log_box = tk.Text(tab3)
        self.log_box.pack(fill="both", expand=True)

        # Dashboard Tab
        self.dashboard = tk.Text(tab4)
        self.dashboard.pack(fill="both", expand=True)
        ttk.Button(self, text="Refresh Dashboard", command=self.refresh_dashboard).pack(side="bottom", pady=5)

    def add_beneficiary(self):
        row = {k: v.get().strip() for k,v in self.bene_vars.items()}
        if not row[list(row.keys())[0]] or not row[list(row.keys())[1]] or not row[list(row.keys())[2]]:
            messagebox.showerror("Missing", "Mandatory fields required (Name, Account, IFSC)")
            return
        if not is_valid_ifsc(row[list(row.keys())[2]]):
            messagebox.showerror("IFSC", "Invalid IFSC format")
            return
        df = read_master()
        if df.empty:
            df = pd.DataFrame(columns=list(row.keys())+["Beneficiary Added"])
        # check duplicates by account+ifsc
        if not df.empty and any((df.iloc[:,1].astype(str).str.strip()==row[list(row.keys())[1]].strip()) & (df.iloc[:,2].astype(str).str.upper()==row[list(row.keys())[2]].strip().upper())):
            messagebox.showwarning("Duplicate", "Beneficiary with same account+IFSC exists in Master")
            return
        new = {k: row.get(k,"") for k in df.columns if k!="Beneficiary Added"}
        new["Beneficiary Added"] = "Y"
        df = df.append(new, ignore_index=True)
        save_master(df)
        append_log("Add Beneficiary", f"Added {row[list(row.keys())[0]]}", "", 1)
        messagebox.showinfo("Added", "Beneficiary added to Master list")
        self.refresh_master()

    def refresh_master(self):
        for i in self.master_tree.get_children(): self.master_tree.delete(i)
        df = read_master()
        if df is None or df.empty: return
        for idx, r in df.iterrows():
            acc = r.iloc[1] if len(r)>1 else ""; ifsc = r.iloc[2] if len(r)>2 else ""; lim = r.iloc[13] if len(r)>13 else ""; name = r.iloc[0] if len(r)>0 else ""
            self.master_tree.insert("", "end", values=(acc, ifsc, lim, name))

    def add_payment(self):
        row = {k: v.get().strip() for k,v in self.pay_vars.items()}
        if not row["Party Name"] or not row["Account Number"] or not row["IFSC"] or not row["Amount"]:
            messagebox.showerror("Missing", "Party, Account, IFSC and Amount required")
            return
        try:
            amt = float(row["Amount"])
        except:
            messagebox.showerror("Amount", "Invalid amount")
            return
        if not is_valid_ifsc(row["IFSC"]):
            messagebox.showerror("IFSC", "Invalid IFSC")
            return
        df = read_payments()
        if df.empty:
            df = pd.DataFrame(columns=["EntryID","Date","Party Name","Account Number","IFSC","Bank Name","Mode","Amount","Narration","Status","Exported_File"])
        entryid = len(df)+1
        new = {"EntryID":entryid, "Date": datetime.date.today(), "Party Name": row["Party Name"], "Account Number": row["Account Number"],
               "IFSC": row["IFSC"], "Bank Name": row["Bank Name"], "Mode": row["Mode"], "Amount": amt, "Narration": row["Narration"], "Status": "", "Exported_File": ""}
        df = df.append(new, ignore_index=True)
        save_payments(df)
        append_log("Add Payment", f"Added payment to {row['Party Name']}", "", 1)
        messagebox.showinfo("Added", "Payment added")
        self.refresh_payments()

    def refresh_payments(self):
        for i in self.pay_tree.get_children(): self.pay_tree.delete(i)
        df = read_payments()
        if df is None or df.empty: return
        for idx, r in df.iterrows():
            self.pay_tree.insert("", "end", values=(r.get("Party Name",""), r.get("Account Number",""), r.get("IFSC",""), r.get("Mode",""), r.get("Amount",""), r.get("Status","")))

    def export_beneficiaries(self):
        df = read_master()
        if df.empty:
            messagebox.showinfo("No Data", "Master list empty")
            return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="bulk-bene-file-upload.xlsx", filetypes=[("Excel","*.xlsx")])
        if not out: return
        df.to_excel(out, index=False)
        append_log("Export Beneficiaries", f"Wrote {out}", out, len(df))
        messagebox.showinfo("Exported", f"Beneficiaries exported to {out}")

    def export_payments(self):
        df = read_payments()
        if df.empty:
            messagebox.showinfo("No Data", "No payments to export")
            return
        total = df["Amount"].sum()
        if total > DAILY_LIMIT:
            if not messagebox.askyesno("Daily limit", f"Total {total} exceeds ₹50,00,000. Continue?"): return
        base = Path(APP_DIR)
        folder = base / f"Payments_{datetime.date.today().strftime('%d%m%Y')}"
        folder.mkdir(parents=True, exist_ok=True)
        payout_headers = list(pd.read_excel(TEMPLATE_PAYOUT).columns) if TEMPLATE_PAYOUT.exists() else ["SNo","AccountNumber","IFSC","AccountName","Amount","Mode","ValueDate","Narration"]
        records = df.to_dict('records')
        batch = []
        filecount = 1
        written = 0
        for rec in records:
            batch.append(rec)
            if len(batch) == BATCH_SIZE:
                outp = folder / f"Batch_{datetime.date.today().strftime('%d%m%Y')}_{filecount}.xlsx"
                out_df = pd.DataFrame(batch)
                final = pd.DataFrame()
                for h in payout_headers:
                    if "acc" in h.lower() or "account" in h.lower(): final[h] = out_df.get("Account Number","")
                    elif "ifsc" in h.lower(): final[h] = out_df.get("IFSC","")
                    elif "name" in h.lower(): final[h] = out_df.get("Party Name","")
                    elif "amount" in h.lower(): final[h] = out_df.get("Amount","")
                    elif "mode" in h.lower(): final[h] = out_df.get("Mode","")
                    elif "narr" in h.lower(): final[h] = out_df.get("Narration","")
                    else: final[h] = ""
                final.to_excel(outp, index=False)
                append_log("Export Payments", f"Wrote {outp}", str(outp), len(batch))
                written += len(batch)
                filecount += 1
                batch = []
        if batch:
            outp = folder / f"Batch_{datetime.date.today().strftime('%d%m%Y')}_{filecount}.xlsx"
            out_df = pd.DataFrame(batch)
            final = out_df
            final.to_excel(outp, index=False)
            append_log("Export Payments", f"Wrote {outp}", str(outp), len(batch))
            written += len(batch)
        messagebox.showinfo("Exported", f"Wrote {written} payments in {filecount} files to {folder}")
        df["Status"] = "Exported"
        df["Exported_File"] = "Batches in " + str(folder)
        df.to_excel(PAYMENTS_FILE, index=False)
        self.refresh_payments()
        self.show_logs()

    def show_logs(self):
        df = read_logs()
        self.log_box.delete(1.0, "end")
        for idx,row in df.iterrows():
            self.log_box.insert("end", f"{row['Timestamp']} | {row['Action']} | {row['Details']} | {row['FileName']}\n")

    def refresh_dashboard(self):
        dfm = read_master(); dfp = read_payments()
        self.dashboard.delete(1.0, "end")
        self.dashboard.insert("end", f"Total beneficiaries: {len(dfm)}\n")
        self.dashboard.insert("end", f"Payments today: {len(dfp)} | Total amount: {dfp['Amount'].sum() if not dfp.empty else 0}\n")
        last = read_logs()
        if not last.empty:
            self.dashboard.insert("end", f"Last export: {last.iloc[-1]['Timestamp']} - {last.iloc[-1]['Action']}\n")

if __name__ == '__main__':
    App().mainloop()
