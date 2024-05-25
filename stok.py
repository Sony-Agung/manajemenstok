import tk 


def auth_gspread():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("stok-423416-d7096f8ed2d4.json", scope)
    client = gspread.authorize(creds)
    return client

def update_counts():
    try:
        count_out = float(len(ws_out.get_all_values()) - 1)
        count_in = float(len(ws_in.get_all_values()) - 1)
    except ValueError:
        count_out = 0.0
        count_in = 0.0
    
    label_count_out.config(text=f"Total Material OUT: {count_out}")
    label_count_in.config(text=f"Total Material IN: {count_in}")
    total_count = count_out + count_in
    label_total_count.config(text=f"Total Material IN and OUT: {total_count:.1f}")

def to_float(value):
    try:
        return float(value)
    except ValueError:
        return None if value == "" else 0.0

def simpan_data_out():
    barcode = entry_out_barcode.get()
    rit1 = to_float(entry_out_rit1.get())
    rit2 = to_float(entry_out_rit2.get())
    rit3 = to_float(entry_out_rit3.get())
    plant = to_float(entry_out_plant.get())

    try:
        cell = ws_stokrm.find(barcode)
        nama_barang = ws_stokrm.cell(cell.row, 3).value
        ws_out.append_row([barcode, barcode, nama_barang, rit1, rit2, rit3, plant])

        entry_out_barcode.delete(0, tk.END)
        entry_out_rit1.delete(0, tk.END)
        entry_out_rit2.delete(0, tk.END)
        entry_out_rit3.delete(0, tk.END)
        entry_out_plant.delete(0, tk.END)
        
        update_counts()
    except gspread.exceptions.CellNotFound:
        messagebox.showerror("Error", "Barcode tidak ditemukan di StokRM")

def simpan_data_in():
    barcode = entry_in_barcode.get()
    rit1 = to_float(entry_in_rit1.get())
    rit2 = to_float(entry_in_rit2.get())

    try:
        cell = ws_stokrm.find(barcode)
        nama_barang = ws_stokrm.cell(cell.row, 3).value
        
        # Check if adding a new row will exceed the row limit
        if len(ws_in.get_all_values()) < 1000:
            ws_in.append_row([barcode, barcode, nama_barang, rit1, rit2])
            entry_in_barcode.delete(0, tk.END)
            entry_in_rit1.delete(0, tk.END)
            entry_in_rit2.delete(0, tk.END)
            update_counts()
        else:
            messagebox.showerror("Error", "Jumlah baris maksimum telah tercapai di lembar IN.")
    except gspread.exceptions.APIError as e:
        if 'exceeds grid limits' in str(e):
            messagebox.showerror("Error", "Jumlah baris maksimum telah tercapai di lembar IN.")
        else:
            messagebox.showerror("Error", "Terjadi kesalahan API saat menyimpan data IN.")

def reset_entries():
    # Hapus semua data dari material OUT
    ws_out.delete_rows(2, ws_out.row_count)  # Hapus semua baris kecuali header
    
    # Hapus semua data dari material IN
    ws_in.delete_rows(2, ws_in.row_count)  # Hapus semua baris kecuali header
    
    # Kosongkan semua field input
    entry_out_barcode.delete(0, tk.END)
    entry_out_rit1.delete(0, tk.END)
    entry_out_rit2.delete(0, tk.END)
    entry_out_rit3.delete(0, tk.END)
    entry_out_plant.delete(0, tk.END)

    entry_in_barcode.delete(0, tk.END)
    entry_in_rit1.delete(0, tk.END)
    entry_in_rit2.delete(0, tk.END)

    # Perbarui tampilan jumlah total
    update_counts()

def authenticate():
    password = password_entry.get()
    if password == "12345":
        login_window.destroy()
        main_window()
    else:
        messagebox.showerror("Error", "Password salah!")

def main_window():
    global entry_out_barcode, entry_out_rit1, entry_out_rit2, entry_out_rit3, entry_out_plant
    global entry_in_barcode, entry_in_rit1, entry_in_rit2
    global label_count_out, label_count_in, label_total_count
    global ws_out, ws_in, ws_stokrm
    
    client = auth_gspread()
    spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1wHqdFtx9hc1sT6TNTCCMMavSHbIld3mDIpIuS9CX_X0/edit#gid=1743056787")
    ws_out = spreadsheet.worksheet("OUT")
    ws_in = spreadsheet.worksheet("IN")
    ws_stokrm = spreadsheet.worksheet("Exp")

    root = tk.Tk()
    root.title("Manajemen Data Material")

    tabControl = ttk.Notebook(root)

    tab_out = ttk.Frame(tabControl)
    tabControl.add(tab_out, text='Material OUT')

    tab_in = ttk.Frame(tabControl)
    tabControl.add(tab_in, text='Material IN')

    tab_expired = ttk.Frame(tabControl)
    tabControl.add(tab_expired, text='Expired Items')

    tabControl.pack(expand=1, fill="both")

    # Material OUT
    tk.Label(tab_out, text="Barcode", font=("Helvetica", 14)).grid(column=0, row=0, padx=10, pady=10)
    entry_out_barcode = tk.Entry(tab_out, font=("Helvetica", 14))
    entry_out_barcode.grid(column=1, row=0, padx=10, pady=10)

    tk.Label(tab_out, text="Rit1", font=("Helvetica", 14)).grid(column=0, row=1, padx=10, pady=10)
    entry_out_rit1 = tk.Entry(tab_out, font=("Helvetica", 14))
    entry_out_rit1.grid(column=1, row=1, padx=10, pady=10)

    tk.Label(tab_out, text="Rit2", font=("Helvetica", 14)).grid(column=0, row=2, padx=10, pady=10)
    entry_out_rit2 = tk.Entry(tab_out, font=("Helvetica", 14))
    entry_out_rit2.grid(column=1, row=2, padx=10, pady=10)

    tk.Label(tab_out, text="Rit3", font=("Helvetica", 14)).grid(column=0, row=3, padx=10, pady=10)
    entry_out_rit3 = tk.Entry(tab_out, font=("Helvetica", 14))
    entry_out_rit3.grid(column=1, row=3, padx=10, pady=10)

    tk.Label(tab_out, text="Plant", font=("Helvetica", 14)).grid(column=0, row=4, padx=10, pady=10)
    entry_out_plant = tk.Entry(tab_out, font=("Helvetica", 14))
    entry_out_plant.grid(column=1, row=4, padx=10, pady=10)

    button_out = tk.Button(tab_out, text="Simpan", font=("Helvetica", 14), command=simpan_data_out)
    button_out.grid(column=1, row=5, padx=10, pady=10)

    button_reset_out = tk.Button(tab_out, text="Reset", font=("Helvetica", 14), command=reset_entries)
    button_reset_out.grid(column=1, row=6, padx=10, pady=10)

    label_count_out = tk.Label(tab_out, text="Total Material OUT: 0.0", font=("Helvetica", 14))
    label_count_out.grid(column=1, row=7, padx=10, pady=10)

    # Material IN
    tk.Label(tab_in, text="Barcode", font=("Helvetica", 14)).grid(column=0, row=0, padx=10, pady=10)
    entry_in_barcode = tk.Entry(tab_in, font=("Helvetica", 14))
    entry_in_barcode.grid(column=1, row=0, padx=10, pady=10)

    tk.Label(tab_in, text="Rit1", font=("Helvetica", 14)).grid(column=0, row=1, padx=10, pady=10)
    entry_in_rit1 = tk.Entry(tab_in, font=("Helvetica", 14))
    entry_in_rit1.grid(column=1, row=1, padx=10, pady=10)

    tk.Label(tab_in, text="Rit2", font=("Helvetica", 14)).grid(column=0, row=2, padx=10, pady=10)
    entry_in_rit2 = tk.Entry(tab_in, font=("Helvetica", 14))
    entry_in_rit2.grid(column=1, row=2, padx=10, pady=10)

    button_in = tk.Button(tab_in, text="Simpan", font=("Helvetica", 14), command=simpan_data_in)
    button_in.grid(column=1, row=3, padx=10, pady=10)

    label_count_in = tk.Label(tab_in, text="Total Material IN: 0.0", font=("Helvetica", 14))
    label_count_in.grid(column=1, row=4, padx=10, pady=10)

    # Expired Items
    tree = ttk.Treeview(tab_expired, columns=("kode", "nama_barang", "expired", "jumlah"), show="headings")
    tree.heading("kode", text="Kode")
    tree.heading("nama_barang", text="Nama Barang")
    tree.heading("expired", text="Expired")
    tree.heading("jumlah", text="Jumlah")
    tree.pack(fill="both", expand=True)

    scrollbar = ttk.Scrollbar(tab_expired, orient="vertical", command=tree.yview)
    scrollbar.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scrollbar.set)

    expired_items = ws_stokrm.get_all_values()
    expired_items.pop(0)  # Remove header row

    for item in expired_items:
        # Assuming the relevant columns are Kode, Nama Barang, Expired, and Jumlah
        kode = item[1]
        nama_barang = item[2]
        expired = item[3]
        jumlah = item[4]
        tree.insert("", "end", values=(kode, nama_barang, expired, jumlah))

    label_total_count = tk.Label(root, text="Total Material IN and OUT: 0.0", font=("Helvetica", 16))
    label_total_count.pack(pady=20)

    root.mainloop()

login_window = tk.Tk()
login_window.title("Login")

tk.Label(login_window, text="Password", font=("Helvetica", 14)).pack(pady=10)
password_entry = tk.Entry(login_window, font=("Helvetica", 14), show='*')
password_entry.pack(pady=10)

tk.Button(login_window, text="Login", font=("Helvetica", 14), command=authenticate).pack(pady=10)

login_window.mainloop()
