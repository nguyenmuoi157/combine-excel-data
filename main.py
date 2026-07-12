import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# --- LOGIC TAB 1: GỘP EXCEL ---
def process_excel_files(folder_path, progress_var, status_label, root, btn_start, btn_browse):
    try:
        excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f.lower().startswith('iss') and f != 'result.xlsx']
        
        if not excel_files:
            root.after(0, lambda: messagebox.showinfo("Thông báo", "Không tìm thấy file .xlsx nào bắt đầu bằng 'iss' trong thư mục!"))
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            root.after(0, lambda: btn_start.config(state=tk.NORMAL))
            root.after(0, lambda: btn_browse.config(state=tk.NORMAL))
            return

        all_data = []
        total_files = len(excel_files)
        
        for idx, file in enumerate(excel_files):
            root.after(0, lambda f=file: status_label.config(text=f"Đang xử lý: {f}..."))
            file_path = os.path.join(folder_path, file)
            
            try:
                df = pd.read_excel(file_path, usecols=['REF_NO', 'REQUEST_REF_NO'], engine='calamine')
                if 'REF_NO' not in df.columns or 'REQUEST_REF_NO' not in df.columns:
                    continue
                
                ref_data = df[df['REF_NO'].notna()]['REF_NO']
                req_data = df[df['REF_NO'].isna()]['REQUEST_REF_NO']
                combined = pd.concat([ref_data, req_data], ignore_index=True).dropna()
                all_data.append(combined)
            except Exception as e:
                print(f"Lỗi khi đọc file {file}: {e}")
            
            progress = ((idx + 1) / total_files) * 50
            root.after(0, lambda p=progress: progress_var.set(p))
            
        if not all_data:
            root.after(0, lambda: messagebox.showinfo("Thông báo", "Không có dữ liệu hợp lệ nào được tìm thấy."))
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            root.after(0, lambda: progress_var.set(0))
            root.after(0, lambda: btn_start.config(state=tk.NORMAL))
            root.after(0, lambda: btn_browse.config(state=tk.NORMAL))
            return

        root.after(0, lambda: status_label.config(text="Đang gộp dữ liệu..."))
        final_series = pd.concat(all_data, ignore_index=True)
        final_df = pd.DataFrame({'COMBINED_REF': final_series})
        final_df['PREFIX_FT'] = final_df['COMBINED_REF'].astype(str).apply(lambda x: x.split('FT')[0] if 'FT' in x else '')
                
        root.after(0, lambda: status_label.config(text="Đang ghi ra file result.xlsx..."))
        result_path = os.path.join(folder_path, 'result.xlsx')
        max_rows_per_sheet = 1_000_000
        total_rows = len(final_df)
        
        with pd.ExcelWriter(result_path, engine='openpyxl') as writer:
            if total_rows == 0:
                final_df.to_excel(writer, sheet_name='Sheet1', index=False)
            else:
                num_sheets = (total_rows // max_rows_per_sheet) + (1 if total_rows % max_rows_per_sheet != 0 else 0)
                for i in range(num_sheets):
                    start_row = i * max_rows_per_sheet
                    end_row = min((i + 1) * max_rows_per_sheet, total_rows)
                    chunk = final_df.iloc[start_row:end_row]
                    sheet_name = f'Sheet{i+1}'
                    chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                    progress = 50 + ((i + 1) / num_sheets) * 50
                    root.after(0, lambda p=progress: progress_var.set(p))
        
        root.after(0, lambda: progress_var.set(100))
        root.after(0, lambda: status_label.config(text="Hoàn thành!"))
        root.after(0, lambda: messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra file:\n{result_path}\nTổng số dòng: {total_rows}"))
        
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{err}"))
        root.after(0, lambda: status_label.config(text="Sẵn sàng"))
        root.after(0, lambda: progress_var.set(0))
    finally:
        root.after(0, lambda: btn_start.config(state=tk.NORMAL))
        root.after(0, lambda: btn_browse.config(state=tk.NORMAL))


# --- LOGIC TAB 2: NỐI FILE DAT ---
def process_dat_files(files_tuple, progress_var, status_label, root, btn_start, btn_browse):
    try:
        files = list(files_tuple)
        if not files:
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            root.after(0, lambda: btn_start.config(state=tk.NORMAL))
            root.after(0, lambda: btn_browse.config(state=tk.NORMAL))
            return
            
        total_files = len(files)
        output_dir = os.path.dirname(files[0])
        output_file = os.path.join(output_dir, 'combined_result.dat')
        
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for idx, file in enumerate(files):
                root.after(0, lambda f=os.path.basename(file): status_label.config(text=f"Đang xử lý: {f}..."))
                
                with open(file, 'r', encoding='utf-8') as infile:
                    for line_num, line in enumerate(infile):
                        if line.startswith('HR'):
                            if idx == 0:
                                outfile.write(line)
                            else:
                                continue
                        else:
                            outfile.write(line)
                
                progress = ((idx + 1) / total_files) * 100
                root.after(0, lambda p=progress: progress_var.set(p))
                
        root.after(0, lambda: progress_var.set(100))
        root.after(0, lambda: status_label.config(text="Hoàn thành!"))
        root.after(0, lambda: messagebox.showinfo("Thành công", f"Đã nối xong các file .dat!\nFile kết quả: {output_file}"))
        
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{err}"))
        root.after(0, lambda: status_label.config(text="Sẵn sàng"))
        root.after(0, lambda: progress_var.set(0))
    finally:
        root.after(0, lambda: btn_start.config(state=tk.NORMAL))
        root.after(0, lambda: btn_browse.config(state=tk.NORMAL))


# --- LOGIC TAB 3: BÁO CÁO ATM ---
def get_df_with_auto_header(file, sheet_name):
    try:
        df_head = pd.read_excel(file, sheet_name=sheet_name, engine='calamine', header=None, nrows=10)
        header_idx = 0
        for i in range(len(df_head)):
            if 'TRN_DT' in df_head.iloc[i].values or 'LCY_AMOUNT' in df_head.iloc[i].values:
                header_idx = i
                break
        return pd.read_excel(file, sheet_name=sheet_name, engine='calamine', header=header_idx)
    except Exception as e:
        print(f"Lỗi khi đọc file {file} sheet {sheet_name}: {e}")
        return pd.DataFrame()

def process_atm_files(files_tuple, progress_var, status_label, root, btn_start, btn_browse):
    try:
        files = list(files_tuple)
        if not files:
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            root.after(0, lambda: btn_start.config(state=tk.NORMAL))
            root.after(0, lambda: btn_browse.config(state=tk.NORMAL))
            return
            
        total_files = len(files)
        output_dir = os.path.dirname(files[0])
        output_file = os.path.join(output_dir, 'atm_result.xlsx')
        
        sheet12_data = []
        sheet13_data = []

        for idx, file in enumerate(files):
            filename = os.path.basename(file)
            root.after(0, lambda f=filename: status_label.config(text=f"Đang xử lý: {f}..."))
            
            try:
                # Bước 1: Sheet 12
                df12 = get_df_with_auto_header(file, '12')
                if 'LCY_AMOUNT' in df12.columns and 'TRN_DT' in df12.columns:
                    grouped12 = df12.groupby('TRN_DT', as_index=False)['LCY_AMOUNT'].sum()
                    res12 = pd.DataFrame({
                        'Ngày (TRN_DT)': grouped12['TRN_DT'],
                        'Thông tin (File/REF_NO)': filename,
                        'Số tiền (LCY_AMOUNT)': grouped12['LCY_AMOUNT']
                    })
                    sheet12_data.append(res12)

                # Bước 2: Sheet 13
                df13 = get_df_with_auto_header(file, '13')
                if 'LCY_AMOUNT' in df13.columns and 'TRN_DT' in df13.columns:
                    grouped13 = df13.groupby('TRN_DT', as_index=False)['LCY_AMOUNT'].sum()
                    res13 = pd.DataFrame({
                        'Ngày (TRN_DT)': grouped13['TRN_DT'],
                        'Thông tin (File/REF_NO)': filename,
                        'Số tiền (LCY_AMOUNT)': grouped13['LCY_AMOUNT']
                    })
                    sheet13_data.append(res13)

                # Bước 3 & 4: Sheet 'GD lỗi'
                df_err = get_df_with_auto_header(file, 'GD lỗi')
                required_cols = ['AC_NO', 'TRN_DT', 'TRN_REF_NO', 'LCY_AMOUNT']
                if all(c in df_err.columns for c in required_cols):
                    df_err['AC_NO_str'] = df_err['AC_NO'].astype(str).str.lower().str.strip()
                    
                    # Điều kiện vnd -> lưu vào sheet 12
                    vnd_mask = df_err['AC_NO_str'].str.startswith('vnd')
                    if vnd_mask.any():
                        vnd_data = df_err[vnd_mask]
                        res_err12 = pd.DataFrame({
                            'Ngày (TRN_DT)': vnd_data['TRN_DT'],
                            'Thông tin (File/REF_NO)': vnd_data['TRN_REF_NO'],
                            'Số tiền (LCY_AMOUNT)': vnd_data['LCY_AMOUNT']
                        })
                        sheet12_data.append(res_err12)

                    # Điều kiện 00000162 -> lưu vào sheet 13
                    ac162_mask = df_err['AC_NO_str'].str.startswith('00000162')
                    if ac162_mask.any():
                        ac162_data = df_err[ac162_mask]
                        res_err13 = pd.DataFrame({
                            'Ngày (TRN_DT)': ac162_data['TRN_DT'],
                            'Thông tin (File/REF_NO)': ac162_data['TRN_REF_NO'],
                            'Số tiền (LCY_AMOUNT)': ac162_data['LCY_AMOUNT']
                        })
                        sheet13_data.append(res_err13)
                    
            except Exception as e:
                print(f"Lỗi chung file {filename}: {e}")
                
            progress = ((idx + 1) / total_files) * 90
            root.after(0, lambda p=progress: progress_var.set(p))
            
        root.after(0, lambda: status_label.config(text="Đang ghi ra file atm_result.xlsx..."))
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            if sheet12_data:
                final12 = pd.concat(sheet12_data, ignore_index=True)
                final12.to_excel(writer, sheet_name='12', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='12', index=False)
                
            if sheet13_data:
                final13 = pd.concat(sheet13_data, ignore_index=True)
                final13.to_excel(writer, sheet_name='13', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='13', index=False)
                
        root.after(0, lambda: progress_var.set(100))
        root.after(0, lambda: status_label.config(text="Hoàn thành!"))
        root.after(0, lambda: messagebox.showinfo("Thành công", f"Đã xử lý xong báo cáo ATM!\nFile kết quả: {output_file}"))
        
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{err}"))
        root.after(0, lambda: status_label.config(text="Sẵn sàng"))
        root.after(0, lambda: progress_var.set(0))
    finally:
        root.after(0, lambda: btn_start.config(state=tk.NORMAL))
        root.after(0, lambda: btn_browse.config(state=tk.NORMAL))


# --- GIAO DIỆN GUI ---
def setup_gui():
    root = tk.Tk()
    root.title("Công cụ Xử lý Dữ liệu")
    root.geometry("550x330")
    root.resizable(False, False)

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # --- TAB 1: EXCEL ---
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Gộp Excel (REF_NO)")
    
    t1_folder_var = tk.StringVar()
    t1_progress_var = tk.DoubleVar()

    def t1_browse():
        folder = filedialog.askdirectory()
        if folder:
            t1_folder_var.set(folder)

    def t1_start():
        folder = t1_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một thư mục hợp lệ.")
            return
            
        t1_progress_var.set(0)
        t1_btn_start.config(state=tk.DISABLED)
        t1_btn_browse.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=process_excel_files, args=(folder, t1_progress_var, t1_status, root, t1_btn_start, t1_btn_browse))
        thread.daemon = True
        thread.start()

    tk.Label(tab1, text="Chọn thư mục chứa các file Excel (.xlsx):", anchor="w").pack(fill=tk.X, padx=10, pady=(10, 5))
    
    t1_frame_folder = tk.Frame(tab1)
    t1_frame_folder.pack(fill=tk.X, padx=10, pady=(0, 15))
    tk.Entry(t1_frame_folder, textvariable=t1_folder_var, state='readonly', width=45).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
    t1_btn_browse = tk.Button(t1_frame_folder, text="Chọn thư mục", command=t1_browse)
    t1_btn_browse.pack(side=tk.RIGHT)
    
    t1_btn_start = tk.Button(tab1, text="Bắt đầu gộp dữ liệu", command=t1_start, height=2, bg="#4CAF50", fg="black")
    t1_btn_start.pack(fill=tk.X, padx=10, pady=(0, 15))
    
    ttk.Progressbar(tab1, variable=t1_progress_var, maximum=100).pack(fill=tk.X, padx=10, pady=(0, 10))
    t1_status = tk.Label(tab1, text="Sẵn sàng", fg="blue", anchor="w")
    t1_status.pack(fill=tk.X, padx=10)


    # --- TAB 2: NỐI FILE DAT ---
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Nối File .DAT")
    
    t2_files_var = tk.StringVar()
    t2_progress_var = tk.DoubleVar()
    t2_selected_files = []

    def t2_browse():
        files = filedialog.askopenfilenames(filetypes=[("DAT files", "*.dat"), ("All files", "*.*")])
        if files:
            nonlocal t2_selected_files
            t2_selected_files = files
            t2_files_var.set(f"Đã chọn {len(files)} file")

    def t2_start():
        if not t2_selected_files:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn các file cần gộp.")
            return
            
        t2_progress_var.set(0)
        t2_btn_start.config(state=tk.DISABLED)
        t2_btn_browse.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=process_dat_files, args=(t2_selected_files, t2_progress_var, t2_status, root, t2_btn_start, t2_btn_browse))
        thread.daemon = True
        thread.start()

    tk.Label(tab2, text="Chọn các file .dat cần nối:", anchor="w").pack(fill=tk.X, padx=10, pady=(10, 5))
    
    t2_frame_files = tk.Frame(tab2)
    t2_frame_files.pack(fill=tk.X, padx=10, pady=(0, 15))
    tk.Entry(t2_frame_files, textvariable=t2_files_var, state='readonly', width=45).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
    t2_btn_browse = tk.Button(t2_frame_files, text="Chọn các file", command=t2_browse)
    t2_btn_browse.pack(side=tk.RIGHT)
    
    t2_btn_start = tk.Button(tab2, text="Bắt đầu nối file", command=t2_start, height=2, bg="#4CAF50", fg="black")
    t2_btn_start.pack(fill=tk.X, padx=10, pady=(0, 15))
    
    ttk.Progressbar(tab2, variable=t2_progress_var, maximum=100).pack(fill=tk.X, padx=10, pady=(0, 10))
    t2_status = tk.Label(tab2, text="Sẵn sàng", fg="blue", anchor="w")
    t2_status.pack(fill=tk.X, padx=10)


    # --- TAB 3: BÁO CÁO ATM ---
    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Báo cáo ATM")
    
    t3_files_var = tk.StringVar()
    t3_progress_var = tk.DoubleVar()
    t3_selected_files = []

    def t3_browse():
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if files:
            nonlocal t3_selected_files
            t3_selected_files = files
            t3_files_var.set(f"Đã chọn {len(files)} file")

    def t3_start():
        if not t3_selected_files:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn các file cần xử lý.")
            return
            
        t3_progress_var.set(0)
        t3_btn_start.config(state=tk.DISABLED)
        t3_btn_browse.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=process_atm_files, args=(t3_selected_files, t3_progress_var, t3_status, root, t3_btn_start, t3_btn_browse))
        thread.daemon = True
        thread.start()

    tk.Label(tab3, text="Chọn các file Excel (ATM):", anchor="w").pack(fill=tk.X, padx=10, pady=(10, 5))
    
    t3_frame_files = tk.Frame(tab3)
    t3_frame_files.pack(fill=tk.X, padx=10, pady=(0, 15))
    tk.Entry(t3_frame_files, textvariable=t3_files_var, state='readonly', width=45).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
    t3_btn_browse = tk.Button(t3_frame_files, text="Chọn các file", command=t3_browse)
    t3_btn_browse.pack(side=tk.RIGHT)
    
    t3_btn_start = tk.Button(tab3, text="Bắt đầu xử lý ATM", command=t3_start, height=2, bg="#4CAF50", fg="black")
    t3_btn_start.pack(fill=tk.X, padx=10, pady=(0, 15))
    
    ttk.Progressbar(tab3, variable=t3_progress_var, maximum=100).pack(fill=tk.X, padx=10, pady=(0, 10))
    t3_status = tk.Label(tab3, text="Sẵn sàng", fg="blue", anchor="w")
    t3_status.pack(fill=tk.X, padx=10)

    root.mainloop()

if __name__ == "__main__":
    setup_gui()
