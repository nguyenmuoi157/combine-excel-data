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
        # Lưu file kết quả ở cùng thư mục với file đầu tiên được chọn
        output_dir = os.path.dirname(files[0])
        output_file = os.path.join(output_dir, 'combined_result.dat')
        
        with open(output_file, 'w', encoding='utf-8') as outfile:
            for idx, file in enumerate(files):
                root.after(0, lambda f=os.path.basename(file): status_label.config(text=f"Đang xử lý: {f}..."))
                
                with open(file, 'r', encoding='utf-8') as infile:
                    for line_num, line in enumerate(infile):
                        # Giữ lại Header (dòng bắt đầu bằng HR) nếu là file đầu tiên, bỏ qua nếu là các file sau
                        if line.startswith('HR'):
                            if idx == 0:
                                outfile.write(line)
                            else:
                                continue # Bỏ qua header của file thứ 2 trở đi
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

# --- GIAO DIỆN GUI ---
def setup_gui():
    root = tk.Tk()
    root.title("Công cụ Xử lý Dữ liệu")
    root.geometry("550x300")
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

    root.mainloop()

if __name__ == "__main__":
    setup_gui()
