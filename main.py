import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def process_excel_files(folder_path, progress_var, status_label, root):
    try:
        # Lấy danh sách các file excel bắt đầu bằng chữ 'iss' (không phân biệt hoa/thường)
        excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f.lower().startswith('iss') and f != 'result.xlsx']
        
        if not excel_files:
            root.after(0, lambda: messagebox.showinfo("Thông báo", "Không tìm thấy file .xlsx nào trong thư mục!"))
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            return

        all_data = []
        total_files = len(excel_files)
        
        for idx, file in enumerate(excel_files):
            root.after(0, lambda f=file: status_label.config(text=f"Đang xử lý: {f}..."))
            file_path = os.path.join(folder_path, file)
            
            try:
                df = pd.read_excel(file_path)
                
                # Kiểm tra cột bắt buộc
                if 'REF_NO' not in df.columns or 'REQUEST_REF_NO' not in df.columns:
                    print(f"Bỏ qua file {file} vì không có đủ 2 cột REF_NO và REQUEST_REF_NO.")
                    continue
                
                # Lấy dữ liệu ở cột REF_NO (những dòng có dữ liệu)
                ref_data = df[df['REF_NO'].notna()]['REF_NO']
                
                # Lấy dữ liệu ở cột REQUEST_REF_NO (khi REF_NO không có dữ liệu)
                req_data = df[df['REF_NO'].isna()]['REQUEST_REF_NO']
                
                # Nối 2 chuỗi dữ liệu này lại với nhau (req_data nằm ngay dưới ref_data)
                combined = pd.concat([ref_data, req_data], ignore_index=True)
                
                # Xoá những dòng trống (nếu cả 2 đều trống)
                combined = combined.dropna()
                
                all_data.append(combined)
            except Exception as e:
                print(f"Lỗi khi đọc file {file}: {e}")
            
            # Cập nhật tiến trình
            progress = ((idx + 1) / total_files) * 50 # 50% quá trình là đọc file
            root.after(0, lambda p=progress: progress_var.set(p))
            
        if not all_data:
            root.after(0, lambda: messagebox.showinfo("Thông báo", "Không có dữ liệu hợp lệ nào được tìm thấy."))
            root.after(0, lambda: status_label.config(text="Sẵn sàng"))
            root.after(0, lambda: progress_var.set(0))
            return

        root.after(0, lambda: status_label.config(text="Đang gộp dữ liệu..."))
        
        # Gộp tất cả dữ liệu từ các file lại thành 1 Series lớn
        final_series = pd.concat(all_data, ignore_index=True)
        final_df = pd.DataFrame({'COMBINED_REF': final_series})
        
        # Thêm cột bên cạnh lấy phần trước chữ 'FT'
        final_df['PREFIX_FT'] = final_df['COMBINED_REF'].astype(str).apply(lambda x: x.split('FT')[0] if 'FT' in x else '')
                
        root.after(0, lambda: status_label.config(text="Đang ghi ra file result.xlsx..."))
        
        # Đường dẫn file kết quả
        result_path = os.path.join(folder_path, 'result.xlsx')
        
        # Giới hạn dòng mỗi sheet
        max_rows_per_sheet = 1_000_000
        total_rows = len(final_df)
        
        # Ghi ra file Excel, hỗ trợ chia sheet
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
                    
                    # Cập nhật tiến trình phần ghi file (50% còn lại)
                    progress = 50 + ((i + 1) / num_sheets) * 50
                    root.after(0, lambda p=progress: progress_var.set(p))
        
        root.after(0, lambda: progress_var.set(100))
        root.after(0, lambda: status_label.config(text="Hoàn thành!"))
        root.after(0, lambda: messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra file:\n{result_path}\nTổng số dòng: {total_rows}"))
        
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{err}"))
        root.after(0, lambda: status_label.config(text="Sẵn sàng"))
        root.after(0, lambda: progress_var.set(0))

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path_var.set(folder_selected)

def start_processing():
    folder = folder_path_var.get()
    if not folder or not os.path.isdir(folder):
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn một thư mục hợp lệ.")
        return
        
    progress_var.set(0)
    btn_start.config(state=tk.DISABLED)
    btn_browse.config(state=tk.DISABLED)
    
    # Chạy xử lý trong một thread riêng để không block giao diện
    thread = threading.Thread(target=run_process_thread, args=(folder,))
    thread.daemon = True
    thread.start()

def run_process_thread(folder):
    process_excel_files(folder, progress_var, status_label, root)
    root.after(0, lambda: btn_start.config(state=tk.NORMAL))
    root.after(0, lambda: btn_browse.config(state=tk.NORMAL))

# --- Cài đặt Giao diện GUI ---
root = tk.Tk()
root.title("Công cụ Gộp dữ liệu Excel (REF_NO)")
root.geometry("500x250")
root.resizable(False, False)

# Biến lưu trữ
folder_path_var = tk.StringVar()
progress_var = tk.DoubleVar()

# Giao diện chính
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(fill=tk.BOTH, expand=True)

lbl_instruct = tk.Label(frame, text="Chọn thư mục chứa các file Excel (.xlsx):", anchor="w")
lbl_instruct.pack(fill=tk.X, pady=(0, 5))

frame_folder = tk.Frame(frame)
frame_folder.pack(fill=tk.X, pady=(0, 15))

entry_folder = tk.Entry(frame_folder, textvariable=folder_path_var, state='readonly', width=45)
entry_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

btn_browse = tk.Button(frame_folder, text="Chọn thư mục", command=browse_folder)
btn_browse.pack(side=tk.RIGHT)

btn_start = tk.Button(frame, text="Bắt đầu gộp dữ liệu", command=start_processing, height=2, bg="#4CAF50", fg="black")
btn_start.pack(fill=tk.X, pady=(0, 15))

progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100)
progress_bar.pack(fill=tk.X, pady=(0, 10))

status_label = tk.Label(frame, text="Sẵn sàng", fg="blue", anchor="w")
status_label.pack(fill=tk.X)

root.mainloop()
