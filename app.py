import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import sys
import os
from tkinter import ttk

def is_valid_phone(value):
    if pd.isna(value):
        return False
    phone = re.sub(r'\D', '', str(value))
    if len(phone) == 9:
        return True
    if len(phone) == 10 and phone.startswith('0'):
        if phone.startswith(('00', '01', '02', '030', '031')):
            return False
        return True
    return False

def standardize_phone_number(value):
    if pd.isna(value):
        return None
    phone = re.sub(r'\D', '', str(value))
    
    if len(phone) == 9:
        standardized = '0' + phone
        if standardized.startswith(('00', '01', '02', '030', '031')):
            return None
        return standardized
    elif len(phone) == 10 and phone.startswith('0'):
        if phone.startswith(('00', '01', '02', '030', '031')):
            return None
        return phone
    return None

def process_excel(file_path, name_column, phone_column, sheet_name):
    try:
        results = []
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Xử lý từng dòng trong DataFrame
        for index, row in df.iterrows():
            phone = standardize_phone_number(row[phone_column])
            if phone and not phone.startswith(('00', '01', '02', '030', '031')):
                results.append({
                    'id': len(results) + 1,
                    'customer_name': row[name_column],
                    'phone': phone
                })
        
        if results:
            # Tạo DataFrame mới và xuất ra Excel
            output_df = pd.DataFrame(results)
            output_df.columns = ['No id', 'Khách hàng', 'sdt']
            
            output_file = os.path.join(
                os.path.dirname(file_path),
                f'phone_numbers_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            )
            
            output_df.to_excel(output_file, index=False)
            
            messagebox.showinfo("Thành công", 
                              f"Đã tìm thấy {len(results)} số điện thoại.\nĐã lưu vào file {output_file}")
        else:
            messagebox.showinfo("Thông báo", "Không tìm thấy số điện thoại hợp lệ nào.")
    
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

class ExcelProcessor:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.columns = []
        self.sheets = []
        self.current_df = None
        self.setup_gui()
    
    def setup_gui(self):
        self.root.title("Trích xuất số điện thoại")
        
        window_width = 500
        window_height = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # File selection
        label = tk.Label(self.root, text="Chọn file Excel để trích xuất số điện thoại", font=("Arial", 10))
        label.pack(pady=20)

        select_button = tk.Button(
            self.root,
            text="Chọn file Excel",
            command=self.select_file,
            width=20,
            height=2
        )
        select_button.pack(pady=10)
        
        # Sheet selection
        self.sheet_frame = tk.Frame(self.root)
        self.sheet_frame.pack(pady=10)
        
        tk.Label(self.sheet_frame, text="Chọn Sheet:").grid(row=0, column=0, padx=5)
        self.sheet_cb = ttk.Combobox(self.sheet_frame, state='disabled')
        self.sheet_cb.grid(row=0, column=1, padx=5)
        self.sheet_cb.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        
        # Column selection (move after sheet selection)
        self.column_frame = tk.Frame(self.root)
        self.column_frame.pack(pady=10)
        
        # Comboboxes for column selection
        tk.Label(self.column_frame, text="Cột tên:").grid(row=0, column=0, padx=5)
        self.name_column_cb = ttk.Combobox(self.column_frame, state='disabled')
        self.name_column_cb.grid(row=0, column=1, padx=5)
        
        tk.Label(self.column_frame, text="Cột SĐT:").grid(row=1, column=0, padx=5)
        self.phone_column_cb = ttk.Combobox(self.column_frame, state='disabled')
        self.phone_column_cb.grid(row=1, column=1, padx=5)
        
        # Process button
        self.process_button = tk.Button(
            self.root,
            text="Xử lý",
            command=self.process_file,
            width=20,
            height=2,
            state='disabled'
        )
        self.process_button.pack(pady=10)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            try:
                # Đọc danh sách sheets
                xls = pd.ExcelFile(self.file_path)
                self.sheets = xls.sheet_names
                
                # Reset các combobox
                self.sheet_cb['values'] = self.sheets
                self.sheet_cb['state'] = 'readonly'
                
                # Reset column comboboxes
                self.name_column_cb['state'] = 'disabled'
                self.name_column_cb.set('')
                self.phone_column_cb['state'] = 'disabled'
                self.phone_column_cb.set('')
                self.process_button['state'] = 'disabled'
                
                # Nếu chỉ có 1 sheet, tự động chọn
                if len(self.sheets) == 1:
                    self.sheet_cb.set(self.sheets[0])
                    self.on_sheet_selected(None)
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")

    def on_sheet_selected(self, event):
        if self.file_path and self.sheet_cb.get():
            try:
                # Đọc sheet được chọn
                self.current_df = pd.read_excel(self.file_path, sheet_name=self.sheet_cb.get())
                self.columns = self.current_df.columns.tolist()
                
                # Cập nhật combobox columns
                self.name_column_cb['values'] = self.columns
                self.phone_column_cb['values'] = self.columns
                
                self.name_column_cb['state'] = 'readonly'
                self.phone_column_cb['state'] = 'readonly'
                self.process_button['state'] = 'normal'
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc sheet: {str(e)}")

    def process_file(self):
        if not self.sheet_cb.get() or not self.name_column_cb.get() or not self.phone_column_cb.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sheet, cột tên và cột số điện thoại")
            return
        
        process_excel(self.file_path, self.name_column_cb.get(), self.phone_column_cb.get(), self.sheet_cb.get())

def create_gui():
    root = tk.Tk()
    app = ExcelProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    create_gui()