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
    if phone.startswith('84'):
        phone = '0' + phone[2:]
    if len(phone) == 9:
        return True
    if len(phone) == 10 and phone.startswith('0'):
        if phone.startswith(('00', '01', '02', '030', '031')):
            return False
        return True
    if len(phone) > 10 and phone.startswith('84'):
        phone = '0' + phone[2:]
        if len(phone) == 10 and not phone.startswith(('00', '01', '02', '030', '031')):
            return True
    return False

def standardize_phone_number(value):
    if pd.isna(value):
        return None
    phone = re.sub(r'\D', '', str(value))
    
    if phone.startswith('84'):
        phone = '0' + phone[2:]
    
    if len(phone) == 9:
        phone = '0' + phone
    elif len(phone) > 10:
        if phone.startswith('84'):
            phone = '0' + phone[2:]
        else:
            return None
    
    if len(phone) == 10 and phone.startswith('0'):
        if phone.startswith(('00', '01', '02', '030', '031')):
            return None
        return phone
    
    return None

def process_excel(file_path, name_column, phone_column, sheet_name):
    try:
        results = []
        # Đọc Excel với tất cả các cột dưới dạng chuỗi và thêm converters để xử lý cột số điện thoại
        df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name,
            dtype=str,  # Đọc tất cả các cột dưới dạng chuỗi
            converters={phone_column: lambda x: str(x).strip()}  # Đảm bảo cột điện thoại được xử lý đúng
        )
        
        # Xử lý từng dòng trong DataFrame
        for index, row in df.iterrows():
            # Chuyển đổi giá trị số điện thoại thành chuỗi và loại bỏ các khoảng trắng
            phone_value = str(row[phone_column]).strip() if not pd.isna(row[phone_column]) else None
            
            # Thêm log để kiểm tra giá trị trước khi chuẩn hóa
            print(f"Raw phone value: {phone_value}")
            
            phone = standardize_phone_number(phone_value)
            
            # Thêm log để kiểm tra giá trị sau khi chuẩn hóa
            print(f"Standardized phone: {phone}")
            
            if phone and not phone.startswith(('00', '01', '02', '030', '031')):
                # Đảm bảo tên khách hàng cũng được xử lý đúng
                customer_name = str(row[name_column]).strip() if not pd.isna(row[name_column]) else ""
                results.append({
                    'customer_name': customer_name,
                    'phone': phone
                })
        
        if results:
            output_file = os.path.join(os.path.dirname(file_path), 'results.xlsx')
            
            # Kiểm tra xem file results.xlsx đã tồn tại chưa
            if os.path.exists(output_file):
                # Đọc file hiện có với dtype=str để tránh mất số 0
                existing_df = pd.read_excel(output_file, dtype=str)
                last_id = int(existing_df['No id'].max()) if not existing_df.empty else 0
                
                # Tạo DataFrame mới với ID tiếp theo
                new_results = []
                for idx, result in enumerate(results, start=1):
                    new_results.append({
                        'No id': str(last_id + idx),  # Chuyển ID thành chuỗi
                        'Khách hàng': result['customer_name'],
                        'sdt': result['phone']
                    })
                
                new_df = pd.DataFrame(new_results)
                # Gộp DataFrame cũ và mới
                output_df = pd.concat([existing_df, new_df], ignore_index=True)
            else:
                # Tạo file mới nếu chưa tồn tại
                new_results = []
                for idx, result in enumerate(results, start=1):
                    new_results.append({
                        'No id': str(idx),  # Chuyển ID thành chuỗi
                        'Khách hàng': result['customer_name'],
                        'sdt': result['phone']
                    })
                output_df = pd.DataFrame(new_results)
            
            # Ghi ra file Excel với định dạng chuỗi cho tất cả các cột
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                
                # Đặt định dạng text cho cột số điện thoại
                for idx, cell in enumerate(worksheet['C'], 1):  # Cột C là cột 'sdt'
                    cell.number_format = '@'
            
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