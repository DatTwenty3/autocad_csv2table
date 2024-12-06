import tkinter as tk
from tkinter import filedialog, messagebox
from pyautocad import Autocad, APoint
import csv


# Hàm tạo bảng từ file CSV
def create_table_from_csv(file_path, insertion_point=(0, 0), row_height=3, column_width=10, text_height=1.5,
                          title="Bảng", column_names=None):
    """
    Đọc file CSV và tạo bảng trong AutoCAD với tiêu đề bảng và tên các cột được nhập từ giao diện người dùng.

    :param file_path: Đường dẫn đến file CSV
    :param insertion_point: Điểm chèn bảng (tọa độ X, Y)
    :param row_height: Chiều cao của mỗi hàng trong bảng
    :param column_width: Chiều rộng của mỗi cột trong bảng
    :param text_height: Kích thước chữ bên trong bảng
    :param title: Tiêu đề bảng
    :param column_names: Danh sách tên các cột
    """
    try:
        # Kết nối AutoCAD
        acad = Autocad(create_if_not_exists = True)
        print("Kết nối AutoCAD thành công.")

        # Đọc file CSV
        with open(file_path, 'r', encoding = 'utf-8') as file:
            reader = csv.reader(file)
            data = list(reader)

        # Xác định số hàng và số cột
        num_rows = len(data)
        num_cols = len(data[0]) if data else 0

        # Thêm 1 cột cho "TT" và tạo một ô tiêu đề bảng
        num_cols += 1  # Thêm 1 cột cho "TT"

        # Tạo bảng trong AutoCAD
        if num_rows > 0 and num_cols > 0:
            table = acad.model.AddTable(APoint(*insertion_point), num_rows + 2, num_cols, row_height, column_width)

            # Thiết lập kích thước chữ
            table.SetTextHeight(0, text_height)

            # Thiết lập ô tiêu đề gộp
            table.MergeCells(0, 0, 0, num_cols - 1)  # Gộp tất cả các ô trong hàng đầu tiên

            # Điền tiêu đề vào ô gộp
            table.SetText(0, 0, title)

            # Điền tên cột vào hàng thứ 1
            table.SetText(1, 0, "TT")  # Cột thứ tự
            for col_idx in range(1, num_cols):
                if col_idx - 1 < len(column_names):
                    table.SetText(1, col_idx, column_names[col_idx - 1])  # Điền tên cột
                else:
                    table.SetText(1, col_idx, "")  # Các cột còn lại để trống

            # Điền dữ liệu vào bảng và thêm số thứ tự vào cột "TT"
            for row_idx, row in enumerate(data):
                table.SetText(row_idx + 2, 0, str(row_idx + 1))  # Cột "TT" (Số thứ tự)
                for col_idx, cell in enumerate(row):
                    table.SetText(row_idx + 2, col_idx + 1, cell)  # Điền dữ liệu vào các cột còn lại

            print(f"Đã tạo bảng từ file CSV tại tọa độ {insertion_point}.")
        else:
            print("File CSV không có dữ liệu.")

    except Exception as e:
        print(f"Lỗi: {e}")


# Hàm mở hộp thoại chọn file CSV
def select_csv_file():
    file_path = filedialog.askopenfilename(title = "Chọn file CSV", filetypes = [("CSV Files", "*.csv")])
    return file_path


# Hàm cập nhật giao diện người dùng dựa trên số cột trong file CSV
def update_column_entries(file_path, columns_frame):
    """
    Cập nhật giao diện để hiển thị các trường nhập tên cột dựa trên số cột trong file CSV.

    :param file_path: Đường dẫn đến file CSV
    :param columns_frame: Frame chứa các trường nhập liệu tên cột
    """
    # Đọc file CSV để xác định số cột
    with open(file_path, 'r', encoding = 'utf-8') as file:
        reader = csv.reader(file)
        data = list(reader)

    # Xác định số cột
    num_cols = len(data[0]) if data else 0

    # Xóa các trường nhập liệu cũ
    for widget in columns_frame.winfo_children():
        widget.destroy()

    # Tạo các trường nhập liệu tên cột
    for col_idx in range(num_cols):
        label = tk.Label(columns_frame, text = f"Tên cột {col_idx + 1}:")
        label.grid(row = col_idx, column = 0, padx = 10, pady = 5)
        entry = tk.Entry(columns_frame, width = 30)
        entry.grid(row = col_idx, column = 1, padx = 10, pady = 5)


# Hàm xử lý click chuột để chọn vị trí trong AutoCAD
def get_insertion_point(acad):
    """
    Lấy vị trí người dùng click chuột trong AutoCAD để chèn bảng.
    """
    print("Chọn điểm chèn bảng trong AutoCAD.")
    messagebox.showinfo("Cảnh báo","Chọn điểm chèn bảng trong AutoCAD. Nhấn 'OK' để tiếp tục")
    point = acad.doc.Utility.GetPoint()  # Lấy điểm click từ AutoCAD
    return point


# Hàm tạo giao diện người dùng
def create_gui():
    # Tạo cửa sổ
    root = tk.Tk()
    root.title("LEDAT - Tạo Bảng trong AutoCAD")

    # Tạo frame cho phần chọn file CSV và tên bảng
    csv_frame = tk.Frame(root)
    csv_frame.grid(row = 0, column = 0, columnspan = 2, padx = 10, pady = 5)

    # Nút chọn file CSV
    select_file_button = tk.Button(csv_frame, text = "Chọn file CSV", command = lambda: on_select_file(csv_frame))
    select_file_button.grid(row = 0, column = 0, padx = 10, pady = 5)

    # Hiển thị đường dẫn file
    file_path_label = tk.Label(csv_frame, text = "Chưa chọn file", width = 50, anchor = "w")
    file_path_label.grid(row = 0, column = 1, padx = 10, pady = 5)

    # Tiêu đề bảng
    title_label = tk.Label(csv_frame, text = "Tiêu đề bảng:")
    title_label.grid(row = 1, column = 0, padx = 10, pady = 5)
    title_entry = tk.Entry(csv_frame, width = 50)
    title_entry.grid(row = 1, column = 1, padx = 10, pady = 5)

    # Tạo frame cho các trường nhập tên cột
    columns_frame = tk.Frame(root)
    columns_frame.grid(row = 1, column = 0, columnspan = 2, padx = 10, pady = 5)

    # Hàm xử lý chọn file CSV
    def on_select_file(csv_frame):
        file_path = select_csv_file()
        if file_path:
            file_path_label.config(text = file_path)
            update_column_entries(file_path, columns_frame)

    # Nút chọn vị trí
    def on_select_location():
        acad = Autocad(create_if_not_exists = True)
        insertion_point = get_insertion_point(acad)
        print(f"Vị trí chèn bảng đã chọn: {insertion_point}")
        return insertion_point

    # Nút tạo bảng
    def on_create_table():
        file_path = file_path_label.cget("text")
        title = title_entry.get()

        # Lấy tên cột từ các trường nhập liệu
        column_names = []
        for widget in columns_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                column_names.append(widget.get())

        if not file_path:
            print("Vui lòng chọn file CSV.")
            return

        # Lấy vị trí chèn bảng từ AutoCAD
        insertion_point = on_select_location()

        # Tạo bảng từ file CSV và các giá trị nhập
        create_table_from_csv(file_path, insertion_point = insertion_point, row_height = 2.5, column_width = 15,
                              text_height = 1, title = title, column_names = column_names)

    ledat_label = tk.Label(root, text = "...:::PHẦN MỀM ĐƯỢC PHÁT TRIỂN BỞI LÊ ĐẠT - ÂU LẠC CONS:::...", fg = "green",
                        font = ("Arial", 10))
    ledat_label.grid(row = 2, column = 0, columnspan = 3)

    create_table_button = tk.Button(root, text = "Tạo Bảng", command = on_create_table, bg = "blue", fg = "white")
    create_table_button.grid(row = 3, column = 0, columnspan = 2, pady = 10)

    # Hiển thị cửa sổ
    root.mainloop()


# Gọi hàm tạo GUI
create_gui()