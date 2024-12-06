from cx_Freeze import setup, Executable

# Đảm bảo rằng bạn đã có file .ico (ví dụ "app_icon.ico") trong cùng thư mục với script
setup(
    name="LEDAT",
    version="1.0",
    description="Ứng dụng được phát triển bởi LEDAT",
    executables=[Executable("main.py", base="Win32GUI", icon="Beer-icon_30353.ico")]
)