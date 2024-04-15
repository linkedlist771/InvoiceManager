from cx_Freeze import setup, Executable
import sys
import os
sys.setrecursionlimit(5000)  # Increase the recursion limit of the Python interpreter

# GUI applications require a different base on Windows (the default is for a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

import os

build_exe_options = {
    "packages": ["os", "sys", "PyQt5", "openpyxl", "pandas", "fitz"],  # Include external packages used
    "include_files": [
        (os.path.join("export", "custom_form.py"), "export/custom_form.py"),
        (os.path.join("export", "gui.py"), "export/gui.py"),
        (os.path.join("export", "invoice.py"), "export/invoice.py"),
        (os.path.join("export", "pdf_utils.py"), "export/pdf_utils.py"),
        (os.path.join("export", "xlsx_utils.py"), "export/xlsx_utils.py"),

    ],
# "include_files": [
#     (os.path.join("output_directory", "custom_form.py"), "export/custom_form.py"),
#     (os.path.join("output_directory", "gui.py"), "export/gui.py"),
#     (os.path.join("output_directory", "invoice.py"), "export/invoice.py"),
#     (os.path.join("output_directory", "pdf_utils.py"), "export/pdf_utils.py"),
#     (os.path.join("output_directory", "xlsx_utils.py"), "export/xlsx_utils.py"),
# ],

    # "excludes": ["tkinter"],  # Exclude modules that are not used to reduce size
    "excludes": ["tkinter", "http", "email"],
    # "zip_include_packages": ["*"],  # 压缩所有包
    # "zip_exclude_packages": [],  # 不排除任何包从压缩中
    # "optimize": 2  # Python 代码优化级别

}


# Executable
executables = [
    Executable(
        script=os.path.join("export", "gui.py"),
        base=base,
    )
]

setup(
    name="YourApp",
    version="1.0",
    description="Your Application Description",
    options={"build_exe": build_exe_options},
    executables=executables
)

# 使用pyarmour加密build/export 下面的所有文件
