import sys
from cx_Freeze import setup, Executable
from tkinter import messagebox

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "excludes": ["tkinter", "unittest"],
    "zip_include_packages": ["encodings", "PySide6"],
}

additional_files = [ ("D:\\aneliqAPP\main.py"), ("D:\\aneliqAPP\data"), ("D:\\aneliqAPP\gеnerated-files"), ("D:\\aneliqAPP\log.py"), ("D:\\aneliqAPP\word.py")]

include_packages = ['tkinter', 'tkcalendar', 'openpyxl', 'pandas']


# base="Win32GUI" should be used only for Windows GUI app
base = "Win32GUI" if sys.platform == "win32" else None
try:
    setup(
        name="my_app",
        version="1.3",
        description="My GUI application!",
        options={"build_exe": {'include_files': additional_files,
                               'packages': include_packages,}},
        executables=[Executable("D:\\aneliqAPP\main.py", base=base,
                                target_name='Програма за генериране на ДС',
                                icon="D:\\aneliqAPP\\app_ico.ico"
                                )],
    )
except Exception:
    messagebox.showerror("Error Setup", Exception)
    raise Exception