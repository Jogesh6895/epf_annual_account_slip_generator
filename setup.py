from cx_Freeze import setup, Executable

base = None
# For Windows GUI, uncomment the following line:
# if sys.platform == 'win32':
#     base = "Win32GUI"

build_exe_options = {
    "packages": ["csv", "subprocess", "time", "sys", "os"],
    "include_files": ["InputFiles"],
    "excludes": [],
    "optimize": 0,
}

setup(
    name="EPF Report Generator",
    version="2.0.0",
    description="Comprehensive tool for generating EPF annual account slips from Excel data",
    author="EPF Calculator Team",
    options={"build_exe": build_exe_options},
    executables=[
        Executable("epf_calculator.py", base=base, targetName="EPFReportGenerator.exe")
    ],
)
