from cx_Freeze import setup, Executable

target = Executable(
    script="geocoder.py",
    base="Win32GUI",
    compress=False,
    copyDependentFiles=True,
    appendScriptToExe=True,
    appendScriptToLibrary=False,
    icon="icon.ico"
    )

setup(
    name="Geocoder",
    version="2.1",
    description="Automation for Geocoding services from excel file",
    author="Aditya Sidharta",
    executables=[target]
    )
