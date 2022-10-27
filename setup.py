import cx_Freeze
import sys
import matplotlib

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main.py", base=base, icon="beehive_icon.png")]
packages = ['tkinter', 'sys', 'os', 'time', 'pandas', 'simplekml',
            'pyproj', 'shapely', 'polycircles', 'pykml']
cx_Freeze.setup(
    name = "Exchange Bees Toolkit",
    options = {"build_exe": {"packages":packages, "include_files":["beehive_icon.png"]}},
    version = "0.01",
    description = "Exchange Bees Internal Tools",
    executables = executables
    )



