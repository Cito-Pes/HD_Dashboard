import sys, os
from cx_Freeze import setup, Executable
from PySide6.QtCore import QLibraryInfo
import shiboken6
import PySide6

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

f = open(resource_path("./config.txt"), 'r', encoding='utf-8')
lines = f.readlines()
for line in lines:
    line = line.strip()  # 줄 끝의 줄 바꿈 문자를 제거한다.

    if line[:11] == "app_title =":
        App_Title = line[11:]

    if line[:9] == "app_ver =":
        App_Ver = line[9:]

# os.environ['NLS_LANG'] = '.UTF8'  # 또는
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 현재 가상환경(.venv)의 루트
venv_root = sys.prefix
shiboken_pyd = shiboken6.__file__
pyside6_pyd = PySide6.__file__

buildOptions = dict(zip_exclude_packages = ["PySide6", "shiboken6"],
                    zip_include_packages = [],
                    include_msvcr = True,
                    packages = ["PySide6.QtWidgets","pandas","datetime","pyodbc","os","sys","PySide6.QtGui","math","PySide6.QtCore", "shiboken6","matplotlib"],
                    excludes = [],
                    include_files = [('./static'),"./config.txt",
                                     # Shiboken6 DLL
                                     (shiboken6.__file__, os.path.basename(shiboken6.__file__)),
                                     # PySide6 핵심 DLL
                                     # (os.path.join(venv_root, "Lib", "site-packages", "PySide6", "pyside6.dll"),"pyside6.dll"),
                                        (pyside6_pyd, os.path.basename(pyside6_pyd)),
                                     (os.path.join(venv_root, "Lib", "site-packages", "PySide6", "Qt6Core.dll"),
                                      "Qt6Core.dll"),
                                     (os.path.join(venv_root, "Lib", "site-packages", "PySide6", "Qt6Gui.dll"),
                                      "Qt6Gui.dll"),
                                     (os.path.join(venv_root, "Lib", "site-packages", "PySide6", "Qt6Widgets.dll"),
                                      "Qt6Widgets.dll"),
                                     # Qt 플러그인(플랫폼, 이미지 포맷 등)
                                     (QLibraryInfo.path(QLibraryInfo.PluginsPath), "plugins"),
                                     # (옵션) shiboken 지원 스크립트
                                     #(os.path.join(venv_root, "Lib", "site-packages", "shibokensupport"),"shibokensupport"),
                                    "./static/D2Coding.ttc"  # 한글 폰트 경로
                                     ])


base = "Win32GUI" if sys.platform == "win32" else None
exe = [Executable("4X4_DashBoard.py", base=base, target_name=f"{App_Title} ({App_Ver}).exe", icon="./static/ht.ico")]

setup(
    name= App_Title,
    version = App_Ver,
    author = "ChoiJO",
    description = "TM센터 실적확인",
    options = dict(build_exe = buildOptions),
    executables = exe
)

'''
version

0.1 : 20250721
'''
#빌드 사용 방법
#python 4X4_Board_Install.py build