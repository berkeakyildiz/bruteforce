#!"D:\Program Files\PycharmProjects\bruteforce\venv\Scripts\python.exe"
# EASY-INSTALL-ENTRY-SCRIPT: 'oletools==0.54.2','console_scripts','olebrowse'
__requires__ = 'oletools==0.54.2'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('oletools==0.54.2', 'console_scripts', 'olebrowse')()
    )
