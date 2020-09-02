#!"D:\Program Files\PycharmProjects\bruteforce\venv\Scripts\python.exe"
# EASY-INSTALL-ENTRY-SCRIPT: 'msoffcrypto-tool==4.10.0','console_scripts','msoffcrypto-tool'
__requires__ = 'msoffcrypto-tool==4.10.0'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('msoffcrypto-tool==4.10.0', 'console_scripts', 'msoffcrypto-tool')()
    )
