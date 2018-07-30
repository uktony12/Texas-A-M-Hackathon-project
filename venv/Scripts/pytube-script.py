#!C:\Users\moonbook\PycharmProjects\Hackathon\venv\Scripts\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'pytube==8.0.2','console_scripts','pytube'
__requires__ = 'pytube==8.0.2'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('pytube==8.0.2', 'console_scripts', 'pytube')()
    )
