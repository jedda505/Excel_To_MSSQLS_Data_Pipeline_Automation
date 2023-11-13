'''
This script will automatically install packages via pip to a venv
'''

import subprocess

# Activate venv and freeze to requirements.txt 


p1 = subprocess.run('".venv/Scripts/activate.bat" & pip freeze > requirements.txt', shell=True, capture_output = True, text = True)

print(p1.returncode, p1.stdout, p1.stderr)
