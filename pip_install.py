'''
This script will automatically install packages via pip to a venv
'''

import subprocess

# use this line to run venv and install packages via pip 

p1 = subprocess.run('".venv/Scripts/activate.bat" & pip install boto3', shell=True, capture_output = True, text = True)

print(p1.returncode, p1.stdout, p1.stderr)
