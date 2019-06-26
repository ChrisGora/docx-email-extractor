import re

import docx2txt
from pathlib import Path

if __name__ == '__main__':
    excluded = []
    emails = []
    for filename in Path('docs').glob('**/*.docx'):
        text = docx2txt.process(filename)
        for email in re.findall(r'\w+@\w+\.\w+\.*\w*', text):
            if email not in excluded:
                emails.append(email)

    print(emails)
