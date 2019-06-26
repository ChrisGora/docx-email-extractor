import csv
import re
from pathlib import Path

import docx2txt

if __name__ == '__main__':
    excluded = []
    emails = set()
    for filename in Path('docs').glob('**/*.docx'):
        text = docx2txt.process(filename)
        for email in re.findall(r'\w+@\w+\.\w+\.?\w*', text):
            if email not in excluded:
                emails.add(email)

    print(emails)

    with open('emails.csv', 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        for email in emails:
            writer.writerow([email])
