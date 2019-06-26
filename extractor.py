import csv
import re
from pathlib import Path

import docx2txt

if __name__ == '__main__':
    excluded = []
    emails = set()
    for file in Path('docs').glob('**/*.docx'):
        if re.fullmatch(r'8\d\d\d.docx', file.name) is not None:
            print("found file", file)
            text = docx2txt.process(file)
            for email in re.findall(r'\w+@\w+\.\w+\.*\w*', text):
                if email not in excluded:
                    emails.add(email)

    print(emails)

    with open('emails.csv', 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        for email in emails:
            writer.writerow([email])
