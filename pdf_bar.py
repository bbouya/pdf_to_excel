import re
import pdfplumber
import pandas as pd
from collections import namedtuple
import PyPDF2
from tqdm import tqdm

#Line = namedtuple('Line', 'First name, last name, MID, CIN, date of service and reject code')

data = []
columns = ['First name', 'Last name', 'MID', 'ACNT', 'ICN', 'Date of service', 'code']

moa_pattern = re.compile(r'MOA ([A-Z\d\s]+)')
code_pattern = re.compile(r'\b([A-Z]{2}-\d{2})\b')
pattern = re.compile(r'NAME (.*?), (.*?) MID (.*?) ACNT (.*?) ICN (.*?) .*?(\d{4} \d{6})')

extracted_text = []

with open('test2.pdf', 'rb') as pdf_file:
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    total_pages = len(pdf_reader.pages)
    for page in tqdm(pdf_reader.pages, total=total_pages, desc='Extracting text'):
        text = page.extract_text()
        extracted_text.append(text)
    extracted_text = ''.join(extracted_text)

moa_match = [info[:20].strip() for info in moa_pattern.findall(extracted_text)]
code_matches = code_pattern.findall(extracted_text)

matches = pattern.findall(extracted_text)
total_matches = len(matches)

with tqdm(total=total_matches, desc='Processing matches') as pbar:
    for imatch, match in enumerate(matches):
        first_name, last_name, mid, acnt, icn, date_of_service = match[:6]
        dates = date_of_service.split(' ')[1]
        code = f'{moa_match[imatch]} {code_matches[imatch]}'.replace('  ', ',')
        print("First name:", first_name)
        print("Last name:", last_name)
        print("MID:", mid)
        print("ACNT:", acnt)
        print("ICN:", icn)
        print("Date of service:", date_of_service)
        print('code : ', code)
        print("______________________________")
        data.append([first_name, last_name, mid, acnt, icn, dates, code])
        pbar.update(1)

df = pd.DataFrame(data, columns=columns)
df.to_excel('test2.xlsx', index=False)

print('-----------')
