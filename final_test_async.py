import asyncio
import re
import pdfplumber
import pandas as pd
from collections import namedtuple
import PyPDF2
from tqdm import tqdm
import re
import pdfplumber
import pandas as pd
from collections import namedtuple


import re
import pdfplumber
import pandas as pd
from collections import namedtuple
import re
import pdfplumber
import pandas as pd
from collections import namedtuple
from tqdm import tqdm

Line = namedtuple('Line', ['First_name', 'Last_name', 'MID', 'ACNT', 'ICN', 'date_of_service_and_reject_code'])

data = []
columns = ['First name', 'Last name', 'MID', 'ACNT', 'ICN', 'Date of service', 'code']

extracted_text = ''

# Open the PDF file in read-binary mode
with open('test2.pdf', 'rb') as pdf_file:
    # Create a PdfFileReader object to read the PDF
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    
    # Loop through each page in the PDF
    for page_num in tqdm(range(len(pdf_reader.pages)), desc='Processing pages'):
        # Get the page object for the current page
        page = pdf_reader.pages[page_num]
        # Extract the text from the page
        text = page.extract_text()
        # Concatenate the text from each page
        extracted_text += text

moa_pattern = r'MOA ([A-Z\d\s]+)'
code_pattern = r"\b([A-Z]{2}-\d{2})\b"

moa_match = re.findall(moa_pattern, extracted_text)
moa_matchss = [info[:20].strip() for info in moa_match]
code_matches = re.findall(code_pattern, extracted_text)

pattern = r'NAME (.*?), (.*?) MID (.*?) ACNT (.*?) ICN (.*?) .*?(\d{4} \d{6})'
matches = re.findall(pattern, extracted_text, re.DOTALL)

for match in tqdm(range(len(matches)), desc='Processing lines'):
    first_name = matches[match][0]
    last_name = matches[match][1]
    mid = matches[match][2]
    acnt = matches[match][3]
    icn = matches[match][4]
    date_of_service_and_reject_code = matches[match][5]
    dates = date_of_service_and_reject_code.split(' ')[1]

    if match < len(moa_matchss) and match < len(code_matches):
        code = moa_matchss[match].replace('  ', ',') + ',' + code_matches[match]
    else:
        code = moa_matchss[match].replace('  ', ',')

    data.append([first_name, last_name, mid, acnt, icn, dates, code])

df = pd.DataFrame(data, columns=columns)

df.to_excel('test2.xlsx', index=False)
