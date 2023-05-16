import re
import pdfplumber
import pandas as pd
from collections import namedtuple

#Line = namedtuple('Line' , 'First name, last name, MID, CIN, date of service and reject code')


import PyPDF2



data = []
columns = ['First name', 'Last name', 'MID', 'ACNT', 'ICN', 'Date of service', 'code']


extracted_text = ''

# Open the PDF file in read-binary mode
with open('test2.pdf', 'rb') as pdf_file:
    # Create a PdfFileReader object to read the PDF
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    
    # Loop through each page in the PDF
    for page_num in range(len(pdf_reader.pages)):
        # Get the page object for the current page
        page = pdf_reader.pages[page_num]
        
        # Extract the text from the page
        text = page.extract_text()

        #read the last line : 
        lines = text.strip().split('\n')
        last_line = lines[-1].strip()
        if last_line in 50*'_' :
            print('ayoub')

        # put all the text pdf in one        
        extracted_text += text



    # the patern to extract the MOA and code
    #moa_pattern = r"MOA\s+(\S+)\s*(\S*)\s*"
    #moa_pattern = r"MOA\s+([A-Za-z0-9]+)\s*([A-Za-z0-9]*)\s*"
    #moa_pattern = r"MOA\s+([A-Za-z0-9]+)\s*([A-Za-z0-9]*)\s*"
    #moa_pattern = r"MOA\s+([A-Za-z][A-Za-z0-9]*)\s*([A-Za-z][A-Za-z0-9]*)\s*"
    moa_pattern = r'MOA ([A-Z\d\s]+)'



    #moa_pattern = r'MOA [A-Za-z0-9\s]+\n'
    #moa_pattern = r'MOA [A-Z0-9]+ [A-Z0-9]+'
    code_pattern = r"\b([A-Z]{2}-\d{2})\b"
    #code_pattern = r'\b(\d{1,2}-\d{1,2})\b'
    # Extract MOA information
    moa_match = re.findall(moa_pattern, extracted_text)
    # Stripping the substrings and removing the newline character (\n)
    moa_matchss = [info[:20].strip() for info in moa_match]
    
    # Extract code information
    code_matches = re.findall(code_pattern, extracted_text)

    # print to see the output
    # Print the extracted information
    print("MOA:", moa_match)
    print("Code Matches:", code_matches)
    print('-------------------')

    # the pattern for the rest
    pattern = r'NAME (.*?), (.*?) MID (.*?) ACNT (.*?) ICN (.*?) .*?(\d{4} \d{6})'
    matches = re.findall(pattern, extracted_text, re.DOTALL)
    for match in range(len(matches)):
        first_name = matches[match][0]
        last_name = matches[match][1]
        mid = matches[match][2]
        acnt = matches[match][3]
        icn = matches[match][4]
        date_of_service = matches[match][5]
        dates = date_of_service.split(' ')
        dates = dates[1]
        if match < len(moa_matchss)  and match < len(code_matches):
            code = moa_matchss[match].replace('  ', ',') + ',' + code_matches[match]
        
        else:
            code = moa_matchss[match].replace('  ', ',')
        #code = moa_matchss[match].replace('  ', ',') + ',' + code_matches[match]
   
        print("First name:", first_name)
        print("Last name:", last_name)
        print("MID:", mid)
        print("ACNT:", acnt)
        print("ICN:", icn)
        print("Date of service:", date_of_service)
        print('code : ', code)
        print("______________________________")
        data.append([first_name, last_name, mid, acnt, icn, dates, code])



df = pd.DataFrame(data, columns=columns)

df.to_excel('test2.xlsx', index=False)

print('-----------')