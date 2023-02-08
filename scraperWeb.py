import requests, PyPDF2
from io import BytesIO

url = 'https://journals.sagepub.com/doi/pdf/10.1177/0194599820931041'
#url = 'https://www.medrxiv.org/content/medrxiv/early/2020/05/25/2020.05.22.20108845.full.pdf'
#url = 'https://www.medrxiv.org/content/medrxiv/early/2020/05/27/2020.05.26.20113464.full.pdf'
response = requests.get(url)
my_raw_data = response.content

with BytesIO(my_raw_data) as data:


    read_pdf = PyPDF2.PdfReader(data)
    for page in range(len(read_pdf.pages)):
        print(read_pdf.pages[page].extract_text().encode('utf8'))
        #string = read_pdf.pages[page].extract_text()