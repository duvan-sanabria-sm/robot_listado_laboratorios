import requests
from PyPDF2 import PdfReader
from io import BytesIO

def get_page_count_from_url(url):
    response = requests.get(url)
    pdf_stream = BytesIO(response.content)
    reader = PdfReader(pdf_stream)
    return len(reader.pages)