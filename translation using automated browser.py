
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import time
import re
import os
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium import webdriver


from docx import Document
"""This code copies data from am english pdf file and translates it into japanese langauge pdf file
 using automated browser and deepl.com website to translate stuff from english to japanese."""
def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec1='utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec1, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()


    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)


    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text
#pdf text data
filename=input('Enter filename:')
#filename='Historica' + '.pdf'
text= convert_pdf_to_txt(filename)
data=[]
n=2500
pattern= "[^\x20-\x7E]+"
for i in range(0, len(text), n):
    data.append("".join(text[i:i+n]))
#connect webdriver
#driver = webdriver.Chrome(executable_path=r"E:\Data Science\eng_to_jap\chromedriver.exe")

gecko = os.path.normpath(os.path.join(os.path.dirname(__file__), 'geckodriver'))
binary = FirefoxBinary(r'C:\Program Files\Mozilla Firefox\firefox.exe')
driver = webdriver.Firefox(firefox_binary=binary, executable_path=gecko+'.exe')
#driver = webdriver.Firefox('geckodriver')
#url='https://deepl.com'
url='https://www.deepl.com/translator#en/ja'
driver.get(url)

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

# //tag[@class=''] class is attribute and inside '' is value i.e. //tag[@Attribute='Value']
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//button[@class='lmt__language_select__active'])[2]"))).click()
ahmed_btn=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//button[@dl-lang='JA'])[2]")))
ahmed_btn.click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//textarea[@class='lmt__textarea lmt__source_textarea lmt__textarea_base_style'])"))).click()

trns_text=driver.find_elements_by_tag_name('textarea')

#convert pdf data
document = Document()
for d in data:
    d=re.sub(pattern, '', d)
    trns_text[0].send_keys(d)
    time.sleep(30)
    bttn=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//textarea)[2]")))
    translated_data=bttn.get_attribute('value')

    document.add_paragraph(translated_data)
    trns_text[0].clear()

document.save('Japanese.docx')
driver.close()





