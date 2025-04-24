import csv
import re
import os
import shutil
import asyncio
import subprocess
from pathlib import Path
from playwright.async_api import async_playwright

# Install Playwright browsers if not already installed
print('Installing Playwrite if needed:')
subprocess.run(["playwright", "install", "chromium"], check=True)

#########################
# Remove special characters replace with file safe annotation, ex. -dash-
# "Win, Lose, or Navigate" becomes: Win-comma-_Lose-comma-_or_Navigate.html
#########################

def clean_file_title(title):
    escapes = {"-":"-dash-",
                ".":"-dot-",
                "/":"-slash-",
                ":":"-colon-",
                "#":"-hash-",
                ",":"-comma-",
                "'":"-appos-",
                '"':"-quote-",
                '?':"-question-",
                '=':"-equal-",
                '+':"-plus-",
                ' ':"_",
		'(':"-paren-",
		')':"-paren-",
		'&':"-amper-"
              }
    try:
        for char, uni in escapes.items():
            title = title.replace(char, uni).lower() 
        return title
    except UnicodeError as f:
        print('{f} there was a error continuing')

topicAndpresenters = set()

#########################
# Loop through csv file adding to topicAndpresenters set, skip blanks
#########################

with open('names_and_presentations.csv', newline='') as csvfile:
    FeedbackDictReader = csv.DictReader(csvfile, delimiter=',')
    for line in FeedbackDictReader:
        names = line['Presenters']
        topic = line['Presentation Title']
        dates = line['Date']
        values = names, topic
        if len(names) != 0:
            topicAndpresenters.add(values)

#########################
# List directory, check for "certs" folder remove / create if needed
#########################
dirList = os.listdir()
if "certs" in dirList:
    shutil.rmtree("certs")
    os.mkdir("certs")
elif "certs" not in dirList:
    os.mkdir("certs")

#########################
# Read set, read certificate (template), open template, replace name placeholder, replace topic placeholder. create new html file in certs folder. Create PDF from HTML
#########################
async def html_to_pdf(input_html_path, output_pdf_path):
    file_url = f"file:///certs/{input_html_path}"
    
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(file_url, wait_until="networkidle")
        await page.pdf(
            path=str(output_pdf_path),
            format="A4",
            print_background=True
        )
        await browser.close()

for names, topic in topicAndpresenters:
    certificate = "dev/Certificate_01.html"
    with open(certificate, 'r', newline='', encoding='utf-8') as f:
        info = f.read()
    namesReplace = re.sub(r'{{participant}}', names, info)
    dateReplace = re.sub(r'{{date}}', dates, namesReplace)
    finalCertificate = re.sub(r'{{topic}}', topic, dateReplace)
    with open('certs/' + clean_file_title(topic) + '.html', 'a', newline='', encoding='utf-8') as f:
        f.write(finalCertificate)
        print(f'created: {clean_file_title(topic)}.html')
        f.close()
        #Comment out below line if PDF not needed
        asyncio.run(html_to_pdf(clean_file_title(topic)+'.html',clean_file_title(topic)+'.pdf'))

