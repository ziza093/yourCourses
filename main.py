import urllib.request
import requests

#this is the website from which I download the file (spreadsheet)
site_url = "https://ac.tuiasi.ro/studenti/didactic/orar/"

#the document link, but calculated later
doc_url = ""

#try to reach the website
site_response = urllib.request.urlopen(site_url)
#extract its content
site_content = site_response.read()

#this is the phrase the help me extract the spreadsheet
phrase = "orar ac"

#iterate through every line of code from the website untill I reach the line which has my spreadsheet link and extract the link
for line in site_content.decode("utf-8").splitlines():
    if phrase in line.lower():
        doc_url = line.split("\"")

        for elem in doc_url:
            if "https" in elem:
                doc_url = elem
                break
        break

#extract from the url the spreadsheet id and complete it with the file format extension to export it
doc_url = doc_url.split("/edit")[0]+"/export?format=xlsx"

#make a get request to the google sheet
response = requests.get(doc_url)

#write its contents to an local file
with open("orar.xlsx", "wb") as f:
    f.write(response.content)