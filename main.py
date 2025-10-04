import urllib.request
import requests
import openpyxl


def get_file():
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
    with open("orar_full.xlsx", "wb") as f:
        f.write(response.content)


def extract_table():
    
    wb_general = openpyxl.load_workbook(filename="orar_full.xlsx")


    sheet_name = ""

    print(f"Choose an year and specialization from the list: {wb_general.sheetnames}")
    sheet_name = input()

    while(sheet_name not in wb_general.sheetnames):
        print("Doesnt fit any of the ones in the list! Please write one from the list!")
        print(wb_general.sheetnames)
        sheet_name = input()


    ws_source = wb_general.worksheets[wb_general.sheetnames.index(sheet_name)]


    #create the workbook
    wb_personal = openpyxl.Workbook()

    #select the only sheet in the workbook
    ws_personal = wb_personal.active
  
    #get the dimensions of the 'table' from the source file for copying
    mr = ws_source.max_row
    mc = ws_source.max_column

    #go through every cell in the table and copy it
    for i in range(1, mr+1):
        for j in range(1, mc+1):
            c = ws_source.cell(i, j)
            ws_personal.cell(i,j).value = c.value 


    #save to file
    wb_personal.save("table.xlsx")


def main():
    get_file()
    extract_table()

if __name__ == '__main__':
    main()
