import urllib.request
import requests
import openpyxl
import openpyxl.styles
from openpyxl.utils import get_column_letter
from copy import copy


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
    

    merged_cells_list = [str(item).split(":") for item in ws_source.merged_cells]

    #create the workbook
    wb_personal = create_table()

    #select the only sheet in the workbook
    # ws_personal = wb_personal.active
  
    # ws_personal = create_table()

    #get the dimensions of the 'table' from the source file for copying
    # mr = ws_source.rows
    # mc = ws_source.max_column
    
    #go through every cell in the table and extract the info
    # source_rows = list(ws_source.rows)
    # source_cols = list(ws_source.columns)
    # print(list(ws_source.rows)[0])

    #expand the sublist of merged cells to include all the cells that are included in the merge
    for sublist in merged_cells_list:
        # print(f"before: {sublist}")
        i = 0
        while sublist[i][0] != sublist[-1][0]:
            sublist.insert(i+1,f"{chr(ord(sublist[i][0])+1)}{sublist[i][1:]}")
            i = i+1
            if sublist[i][0] == sublist[-1][0]:
                sublist.remove(sublist[i])

        if sublist[i-1][1:] != sublist[-1][1:]:
            sublist.insert(i, f"{sublist[-1][0]}{int(sublist[-1][1:])-1}")

        # print(f"after: {sublist}")
        i=0
        while sublist[i][1:] != sublist[-1][1:]:
            sublist.insert(i+1, f"{sublist[i][0]}{int(sublist[i][1:])+1}")
            i = i+2
            
            if sublist[i][1:] == sublist[-1][1:]:
                sublist.remove(sublist[i])

            if i == len(sublist):
                i = i-1

    for sublist in merged_cells_list:
        check = True
        for i in range(0, len(sublist)-1):
            print(sublist[i][0])
            print(sublist[i+1][0])
            if sublist[i][0] != sublist[i+1][0]:
                check = False

        if check:
            merged_cells_list.remove(sublist)
        # print(f"after: {sublist}")


    # print(merged_cells_list)

    weekdays = {
    "luni" : [],
    "marți" : [],
    "miercuri" : [],
    "joi" : [],
    "vineri": []
    }

    print(f"Choose an group/class")
    group_name = input()

    group_col = ""

    for ro in range(1, ws_source.max_row):
        for col in range(1, ws_source.max_column):
            source_cell = ws_source.cell(row=ro, column=col)
            source_cell_value = str(source_cell.value).lower()
            if source_cell_value in weekdays:
                # if source_cell_value is "luni":
                    # ws_personal
                # print(cell.value)
                weekdays[source_cell_value] += [f"{ro}"]
                weekdays[source_cell_value] += [f"{col}"]
                # print(f"{ro} -> {col}")

            if source_cell_value == group_name.lower():
                group_col = source_cell.column
                # group_col = get_column_letter(group_col)

    # print(weekdays)

    for ro in range(1, ws_source.max_row):
        source_cell = ws_source.cell(row=ro, column=group_col)
        source_cell_value = str(source_cell.value)
        # cell_coordonates = f"{get_column_letter(ro)}{ro}:{get_column_letter(group_col)}{group_col}"
        cell_coordonates = f"{get_column_letter(group_col)}{ro}"

        found = (cell_coordonates in sublist for sublist in merged_cells_list)
        i=0
        for sublist in merged_cells_list:
            print(f"{i}:{sublist}")
            i = i+1
            if cell_coordonates in sublist:
                # print(f"{cell_coordonates} -> {sublist}")
                pass
                # TODO
                #dau unmerge la cell-ul merged si extrag textul din primul elem al listei
                # print(f"source: {source_cell_value}")

                #unmerge
                # ws_source.unmerge_cells(f"{sublist[0]}:{sublist[-1]}")
                # print(f"merged source: {ws_source.cell(row=ro, column=ord(sublist[0][0])).value}")

        # if found:
            # print(source_cell)
            # print(cell_coordonates)
            # print(ro)

    
        # print(source_cell_value)

    # print(merged_cells_list)

    # for cell in source_cols[0]:
        # if cell.value:
            # print(cell.value)

    # for cell in source_cols[1]:
        # if cell.value:
            # print(cell.value)

    # for row in range(1, ws_source.max_row):
        # for cell in source_rows[0]:
            # if cell.value !=  None:
                # print(cell.value)
    # for row in range(mr):
        # for c in range(row):
            # print(ws_source.c)


    # for (row,col), source_cell in ws_source._cells.items():   
    #     personal_cell = ws_personal.cell(column=col, row=row)
    #     personal_cell.value = source_cell.value
    #     personal_cell.data_type = source_cell.data_type

    #     if source_cell.has_style:
    #         personal_cell.font = copy(source_cell.font)
    #         personal_cell.border = copy(source_cell.border)
    #         personal_cell.fill = copy(source_cell.fill)
    #         personal_cell.number_format = copy(source_cell.number_format)
    #         personal_cell.protection = copy(source_cell.protection)
    #         personal_cell.alignment = copy(source_cell.alignment)




    #save to file
    wb_personal.save("table.xlsx")

def create_table():
      #create the workbook
    wb_personal = openpyxl.Workbook()

    #select the only sheet in the workbook
    ws_personal = wb_personal.active
    
    ws_personal.cell(row=1, column=2).value = "LUNI"
    ws_personal.cell(row=1, column=3).value = "MARTI"
    ws_personal.cell(row=1, column=4).value = "MIERCURI"
    ws_personal.cell(row=1, column=5).value = "JOI"
    ws_personal.cell(row=1, column=6).value = "VINERI"

    #weekdays
    for i in range(1,7):
        ws_personal.cell(row=1, column=i).font = openpyxl.styles.Font(bold=True)

    #hours every day
    j = 8
    for i in range(2, 14):
        ws_personal.cell(row=i, column=1).value = j
        j = j + 1


    for i in range(1,14):
        for j in range(1, 7):
            ws_personal.cell(row=i, column=j).border = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style="thin", color = "000000"),
                                                                            right=openpyxl.styles.Side(border_style="thin", color= "000000"),
                                                                            bottom=openpyxl.styles.Side(border_style="thin", color="000000"),
                                                                            top=openpyxl.styles.Side(border_style="thin", color="000000"))

    #!!!fa niste calcule sa vezi cate litere are fiecare coloana si in functie de dimensiune seteaza un width corespunzator
    for i in range(1, 7):
        size = len(str(ws_personal.cell(row=1, column=i).value))
        letter = get_column_letter(i)
        ws_personal.column_dimensions[letter].width = (size + 2) * 1.2

    return wb_personal

def main():
    get_file()
    extract_table()

if __name__ == '__main__':
    main()
