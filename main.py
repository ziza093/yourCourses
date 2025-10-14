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


def merged_cells_sublists(merged_cells_list):

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

    return merged_cells_list

def get_group_col(ws_source):

    print(f"Choose an group/class")
    group_name = input()

    for ro in range(1, ws_source.max_row):
        for col in range(1, ws_source.max_column):
            source_cell = ws_source.cell(row=ro, column=col)
            source_cell_value = str(source_cell.value).lower()

            if source_cell_value == group_name.lower():
                group_col = source_cell.column

    return group_col


def get_courses(ws_source, group_col):

    courses_list = {}

    for ro in range(1, ws_source.max_row):
        for col in range(1, group_col):
            source_cell = ws_source.cell(row=ro, column=col)
            source_cell_value = str(source_cell.value).lower()
            if "c s" in source_cell_value or "p p" in source_cell_value or "p i" in source_cell_value:
                courses_list.update({source_cell_value:ro})                

    return courses_list

def get_cells(ws_source, group_col):
    cells_list = {}

    for ro in range(1, ws_source.max_row):
        source_cell = ws_source.cell(row=ro, column=group_col)
        source_cell_value = str(source_cell.value).lower()

        if source_cell_value:
            cells_list.update({source_cell_value:ro})
    
    return cells_list

def get_weekdays(ws_source):
    weekdays = {
    "luni" : [],
    "marți" : [],
    "miercuri" : [],
    "joi" : [],
    "vineri": []
    }

    for ro in range(1, ws_source.max_row):
        for col in range(1, ws_source.max_column):
            source_cell = ws_source.cell(row=ro, column=col)
            source_cell_value = str(source_cell.value).lower()

            if source_cell_value in weekdays:
                weekdays[source_cell_value] = ro

    return weekdays

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
    

 

    #get the column pozition in which your group is situated
    group_col = get_group_col(ws_source)

    #extract only the courses the group attends
    courses_list = get_courses(ws_source, group_col)
    
    #extract all the cells 
    cells_list = get_cells(ws_source, group_col)

    #extract the weekdays positions
    weekdays = get_weekdays(ws_source)

    #create the workbook
    wb_personal = create_table(ws_source, courses_list, cells_list, group_col, weekdays)

    #save to file
    wb_personal.save("table.xlsx")


def add_personal_all_data(ws_personal, weekday, courses_list, cells_list, weekdays, week_col):
    #create monday column(using courses and projects)
    for course in courses_list:
        if int(courses_list[course]) >= int(weekdays[weekday]) and int(courses_list[course] <= int(weekdays[weekday]) + 11):
            if courses_list[course] == weekdays[weekday]:
                ws_personal.cell(row=2, column=week_col).value = course    
            else:
                hour = 2 + int(courses_list[course]) - int(weekdays[weekday])
                ws_personal.cell(row=hour, column=week_col).value = course

    #creating monday column(using the labs and the rest)
    for cell in cells_list:
        if int(cells_list[cell]) >= int(weekdays[weekday]) and int(cells_list[cell] <= int(weekdays[weekday]) + 11):
            if cells_list[cell] == weekdays[weekday]:
                ws_personal.cell(row=2, column=week_col).value = cell    
            else:
                hour = 2 + int(cells_list[cell]) - int(weekdays[weekday])
                ws_personal.cell(row=hour, column=week_col).value = cell


def merge_final_cells(ws_personal):
    for col in range(2, ws_personal.max_column + 1):
        for ro in range(2,ws_personal.max_row + 1):
            personal_cell = ws_personal.cell(row=ro, column=col)
            personal_cell_value = personal_cell.value

            if personal_cell_value:
                if "p i" in personal_cell_value or "p p" in personal_cell_value:
                    ro = ro + 2
                else:
                    ws_personal.merge_cells(start_row=ro, start_column=col, end_row=ro+1, end_column=col)
                    ro = ro + 1

def create_table(ws_source, courses_list, cells_list, group_col, weekdays):
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

    
    for i in range(1, 7):
        size = len(str(ws_personal.cell(row=1, column=i).value))
        letter = get_column_letter(i)
        ws_personal.column_dimensions[letter].width = (size + 2) * 1.2

    
    add_personal_all_data(ws_personal, "luni", courses_list, cells_list, weekdays, 2)
    add_personal_all_data(ws_personal, "marți", courses_list, cells_list, weekdays, 3)
    add_personal_all_data(ws_personal, "miercuri", courses_list, cells_list, weekdays, 4)
    add_personal_all_data(ws_personal, "joi", courses_list, cells_list, weekdays, 5)
    add_personal_all_data(ws_personal, "vineri", courses_list, cells_list, weekdays, 6)

    merge_final_cells(ws_personal)
    merge_final_cells(ws_personal)
    merge_final_cells(ws_personal)
    merge_final_cells(ws_personal)
    merge_final_cells(ws_personal)


    return wb_personal

def main():
    get_file()
    extract_table()

if __name__ == '__main__':
    main()
