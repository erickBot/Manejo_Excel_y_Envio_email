import openpyxl
from datetime import datetime, date, time, timedelta
import calendar

list_DaysLel = []
list_DaysH2s = []
count_LEL = []
count_H2S = []

def run():
    name_file = "Lista_de_Sensores_HGAS.xlsx"
    file = openpyxl.load_workbook(name_file)
    sheet1 = file.get_sheet_by_name('LEL')
    sheet2 = file.get_sheet_by_name('H2S')
    cell_lel = sheet1['H4':'H19']
    cell_h2s = sheet2['H4':'H22']
    dateNow = datetime.now()
    lookCell(cell_lel, dateNow, 730)
    lookCell(cell_h2s, dateNow, 365)
    #salta a las funciones para llenar las celdas con datos nuevos
    llenar_cell(sheet1, list_DaysLel, count_LEL, 'I4', 'I19')
    llenar_cell(sheet2, list_DaysH2s, count_H2S, 'I4', 'I22')
    #guarda archivo
    file.save(name_file)

    print(count_LEL)
    print(count_H2S)

def lookCell(celdas, dateNow, totalDays):
    expirados = 0
    porVencer = 0
    vigentes = 0

    for row in celdas:
        for cell in row:
            dateLast= cell.value
            days = dateNow - dateLast
            diferencia = totalDays - days.days
            if totalDays == 730:
                list_DaysLel.append(diferencia)
            if totalDays == 365:
                list_DaysH2s.append(diferencia)
            if diferencia < 0:
                expirados += 1
            if diferencia > 0 and diferencia < 30:
                porVencer += 1
            if diferencia > 30:
                vigentes += 1

    if totalDays == 730:
        count_LEL.append(porVencer)
        count_LEL.append(expirados)
        count_LEL.append(vigentes)


    else:
        count_H2S.append(porVencer)
        count_H2S.append(expirados)
        count_H2S.append(vigentes)

def llenar_cell(sheet, list_Days, countList, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            cell.value = list_Days[i]
            i += 1

    sheet['I24'] = countList[0] #porVencer
    sheet['I25'] = countList[1] #expirados
    sheet['I26'] = countList[2] #vigentes

if __name__ == "__main__":
    run()
