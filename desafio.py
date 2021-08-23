import gspread
import pandas as pd

gc = gspread.service_account(filename='creds.json') # calling the credentials to access my google drive
sh = gc.open_by_key('1nv2ptydfBGZG5wlFysgpd1sgr7RJNh2PeI0hcyJSSrI') # key of my document in google drive
df = sh.sheet1
sheet = df.get_all_values()

students_sheet = pd.DataFrame(sheet)

# students_sheet = pd.read_excel(sheet)

naf_col = students_sheet[7]

for mean in students_sheet.index:

    if (students_sheet.P1[mean] + students_sheet.P2[mean] + students_sheet.P3[mean])/3 >= 70:
        students_sheet.Situação[mean] = 'Aprovado'

    elif 50 <= (students_sheet.P1[mean] + students_sheet.P2[mean] + students_sheet.P3[mean])/3 < 70:
        students_sheet.Situação[mean] = 'Exame Final'

    elif (students_sheet.P1[mean] + students_sheet.P2[mean] + students_sheet.P3[mean])/3 < 50:
        students_sheet.Situação[mean] = 'Reprovado por nota'


for faults in students_sheet.index:

    if students_sheet.Faltas[faults] > 15:
        students_sheet.Situação[faults] = 'Reprovado por falta'


for naf in students_sheet.index:

    if students_sheet.Situação[naf] == 'Exame Final':
        naf_col[naf] = 100 - ((students_sheet.P1[naf] + students_sheet.P2[naf] + students_sheet.P3[naf])/3)
        naf_col[naf] = round(naf_col[naf] + 0.5)

students_sheet.to_excel("Engenharia de Software - Desafio Iago D'Ippolito.xlsx", sheet_name='engenharia_de_software', index=False)