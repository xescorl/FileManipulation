# la llibreria utilitzada serà openpyxl, per lo vist es la primera que he trobat que
# em permet modificar un excel existent i no crear-ne un de nou
import openpyxl
import math

fitxer = openpyxl.load_workbook(r'C:\Users\ftur.DOMINIOEXPO\Desktop\plantilla etiquetes.xlsx')

full_Dades = fitxer.worksheets[0]
full_Enganxines = fitxer.worksheets[1]

# Dades de la plantilla
ALTURA_NOM = 0
ALTURA_CHFUNCIO = 1
ALTURA_PRIMERCANAL = 2
COLUMNES_PER_DISPOSITIU = 2
DISPOSITIUS_PER_FILA = 5

# Dades llegides
nom_dispositius = full_Dades['A2'].value
quantitat_dispositius = full_Dades['B2'].value
quantitat_canals = 0

nom_canals = []

for col in range(3, 18):  # des de la columna C (3) fins a la R (19)
    canal = full_Dades.cell(row=2, column=col).value
    if canal:
        nom_canals.append(canal)
        quantitat_canals += 1

# Dades calculades
MAX_COL = COLUMNES_PER_DISPOSITIU * DISPOSITIUS_PER_FILA
filesXetiqueta = ALTURA_PRIMERCANAL + quantitat_canals
MAX_ROW = math.ceil(quantitat_dispositius/DISPOSITIUS_PER_FILA) * filesXetiqueta

# *******************************************************************
# Ara que hem agafat les dades, podem començar a crear les etiquetes
# *******************************************************************

# Primer de tot ajustem el format de les cel·les per a que s'adaptin al contingut
for row in full_Enganxines.iter_rows(min_row=1, max_row=MAX_ROW, min_col=1, max_col=MAX_COL):
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
        cell.font = openpyxl.styles.Font(name='Arial', size=10)
        cell.fill = openpyxl.styles.PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                             right=openpyxl.styles.Side(style='thin'),
                                             top=openpyxl.styles.Side(style='thin'),
                                             bottom=openpyxl.styles.Side(style='thin'))