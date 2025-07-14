# la llibreria utilitzada serà openpyxl, per lo vist es la primera que he trobat que
# em permet modificar un excel existent i no crear-ne un de nou
import openpyxl
import math

fitxer = openpyxl.load_workbook(r'C:\Users\ftur.DOMINIOEXPO\Desktop\test.xlsx')

full_Dades = fitxer.worksheets[0]
full_Enganxines = fitxer.worksheets[1]

# *******************************************************************
# Lectura de dades del fitxer Excel
# *******************************************************************

# Dades de la plantilla
ALTURA_NOM = 0
ALTURA_CHFUNCIO = 1
ALTURA_PRIMERCANAL = 2
COLUMNES_PER_DISPOSITIU = 2
DISPOSITIUS_PER_FILA = 5
ALTURA_CELA = 15
AMPLADA_CELA_PETITA = 2.71
AMPLADA_CELA_GRAN = 11.86

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

# FORMAT DE LES CEL·LES
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

# Ajustem l'alçada de les files
for row_num in range(1, MAX_ROW + 1):
    full_Enganxines.row_dimensions[row_num].height = ALTURA_CELA

# Ajustem l'amplada de les columnes
for col in range(1, 11):
    sample_cell = full_Enganxines.cell(row=1, column=col)
    col_letter = sample_cell.column_letter
    
    if col % 2 == 0:
        # Columna par - amplada 10
        full_Enganxines.column_dimensions[col_letter].width = AMPLADA_CELA_GRAN
    else:
        # Columna impar - amplada 5
        full_Enganxines.column_dimensions[col_letter].width = AMPLADA_CELA_PETITA

# CONTINGUT DE LES CEL·LES
# *******************************************************************

# Omplim les cel·les amb el contingut corresponent
for n in range(1, quantitat_dispositius + 1):
    columna_actual = ((n - 1) % DISPOSITIUS_PER_FILA) * COLUMNES_PER_DISPOSITIU + 1
    fila_actual = math.floor((n - 1) / DISPOSITIUS_PER_FILA) * filesXetiqueta + 1
    full_Enganxines.cell(row=fila_actual + ALTURA_NOM, column=columna_actual).value = nom_dispositius + str(n)
    full_Enganxines.cell(row=fila_actual + ALTURA_CHFUNCIO, column=columna_actual).value = 'CH'
    full_Enganxines.cell(row=fila_actual + ALTURA_CHFUNCIO, column=columna_actual + 1).value = 'FUNCIO'
    for canal in range(1, quantitat_canals + 1):
        full_Enganxines.cell(row=fila_actual + ALTURA_PRIMERCANAL + canal - 1, column=columna_actual).value = str(canal)
        full_Enganxines.cell(row=fila_actual + ALTURA_PRIMERCANAL + canal - 1, column=columna_actual + 1).value = nom_canals[canal - 1]

# *******************************************************************
# Finalment, guardem i tanquem el fitxer Excel
# *******************************************************************

# Guardar els canvis al fitxer Excel
fitxer.save(r'C:\Users\ftur.DOMINIOEXPO\Desktop\test.xlsx')
print("Fitxer guardat correctament!")

# Opcional: Tancar el fitxer
fitxer.close()