# la llibreria utilitzada serà openpyxl, per lo vist es la primera que he trobat que
# em permet modificar un excel existent i no crear-ne un de nou
import openpyxl

fitxer = openpyxl.load_workbook(r'C:\Users\ftur.DOMINIOEXPO\Desktop\test.xlsx')

# *******************************************************************
# Eliminar totes les fulles excepte la primera
# *******************************************************************

# Guardar la referència de la primera fulla
primera_fulla = fitxer.worksheets[0]

# Obtenir una llista de totes les fulles excepte la primera
fulles_a_eliminar = fitxer.worksheets[1:]

# Eliminar totes les fulles excepte la primera
for fulla in fulles_a_eliminar:
    fitxer.remove(fulla)
    print(f"Fulla eliminada: {fulla.title}")

print(f"Fulles restants: {len(fitxer.worksheets)}")
print(f"Fulla conservada: {fitxer.worksheets[0].title}")

# *******************************************************************
# Finalment, guardem i tanquem el fitxer Excel
# *******************************************************************

# Guardar els canvis al fitxer Excel
fitxer.save(r'C:\Users\ftur.DOMINIOEXPO\Desktop\test.xlsx')
print("Fitxer guardat correctament!")

# Opcional: Tancar el fitxer
fitxer.close()