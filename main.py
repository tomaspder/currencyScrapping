from openpyxl import load_workbook
from requests import get

#Funcion para hacer scrapping del valor actualizado del dolar
def scrap():
    url = 'https://www.bna.com.ar/Personas'
    #Peticion HTTP, devuelve 200 si todo salio ok.
    response = get(url)
    #Extraigo la posicion exacta donde se ubica la cotizacion del usd y le doy formato decimal
    usd_value = float((response.text[31516:31521]).replace(',','.'))
    return usd_value

#Abro el archivo xl y lo leo
wb = load_workbook(filename = 'FileXL.xlsx')
#Modifico la celda correspondiente al valor del dolar actualizado de la funcion scrap()
wb['Table 1']['C5'] = scrap()
wb.save(filename = 'FileXL.xlsx')
wb.close()

