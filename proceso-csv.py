import os
import sys
import csv
import pandas as pd
import re
import unicodedata 

def elimina_tildes(cadena):
    s = ''.join((c for c in unicodedata.normalize('NFD',cadena) if unicodedata.category(c) != 'Mn'))
    return s

"""
def normalizar(cadena):
    # -> NFD y eliminar diacríticos
    cadena = re.sub(
            r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1", 
            unicodedata.normalize( "NFD", cadena), 0, re.I
        )

    # -> NFC
    return unicodedata.normalize( 'NFC', cadena)    
"""

#C://Descargas/
read_file = pd.read_excel('A_GSuite/Usuarios-GSuite-Pasados.xlsx', sheet_name='Hoja1')
read_file.to_csv('A_GSuite/Usuarios-GSuite-Pasados.csv', index = None, header=True)

path_origen = 'A_GSuite/Usuarios-GSuite-Pasados.csv'
path_destino_gfe = 'A_GSuite/Usuarios-GSuite-Para_Importar.csv'
path_destino_ps = 'A_GSuite/Usuarios-GSuite-Para_PowerShell.csv'

cuentas_from_csv = open(path_origen,mode='r',encoding='utf-8')
cuenta_to_gfe = open(path_destino_gfe,mode='w',encoding='utf-8', newline='')
cuenta_to_ps = open(path_destino_ps,mode='w',encoding='utf-8', newline='')

cuentas_csv = csv.reader(cuentas_from_csv)
cuentas_gfe_writer = csv.writer(cuenta_to_gfe, delimiter=',')
cuentas_ps_writer = csv.writer(cuenta_to_ps, delimiter=',')

#Escribo encabezado gfe
header_gfe = ['First Name [Required]','Last Name [Required]','Email Address [Required]','Password [Required]','Password Hash Function [UPLOAD ONLY]','Org Unit Path [Required]','New Primary Email [UPLOAD ONLY]','Recovery Email','Home Secondary Email','Work Secondary Email','Recovery Phone [MUST BE IN THE E.164 FORMAT]','Work Phone','Home Phone','Mobile Phone','Work Address','Home Address','Employee ID','Employee Type','Employee Title','Manager Email','Department','Cost Center','Building ID','Floor Name','Floor Section','Change Password at Next Sign-In','New Status [UPLOAD ONLY]']
cuentas_gfe_writer.writerow(header_gfe)

#Escribo encabezado ps
header_ps = ['nombres','apellido','cuenta','dpto','mail']
cuentas_ps_writer.writerow(header_ps)

#Paso las cuentas
line_out_gfe =  ['name', 'surname','mail','Temporal123','','/SA/','','mail','','mail','','','','','','','','','','','','','','','','true','']

#Cuento las lineas para evitar el encabezado original
line_count = 0

#Skip header row
next(cuentas_csv)  

#Linea para envio de mails
line_out_ps = ['nombres','apellido','cuenta','dpto','mail']

for row in cuentas_csv:
        print(row)
        line_out_gfe[0] =  elimina_tildes(row[0].strip())
        line_out_gfe[1] =  elimina_tildes(row[1].strip())
        line_out_gfe[2] =  row[2].strip()
        line_out_gfe[5] =  row[3].strip()
        line_out_gfe[7] =  row[4].strip()
        line_out_gfe[9] =  row[4].strip()
        cuentas_gfe_writer.writerow(line_out_gfe)

        line_out_ps[0] =  elimina_tildes(row[0].strip())
        line_out_ps[1] =  elimina_tildes(row[1].strip())
        line_out_ps[2] =  row[2].strip()
        line_out_ps[3] =  row[3].strip()
        line_out_ps[4] =  row[4].strip()
        cuentas_ps_writer.writerow(line_out_ps)

print('************************ PROCESO TERMINADO ************************')

files = os.listdir('A_GSuite/')

print('************************ ARCHIVOS GENERADOS ************************')
print('Lista de archivos (Carpeta: /A_GSuite/):')
for file in files:
    print(file)

cuentas_from_csv.close()
cuenta_to_gfe.close()
cuenta_to_ps.close()
