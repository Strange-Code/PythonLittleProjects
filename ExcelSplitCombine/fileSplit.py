# Paquetes a importar
import pandas as pd
import os
from openpyxl import load_workbook
import xlsxwriter
from shutil import copyfile

file=input('Ruta del archivo: ')
extension = os.path.splitext(file)[1]
filename = os.path.splitext(file)[0]
pth=os.path.dirname(file)
newfile=os.path.join(pth,filename+'_2'+extension)
df=pd.read_excel(file)
colpick=input('Seleccione Columna: ')
cols=list(set(df[colpick].values))

def sendtofile(cols):
    for i in cols:
        df[df[colpick] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i, index=False)
    print('\nCompletado 😀')
    print('Gracias por usar este programa. ❤')
    return

def sendtosheet(cols):
    copyfile(file, newfile)
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        for myname in cols:
            mydf = df.loc[df[colpick] == myname]
            mydf.to_excel(writer, sheet_name=myname, index=False)
        writer.save()

    print('\nCompletado 😀')
    print('Gracias por usar este programa. ❤')
    return

print('Tu data sera separada por este valor {} y se creara {} archivos u hojas basadas en la selección. Si estas listo para proceder presiona "Y" y enter. presiona "N" para salir.'.format(', '.join(cols),len(cols)))
while True:
    x=input('Listo para proceder 😎 (Y/N): ').lower()
    if x == 'y':
        while True:
            s = input('Dividir en diferentes hojas o archivos (S/F): ').lower()
            if s == 'f':
                sendtofile(cols)
                break
            elif s == 's':
                sendtosheet(cols)
                break
            else: continue
        break
    elif x=='n':
        print('\nGracias por usar este programa. ❤')
        break

    else: continue