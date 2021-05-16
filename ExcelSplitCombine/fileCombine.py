import glob
import os
import pandas as pd

file=input('Ruta del archivo: ')
pth=os.path.dirname(file)
extension = os.path.splitext(file)[1]
files = glob.glob(os.path.join(pth, '*.xls*'))
newfile=os.path.join(pth,'combinado.xlsx')
df = pd.DataFrame()
for f in files:
    data = pd.read_excel(f)
    df = df.append(data)

df.to_excel(newfile, sheet_name='combinado', index=False)
print('Completado 😁')