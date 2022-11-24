import os
import pandas as pd
import getpass
# os.chdir("C:\\Users\\"+ str(getpass.getuser()) +"\\Box\\..")
filen = input("Enter the file name: ") +'.xlsx'
file_opn = open(filen, mode='r+b')
data = pd.read_excel(file_opn)
data2 = data['SOURCE'].str.split('-', expand=True)
new = pd.concat([data, data2], axis=1)
new["Header"] = ''
new["ROW2"] = ''
new["COLUMN1"] = ''
new["COLUMN2"] = ''
new["VARIABLE_KEY1"] = ''
new["VARIABLE_VALUE1"] = ''
new["FIELDTYPE"] = ''
new["Extracted"] = ''
new["Output"] = ''
new.rename(columns = {0 : 'DATATYPE', 1 :'FILENAME', 2:'HEADER1', 3:'HEADER2', 4:'HEADER', 5:'ROW', 6:'ROW1'}, inplace=True)
new.to_excel('C:\\Users\\'+ str(getpass.getuser()) +'\\Downloads\\YourFile.xlsx', index=False)
