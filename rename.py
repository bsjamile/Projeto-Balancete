import os
folder = r'C:\Users\jamile.santos\Downloads\teste\\'

lista_nomes = [
    'Balancete CD NÉOS.xlsx',
    'Balancete CD PE.xlsx',
    'Balancete CD RN.xlsx'
]

i = 0

for file_name in os.listdir(folder):
    old_name = folder + file_name
    new_name = folder + lista_nomes[i]
    os.rename(old_name,new_name)
    i+=1

print(os.listdir(folder))