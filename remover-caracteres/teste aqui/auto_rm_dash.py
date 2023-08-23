import os
from time import sleep
try:
    folder = os.getcwd() + "\\"

    for file_name in os.listdir(folder):
        if 'auto_rm_dash' in file_name:
            continue
        old_name = folder + file_name
        new_name = folder + \
            file_name.replace('-', '').replace('–',
                                               '').replace('  ', ' ').replace('_', ' ')
        os.rename(old_name, new_name)

    print('-'*40)
    print('Programa Finalidado!!!')
    print('Os arquivos foram renomeados corretamente')
    print('-'*40)
    sleep(3)
except:
    print('[ERROR] - Houve um erro inesperado na execução!!!')
    sleep(3)
