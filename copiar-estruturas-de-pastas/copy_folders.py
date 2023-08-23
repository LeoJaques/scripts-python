import os
from time import sleep

def copy_structure(src_folder, dst_folder):
    if not os.path.exists(dst_folder):
        os.makedirs(dst_folder)
    for item in os.listdir(src_folder):
        s = os.path.join(src_folder, item)
        d = os.path.join(dst_folder, item)
        if os.path.isdir(s):
            copy_structure(s, d)
        else:
            pass  # arquivos são ignorados

src_folder = input("Digite o caminho completo da pasta base: ")
folders = [folder for folder in os.listdir(src_folder) if os.path.isdir(os.path.join(src_folder, folder))]

print("Pastas existentes:")
for i, folder in enumerate(folders):
    print(f"{i+1}. {folder}")

selected = int(input("Selecione a pasta a ser copiada (número): "))
dst_folder = os.path.join(src_folder, f"{folders[selected-1]}-copy")

copy_structure(os.path.join(src_folder, folders[selected-1]), dst_folder)

print("Estrutura de pastas copiada com sucesso.")
sleep(2)