import os
import pandas as pd
import time

start_time = time.time()

# dir_ce = r"C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE"
dir_ce = r"C:\Users\GB675AG\Downloads\excel check\excel check\excel"
list_of_dirs = os.listdir(dir_ce)

for folder in list_of_dirs:
    excel_file_dir = f"{dir_ce}/{folder}"
    list_of_excel_files = os.listdir(excel_file_dir)

    for file_name in list_of_excel_files:
        file_path = f"{excel_file_dir}/{file_name}"
        file_name = os.path.basename(file_path)

        # pd.read_excel(file_path)
        try:
            dataframe1 = pd.read_excel(file_path)
            print("Funcionando:", file_name)
            # if not dataframe1.empty: print('Planilha ok')
        except:
            print("Corrompida:", file_name)

final_time = time.time() - start_time
# print('Tempo gasto:', round(final_time / 60, 3), "m")
print('Tempo gasto:', str(round(final_time, 3)) + "s")
# print('Tempo gasto:', round(final_time * 1000, 3), "ms")
