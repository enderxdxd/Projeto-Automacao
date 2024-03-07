import pyautogui
pyautogui.FAILSAFE = False
import time
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
import pandas as pd
from unidecode import unidecode
import os
import win32com.client as win32
outlook= win32.Dispatch("outlook.application")
from tkinter import filedialog

pyautogui.press("win")
pyautogui.write("Edge")
time.sleep(2)
pyautogui.press("enter")
time.sleep(6)
link = "https://app.pactosolucoes.com.br/login/95f31d847d04ae6b792b1f4394789315/tko=b128e172-4006-41d9-b85f-b4329dc6e264"
pyautogui.write(link)
pyautogui.press("enter")
time.sleep(7)
pyautogui.click(x=1127, y=334)
pyautogui.write("henrique")
time.sleep(5)
pyautogui.click(x=1093, y=385)
pyautogui.write("Henpa678876!")
pyautogui.click(x=1176, y=459)
pyautogui.click(x=1008, y=447)
pyautogui.click(x=1110, y=548)
time.sleep(4)
pyautogui.click(x=657, y=318)
pyautogui.click(x=657, y=357)
time.sleep(2)
pyautogui.click(x=426, y=410)
time.sleep(15)
pyautogui.click(x=716, y=121)
time.sleep(2)
pyautogui.write("Colaborador")
time.sleep(3)
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=860, y=224)
time.sleep(2)
pyautogui.click(x=857, y=406)
time.sleep(2)
pyautogui.click(x=1049, y=225)
time.sleep(2)
pyautogui.click(x=1040, y=250)
time.sleep(3)
pyautogui.click(x=1123, y=264)
pyautogui.click(x=886, y=739)
time.sleep(5)

#ZapSign
pyautogui.press("win")
time.sleep(3)
pyautogui.write("edge")
time.sleep(3)
pyautogui.press("enter")
time.sleep(5)
link_2 = "https://app.zapsign.com.br/acesso/entrar"
pyautogui.write(link_2)
pyautogui.press("enter")
time.sleep(8)
pyautogui.click(x=1241, y=365)
time.sleep(3)
pyautogui.scroll(-350)
time.sleep(3)
pyautogui.click(x=494, y=535)
time.sleep(2)
pyautogui.click(x=598, y=396)
time.sleep(3)
pyautogui.click(x=867, y=472)
time.sleep(3)
pyautogui.click(x=1359, y=0)
time.sleep(5)

pyautogui.click(x=754, y=740)
time.sleep(4)
pyautogui.click(x=79, y=316)
time.sleep(2)
pyautogui.doubleClick(x=408, y=235)
time.sleep(4)
pyautogui.click(x=612, y=426)
time.sleep(3)

pyautogui.click(x=1164, y=85)
time.sleep(4)
pyautogui.click(x=26, y=53)
time.sleep(2)
pyautogui.click(x=92, y=385)
time.sleep(2)
pyautogui.click(x=692, y=145)
time.sleep(3)
pyautogui.write("Colaborador")
time.sleep(3)
pyautogui.click(x=1235, y=200)
time.sleep(2)
pyautogui.click(x=705, y=444)
time.sleep(3)
pyautogui.click(x=1365, y=0)




def processar_arquivos(filepaths):
    nomes_planilha1 = set()
    nomes_planilha2 = set()

    for filepath in filepaths:
        if 'Colaborador' in filepath:
            # Verifica se o arquivo é XLS e converte para XLSX se necessário
            if filepath.lower().endswith('.xls'):
                xls_df = pd.read_excel(filepath, header=None, usecols=[2])
                xlsx_filepath = filepath + 'x'
                xls_df.to_excel(xlsx_filepath, index=False, header=None)
                filepath = xlsx_filepath

            # Adiciona os nomes da primeira planilha (Colaboradores, coluna C) ao conjunto
            df = pd.read_excel(filepath, header=None, usecols=[2])
            nomes_planilha1.update(unidecode(str(nome)) for nome in df[2].dropna())

            
            
            
        elif 'ZapSign' in filepath:
            # Adiciona os nomes da segunda planilha (Signatarios, coluna D) ao conjunto
            df = pd.read_excel(filepath, sheet_name='Signatarios', header=None, usecols=[3])
            nomes_planilha2.update(unidecode(str(nome)) for nome in df[3].dropna())

    # Calcula os diferentes elementos entre as duas planilhas
    different_elements = nomes_planilha1.difference(nomes_planilha2)

    # Aqui você pode fazer o que quiser com os resultados, como imprimir ou salvar
    print("Elementos Diferentes entre as Planilhas:", different_elements)

    return different_elements

class App:
    def __init__(self, master, processar_callback):
        self.master = master
        self.master.title("Analisador de Excel")
        self.processar_callback = processar_callback
        self.master.attributes('-fullscreen', False)

        self.btn_analyze = tk.Button(master, text="Analisar", command=self.analyze_files)
        self.btn_analyze.pack(pady=20)

        self.btn_excel = tk.Button(master, text="Excel", command=self.save_to_excel)
        self.btn_excel.pack(pady=20)

        self.resultados = None

        # Chama automaticamente a função de análise e salva ao iniciar
        self.auto_analyze_and_save()

    def auto_analyze_and_save(self):
        # Modifique o diretório conforme necessário
        diretorio = r"C:\Users\FLEX\Downloads"

        # Encontrar automaticamente os dois arquivos mais recentes
        filepaths = self.encontrar_dois_mais_recentes(diretorio)

        # Se encontrou dois arquivos, realiza a análise automaticamente
        if filepaths:
            self.resultados = self.processar_callback(filepaths)
            print("Análise automática concluída.")
            self.save_to_excel()
        else:
            print("Não foram encontrados dois arquivos para análise automática.")

    def encontrar_dois_mais_recentes(self, diretorio):
        arquivos = [os.path.join(diretorio, arquivo) for arquivo in os.listdir(diretorio)]
        arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.lower().endswith('.xlsx')]
        arquivos_xlsx.sort(key=os.path.getmtime, reverse=True)

        return arquivos_xlsx[:2] if len(arquivos_xlsx) >= 2 else None

    def analyze_files(self):
        print("Análise manual não é mais necessária.")

    def save_to_excel(self):
        if not self.resultados:
            print("Nenhum resultado para salvar em Excel.")
            return

        # Modifique o caminho conforme necessário
        output_filepath = r"C:\Users\FLEX\Documents\novoarquivo\PERSONALSEMCONTRATO.xlsx"

        new_workbook = Workbook()
        new_sheet = new_workbook.active

        new_sheet['A1'] = 'PERSONALSEMCONTRATO'

        for index, element in enumerate(self.resultados, start=2):
            new_sheet.cell(row=index, column=1, value=element)

        new_workbook.save(output_filepath)
        print(f"Arquivo Excel gerado com sucesso em: {output_filepath}")

        # Adicione aqui o código para fechar o aplicativo após salvar o arquivo, se desejado
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root, processar_arquivos)
    root.mainloop()






mail = outlook.CreateItem(0)
mail.To = 'teste@hotmail.com"
mail.Subject = "Segue o anexo"
mail.Body = 'Anexo excel'

attachment = r'C:\Users\FLEX\Documents\novoarquivo\PERSONALSEMCONTRATO.xlsx'
mail.attachments.Add(attachment)

mail.Send()










