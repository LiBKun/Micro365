import pandas as pd # Planilhas
import pyperclip # Armazenar dados na área de transferência
import os # Arquivos
from time import sleep # Pausas
import pyautogui as pg # Bot

#Função pra esperar enquanto determinada imagem não aparece
def detectarImagem(caminho):
    img = pg.locateOnScreen(caminho, confidence=0.7) # Confidence = Porcentagem de semelhança aturada
    while img == None:
        img = pg.locateOnScreen(caminho, confidence=0.7)
        sleep(0.1)
    return img

#Função pra copiar dados de uma determinada planilha
def copiar_para_area_de_transferencia(file_path):
    df = pd.read_excel(file_path)
    dados_como_string = df.to_csv(sep='\t', index=False)
    pyperclip.copy(dados_como_string)

pg.hotkey('alt', 'tab')
origem = r"" # Pasta aonde estão as planilhas
links = [
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=l7lACH&nav=MTVfezkwRDM4QzdELUJGQzktNDZDNS1BMDEwLThGQUJENzQ1RUNDQX0",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=zODa77&nav=MTVfezM4ODBERkVFLURFRkMtNDAzRC1BNTJELTU1QjY2QjkyQzBFN30",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=c9XgrX&nav=MTVfezIzN0QwMTBBLUQ2MzYtNDRDQy1CRTNFLTdBQzI4RUQxRkZDQ30"
] # Links em ordem já preestabelecida, pode ser substituido por um dicionario futuramente
for caminho, subpasta, arquivos in os.walk(origem):
    for contador,(nome) in enumerate(arquivos): # UMA EXECUÇÃO POR ARQUIVO
        arq = caminho+"\\"+nome
        copiar_para_area_de_transferencia(arq)
        pg.hotkey('ctrl', 't')
        pg.hotkey('ctrl', 'l')
        pg.write(links[contador])
        pg.press('enter')
        img = detectarImagem('imagens/selectAll.png')
        pg.click(img)
        img = detectarImagem('imagens/allSelected.png')
        pg.press('delete')
        sleep(3)
        pg.hotkey('ctrl', 'v')
        img = detectarImagem('imagens/pass.png')
