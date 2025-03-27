import pandas as pd # Planilhas
import pyperclip # Armazenar dados na área de transferência
import os # Arquivos
from time import sleep # Pausas
import pyautogui as pg # Bot

#Função pra esperar enquanto determinada imagem não aparece
def detectarImagem(caminho):
    img = pg.locateOnScreen(caminho, confidence=0.7, minSearchTime=10) # Confidence = Porcentagem de semelhança aturada
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
origem = os.path.join(os.getcwd(), 'planilhas') # Pasta aonde estão as planilhas
links = [
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=l7lACH&nav=MTVfezkwRDM4QzdELUJGQzktNDZDNS1BMDEwLThGQUJENzQ1RUNDQX0",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=zODa77&nav=MTVfezM4ODBERkVFLURFRkMtNDAzRC1BNTJELTU1QjY2QjkyQzBFN30",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=c9XgrX&nav=MTVfezIzN0QwMTBBLUQ2MzYtNDRDQy1CRTNFLTdBQzI4RUQxRkZDQ30",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=LhDs5N&nav=MTVfezhGMzFBMTJCLTRFRjAtNDU5MC1BRkFGLUEzMTg2OTZCRDUxNn0",
"https://royalagrocereais-my.sharepoint.com/:x:/g/personal/vitor_ramos_royalagro_com_br/EV7wPKMhqrdBmLTM7ImqXvQBWv7-ILLdlYKvbIRlRWVhWA?e=8gZfKQ&nav=MTVfezY4RTgyM0NELTA1MTQtNDE0RS04MkQwLUU0MjExMzNGQTQ5MX0"
] # Links em ordem já preestabelecida, pode ser substituido por um dicionario futuramente
for arquivo in os.listdir(origem): #For pra excluir todos os arqquivos que não estejam filtrados
    caminho_arquivo = os.path.join(origem, arquivo)
    if 'Filtrado' not in arquivo and os.path.isfile(caminho_arquivo):
        try:
            os.remove(caminho_arquivo)
            print(f'Arquivo {arquivo} excluído com sucesso.')
        except Exception as e:
            print(f'Erro ao excluir o arquivo {arquivo}: {e}')
            
for caminho, subpasta, arquivos in os.walk(origem):
    for contador,(nome) in enumerate(arquivos): # UMA EXECUÇÃO POR ARQUIVO
        print(contador)
        arq = caminho+"\\"+nome
        pg.hotkey('ctrl', 't')
        pg.hotkey('ctrl', 'l')
        pyperclip.copy(links[contador])
        print(arq)
        pg.hotkey('ctrl', 'v')
        pg.press('enter')
        img = detectarImagem('imagens/selectAll.png')
        pg.press('home')
        pg.click(img)
        img = detectarImagem('imagens/allSelected.png')
        pg.press('delete')
        img = detectarImagem('imagens/allDelete.png')
        copiar_para_area_de_transferencia(arq)
        sleep(2)
        pg.hotkey('ctrl', 'v')
        img = detectarImagem('imagens/pass.png')
        
