import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = (linha[0].value)
    pyperclip.copy(nome_produto)
    pyautogui.click(402,353,duration=1)
    pyautogui.hotkey('ctrl','v')

    descrição = (linha[1].value)
    pyperclip.copy(descrição)
    pyautogui.click(220,474,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = (linha[2].value)
    pyperclip.copy(categoria)
    pyautogui.click(209,591,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(189,675,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(190,764,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(211,849,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(184,908,duration=1)
    sleep(3)
    
    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value
    cor = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    fabricante = linha[12].value
    pais_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizacao_armazem = linha[16].value

    # 209,591
    # 189,675
    # 190,764
    # 211,849
# Repetir esses passos para outros campos até preencher #odos os campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página (page #)
# Repetir os mesmos passos e finalizar o cadastro daquele #roduto e clicar em concluir
# Clicar em ok, para finalizar o processo
# Clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"

#PyAutoGUI (automação de clicks e teclado)
#Openpyxl (Leitura e automação de planilhas)

