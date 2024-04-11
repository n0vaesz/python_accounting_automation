import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Enter the spreadsheet
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copy information from a field and paste it into its corresponding field
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = (linha[0].value)
    pyperclip.copy(nome_produto)
    pyautogui.click(402,353,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Descrição
    descrição = (linha[1].value)
    pyperclip.copy(descrição)
    pyautogui.click(220,474,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Categoria
    categoria = (linha[2].value)
    pyperclip.copy(categoria)
    pyautogui.click(209,591,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #Código do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(189,675,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(190,764,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(211,849,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(184,908,duration=1)
    sleep(3)
    
    #Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(208,397,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(210,479,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(225,570,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(205,653,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Tamanho
    tamanho = linha[10].value
    pyautogui.click(282,740,duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(239,770,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(229,797,duration=1)
    else:
        pyautogui.click(265,815,duration=1)

    #Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(274,829,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #Next button
    pyautogui.click(191,882,duration=1)

    #Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(303,415,duration=1)
    pyautogui.hotkey('ctrl','v')

    #País Origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(275,503,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(246,596,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Código de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(311,724,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Localização no armazém
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(300,807,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Finish button
    pyautogui.click(174,865,duration=1)

    #OK button
    pyautogui.click(656,195,duration=1)

    # Add one more
    pyautogui.click(487,626,duration=1)

#PyAutoGUI (click and keyboard automation)
#Openpyxl (Spreadsheet reading and automation)

