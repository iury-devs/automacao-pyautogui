import openpyxl
import pyautogui

planilha_vendas = openpyxl.load_workbook('venda_de_produto.xlsx')
pagina_vendas = planilha_vendas['vendas']

for linha in pagina_vendas.iter_rows:
    #Nome cliente
    pyautogui.click(1808, 452, duration=1.5)
    pyautogui.write(linha[0].value)
    #Produto
    pyautogui.click(1815, 476, duration=1.5)
    pyautogui.write(linha[1].value)
    #Quantidade
    pyautogui.click(1813, 497, duration=1.5)
    pyautogui.write(str(linha[2].value))
    #Categoria
    pyautogui.click(1883, 532, duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(1752, 549, duration=1.5)
    pyautogui.click(1256, 581, duration=1.5)

    
    
    
   