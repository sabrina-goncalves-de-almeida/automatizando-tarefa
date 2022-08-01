# biblioteca que chama processos feitos pelo mouse ou teclado
import pyautogui

# biblioteca que chama funções como copy (ctrl+c)
import pyperclip

# biblioteca que executa processos de tempo
import time

import pandas as pd

# cada vez que pyautogui for chamado vai esperar 1 segundo para executar a função
pyautogui.PAUSE = 1

# pressionar a tecla windows
pyautogui.press("win")

# escrever google no menu
pyautogui.write("google")

# pressionar a tecla enter
pyautogui.press("enter")

# é um copiar, equivale a pressionar as teclas ctrl+c
pyperclip.copy("link para a pasta do drive que contem o arquivo excel")

# pressionar as tecla ctrl+v
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# leva 5 segundos antes de passar para o próximo processo
time.sleep(5)

# para printar a posição que se quer clicar no terminal
# print(pyautogui.position())

# vai clicar na posição da tela com eixo x = n; e eixo y = m.
pyautogui.click(x=340, y=335, clicks=2)
time.sleep(2)

pyautogui.click(x=340, y=335)

pyautogui.click(x=1169, y=234)

pyautogui.click(x=989, y=611)
time.sleep(7)

# o r vai usar o texto exatamente como foi escrito, sem atalhos ou caracteres inesperados
tabela = pd.read_excel(r"path-para-seu-arquivo/nome-do-arquivo")

# imprime a tabela
print(tabela)

# salva uma coluna com seus valores somados na variável
quantidade = tabela["Quantidade"].sum()

faturamento = tabela["Valor Final"].sum()

print(quantidade)
print(faturamento)

# abrir nova guia do google
pyautogui.hotkey("ctrl", "t")

# copiar o email
pyperclip.copy("link para seu email, lembre pegar o link após abrir uma caixa de mensagem nova")

# colar o email
pyautogui.hotkey("ctrl", "v")

# abrir o email
pyautogui.press("enter")

# tempo de espera 7 segundos
time.sleep(7)

# escrever o email e passar para o próximo campo
pyautogui.write("email do destinatario")
pyautogui.press("tab")
pyautogui.press("tab")

# escrever o assunto do email e passar para o próximo campo

    # o pyautogui.write não escreve caracteres especiais, por isso deve usar o pyperclip.copy
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

# escreve o conteúdo do email
texto = f""" Prezado gerente, bom dia
Venho por meio deste informar, que ontem,
o faturamento da empresa foi de: R$ {faturamento:,.2f},
e a quantidade de produtos vendidos foi: {quantidade:,}.

Agradecemos por sua atenção.
Assin.: Equipe de vendas.
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# envia o email
pyautogui.hotkey("ctrl","enter")