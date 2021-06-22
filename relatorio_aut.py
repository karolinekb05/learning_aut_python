#!/usr/bin/env python
# coding: utf-8

# In[40]:


import pyautogui
import time
import pyperclip

# pyautogui e pyperclip automatizam teclado e mouse

pyautogui.PAUSE = 0.5

# 1
#Abrir nova aba
pyautogui.hotkey('ctrl', 't')
# Entrar no link do sistema
link = "https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#2
#Entrar na pasta
time.sleep(5) #tempo de espera
#pyautogui.position() -> mostra a posição do mouse(x,y)
pyautogui.click(398, 385) #clicar no arquivo
pyautogui.click(1185, 179) #clicar nos ... ou button'right'
pyautogui.click(1045, 516) #fazer download
time.sleep(8)

#3
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Karoline\Downloads\Vendas - Dez.xlsx") #sempre colocar o r antes do caminho do arquivo

#4
#calculando faturamento e quantidade vendida
faturamento = tabela["Valor Final"].sum() #soma da coluna Valor FInal
quantidade = tabela["Quantidade"].sum() #soma da coluna Quantidade
#exibindo a tabela e os resultados do cálculo
display(tabela)
print(faturamento) #exibe o resultado
print(quantidade)

#5
pyautogui.hotkey('ctrl', 't')
link = "https://mail.google.com/mail/u/0/#inbox?compose=new"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(10)

pyautogui.write('karolinekb05@gmail.com') #destinatário
pyautogui.press('tab')
time.sleep(2)
pyautogui.press('tab')
assunto = 'Relatório de Vendas' #assunto
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
#corpo do e-mail
texto_email = f"""
Prezados, bom dia

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}
"""
pyperclip.copy(texto_email)
pyautogui.hotkey('ctrl', 'v')

pyautogui.hotkey('ctrl', 'enter') #enviar


# In[36]:





# In[26]:





# In[ ]:




