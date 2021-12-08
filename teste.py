from tkinter import *
import clipboard as cb
import tkinter.font as font
import pyautogui as pg
from tkinter import messagebox
from datetime import date
import pandas as pd
from openpyxl import *
import time
import sys

import re

mat = '1431862'
messagebox.showwarning(title='Erro', message='Nao foi possivel alterar a matricula '+mat)


    # Finalizando laço de repetição



    # pg.write("*/frame.querySelector('input[name=\"txt_cd_lot\"]').value")

# def alterar_linha(arquivo,nome):
#     cont = 0
#     with open(arquivo,'r') as f:
#         texto=f.readlines()
#     with open(arquivo,'w') as f:
#         exist = 0
#         for i in texto:
#             cont = cont + 1
#             if(cont == 11):
#                 f.write(i)
#                 messagebox.showwarning("Error","limite de atestados exedido, favor envialos a junta e somente apos enviados, adicionar novos.")
#                 return True
#             if(str(nome.strip()) in str(texto)):
#                 exist = 1
#             f.write(i)
#         messagebox.showinfo("Registrado", "Ok, pode guardar o atestado.")
#         if(exist == 1):
#             return True
#         f.write(str(nome)+'\n')
#         return True

# def cadastrar():
#     nome = entry.get()
#     alterar_linha('bd.txt', nome)
#     entry.delete('0', END)

# def limpar():
#     open('bd.txt', 'w').close()

# lista = ['POS1', 'POSIC1', 'POS3', 'POS4', 'POS5', 'POS6', 'POS7', 'POS8', 'POS9', 'POSI0', 'POSI1', 'POSIC2']
# doc = Document('K:/Administrativo/SetorPessoal/Raynder/Bot/arquivos_bases/junta.docx')

# def replaceWord(word, replace):
#     num = entry.get()
#     for paragraph in doc.paragraphs:
#         if word in paragraph.text:
#             inline = paragraph.runs
#             for i in range(len(inline)):
#                 if word in inline[i].text:
#                     text = inline[i].text.replace(word, replace)
#                     inline[i].text = text 
#     strcod = str(codigo)
#     strdia = str(dia)
#     strmes = str(mes)
#     strano = str(ano)
#     doc.save("K:/Administrativo/SetorPessoal/Raynder/Atestados/Remessas/relatorio"+strcod+"Junta"+strdia+"-"+strmes+"-"+strano+".doc")

# def gerarRelatorio():
#     novocodigo = int(codigo)+1
#     arquivo2 = open('K:/Administrativo/SetorPessoal/Raynder/Atestados/RegistroAtestados/numAtestado.txt', 'w')
#     arquivo2.write(str(novocodigo)+'\n')

#     contger = 0
#     with open('bd.txt','r') as f:
#         texto=f.readlines()
#     with open('bd.txt','w') as f:
#         for i in texto:
#             replaceWord(lista[contger], i.strip())
#             contger = contger + 1
#         replaceWord('DIA', str(dia))
#         replaceWord('MZ', str(meses[mes]))
#         replaceWord('AZ', str(ano))
#         replaceWord('RZ', str(codigo))

#         for c in lista:
#             replaceWord(c, "")

# myFont = font.Font(size=30)
# myFont2 = font.Font(size=15)

# C = Canvas(app, bg="blue", height=250, width=300)
# filename = PhotoImage(file = "Paisagem.png")
# background_label = Label(app, image=filename)
# background_label.place(x=0, y=0, relwidth=1, relheight=1)

# font1 = font.Font(name='TkCaptionFont', exists=True)
# font1.config(family='courier new', size=20)

# entry = AutocompleteEntry(autocompleteList, app, listboxLength=6, width=32, matchesFunction=matches)
# entry.place(x=150, y=45, width=250, height=20)
# button3 = Button(text='Cadastrar', bg='#145D3B', fg='#ffffff' , command=cadastrar)#bg fundo fg font
# button3['font'] = myFont
# button3.place(x=165, y=70, width=200, height=70)


# button3 = Button(text='Gerar Documento', bg='#FF0000', fg='#ffffff' , command=gerarRelatorio)#bg fundo fg font
# button3['font'] = myFont2
# button3.place(x=410, y=340, width=170, height=35)

# app.mainloop()