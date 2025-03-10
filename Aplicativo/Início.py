import streamlit as st
import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, askdirectory
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
# Use para rodar: python -m streamlit run Início.py

#criar funções de carregamento de dados

#prepara visulizações

#criar interface gráfica

st.set_page_config(
    page_title="Grafaria",
    page_icon="🔎",
)

st.title("Página Inicial")
st.sidebar.success("Selecione a função desejada")

st.write(""" 
# Saudações do Grafaria!
    Este site foi desenvolvido a fim de automatizar processos de geração de gráficos
    e processamento de arquivos.

## Como o site funciona?
    Ao lado esquerdo estão as páginas acessíveis referentes à cada uma das 
    funcionabilidades do site. Em cada aba haverá uma explicação sobre a 
    última atualização lançada.
         
## Quais são as funções disponíveis?
    • Upload de arquivos DTA para conversão para XLSX
    • Plot de gráficos de Impedância x Tempo decorrido
         - Só aceita arquivos XLSX
    • Plot de gráficos Z real x Z Imaginário
         - Só aceita arquivos XLSX


Desenvolvido por Maria Eduarda Nunes em fevereiro de 2025
""") #Markdown

