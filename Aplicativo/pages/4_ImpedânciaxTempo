import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime

# Função para carregar a planilha

def carregar_planilha(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, header=None)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        return None

# Função para obter a data e a hora

def obter_data_hora(df):
    try:
        data = df.iloc[0, 1]  # Célula B0
        hora = df.iloc[1, 1]  # Célula B1

        if isinstance(data, str):
            data = pd.to_datetime(data, format="%d/%m/%Y", errors='coerce')
        elif not isinstance(data, datetime):
            data = pd.to_datetime(data, dayfirst=True, errors='coerce')

        if pd.isna(hora) or hora is None:
            return None

        hora = str(hora).strip()
        
        try:
            hora = datetime.strptime(hora, "%H:%M:%S").time()
        except ValueError:
            try:
                hora = datetime.strptime(hora, "%H:%M").time()
            except ValueError:
                return None

        return datetime.combine(data.date(), hora)
    except Exception:
        return None

# Função para selecionar os dados

def selecionar_dados(df, data_hora_inicial, area):
    try:
        ultima_linha_preenchida = df.iloc[:, 0].last_valid_index()
        if ultima_linha_preenchida is None:
            return None, None

        valor_coluna_a = df.iloc[ultima_linha_preenchida, 0]
        data_hora_atual = obter_data_hora(df)
        if data_hora_atual is None:
            return None, None

        horas_decorridas = max(0, (data_hora_atual - data_hora_inicial).total_seconds() / 3600)
        eixo_y = max(0, valor_coluna_a * area)

        return horas_decorridas, eixo_y
    except Exception:
        return None, None

# Função para gerar o gráfico

def gerar_grafico_combinado(dados_graficos, titulo_grafico_combinado):
    try:
        cores = ['#FF4500', '#FF8C00', '#DAA520', '#D10D0D', '#6E0303', '#520B0B']
        marcadores = ['o']

        dados_graficos = sorted(dados_graficos, key=lambda x: x[0])

        plt.figure(figsize=(10, 6))
        
        for index, (eixo_x, eixo_y, legenda) in enumerate(dados_graficos):
            marcador = marcadores[index % len(marcadores)]
            cor = cores[index % len(cores)]
            plt.scatter(eixo_x, eixo_y, marker=marcador, color=cor, label=legenda)

        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.title(titulo_grafico_combinado)
        plt.xlabel("Tempo (horas)")
        plt.ylabel("Zreal (Ohm.cm²)")
        plt.tight_layout()

        st.pyplot(plt)
    except Exception as e:
        st.error(f"Erro ao gerar gráfico: {e}")

# Interface do Streamlit

def main():
    st.title("Gerador de Gráfico Combinado")
    st.write("Faça upload dos arquivos Excel para processar os dados e gerar o gráfico combinado.")

    uploaded_files = st.file_uploader("Selecione os arquivos Excel", type=["xlsx", "xls"], accept_multiple_files=True)
    titulo_grafico_combinado = st.text_input("Título do Gráfico Combinado", "Gráfico Combinado")
    area = st.number_input("Área do Corpo de Prova", min_value=0.0, value=1.0, step=0.1)

    if st.button("Gerar Gráfico"):
        if not uploaded_files:
            st.warning("Por favor, faça upload de pelo menos um arquivo Excel.")
            return

        datas_horas = []
        dados_graficos = []
        
        for file in uploaded_files:
            df = carregar_planilha(file)
            if df is not None:
                data_hora = obter_data_hora(df)
                if data_hora:
                    datas_horas.append((data_hora, file.name))

        if not datas_horas:
            st.error("Nenhuma data/hora válida foi encontrada nos arquivos.")
            return

        data_hora_inicial, _ = min(datas_horas, key=lambda x: x[0])

        for file in uploaded_files:
            df = carregar_planilha(file)
            if df is not None:
                eixo_x, eixo_y = selecionar_dados(df, data_hora_inicial, area)
                if eixo_x is not None and eixo_y is not None:
                    dados_graficos.append((eixo_x, eixo_y, file.name))

        if dados_graficos:
            gerar_grafico_combinado(dados_graficos, titulo_grafico_combinado)
        else:
            st.error("Nenhum dado válido encontrado nos arquivos.")

if __name__ == "__main__":
    main()

