import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
from io import BytesIO

# Criar diretório temporário para histórico de gráficos
HISTORICO_DIR = "historico_graficos"
os.makedirs(HISTORICO_DIR, exist_ok=True)

def carregar_planilha(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, header=None)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        return None

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

st.title("Gerador de Gráfico Impedância X Tempo")
st.write("Faça upload dos arquivos `.xlsx` convertidos na aba Conversor para processar os dados e visualizar o gráfico Impedância X Tempo.")

uploaded_files = st.file_uploader("Selecione os arquivos Excel", type=["xlsx", "xls"], accept_multiple_files=True)

titulo_grafico_combinado = st.text_input("Título do Gráfico", "Gráfico ImpXTempo")
area = st.number_input("Área do Corpo de Prova", min_value=0.0, value=1.0, step=0.1)

mostrar_legenda = st.checkbox("Mostrar legenda no gráfico", value=True)

if uploaded_files:
    datas_horas = []
    dados_graficos = []
    imagens = []
    
    for file in uploaded_files:
        df = carregar_planilha(file)
        if df is not None:
            data_hora = obter_data_hora(df)
            if data_hora:
                datas_horas.append((data_hora, file.name))
    
    if datas_horas:
        data_hora_inicial, _ = min(datas_horas, key=lambda x: x[0])
        
        for file in uploaded_files:
            df = carregar_planilha(file)
            if df is not None:
                eixo_x, eixo_y = selecionar_dados(df, data_hora_inicial, area)
                if eixo_x is not None and eixo_y is not None:
                    dados_graficos.append((eixo_x, eixo_y, file.name))
    
    if dados_graficos:
        plt.figure(figsize=(10, 6))
        for eixo_x, eixo_y, legenda in dados_graficos:
            plt.scatter(eixo_x, eixo_y, label=legenda, color='#D10D0D', marker='o')
        
        plt.grid(True, linestyle='--', linewidth=0.5)
        plt.title(titulo_grafico_combinado)
        plt.xlabel("Tempo (horas)")
        plt.ylabel("-Zreal (Ohm.cm²)")
        if mostrar_legenda:
            plt.legend()
        
        img_bytes = BytesIO()
        plt.savefig(img_bytes, format="png", dpi=300)
        plt.close()
        imagens.append((titulo_grafico_combinado, img_bytes.getvalue()))
    
    for titulo, img in imagens:
        st.image(img, caption=titulo, use_container_width=True)



# Exibir histórico de gráficos gerados
st.subheader("📜 Histórico de gráficos gerados")
historico_arquivos = os.listdir(HISTORICO_DIR)
if historico_arquivos:
    for arq in historico_arquivos:
        with open(os.path.join(HISTORICO_DIR, arq), "rb") as file:
            st.download_button(
                label=f"Baixar {arq}",
                data=file,
                file_name=arq,
                mime="image/png"
            )
else:
    st.write("Nenhum gráfico gerado ainda.")

# Botão para limpar histórico
st.subheader("🗑️ Gerenciamento do Histórico")
if st.button("Limpar Histórico de Gráficos", key="clear_history_button"):
    for arq in os.listdir(HISTORICO_DIR):
        os.remove(os.path.join(HISTORICO_DIR, arq))
    st.rerun()
