import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import zipfile
from io import BytesIO

# Criar diretório temporário para histórico de gráficos
HISTORICO_DIR = "historico_graficos"
os.makedirs(HISTORICO_DIR, exist_ok=True)

# Função para carregar a planilha
def carregar_planilha(uploaded_file):
    """Carrega a planilha Excel a partir do upload do usuário."""
    try:
        df = pd.read_excel(uploaded_file, skiprows=5, usecols="A:B")  # Pegando colunas A e B, começando na linha 6
        df.columns = ["Zreal", "Zimag"]  # Renomeia colunas
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        return None

# Criar nome seguro para arquivos
def criar_nome_seguro(titulo):
    """Gera um nome seguro para arquivos."""
    return "".join([c if c.isalnum() or c in (' ', '.', '_') else '_' for c in titulo])

# Gerar gráfico combinado
def gerar_grafico_combinado(dados_graficos, titulo, zipf, exibir_rotulos, frequencias, mostrar_preview):
    """Gera e salva um gráfico combinado no ZIP, com opção de rótulos nos últimos pontos."""
    try:
        img_bytes = BytesIO()
        plt.figure(figsize=(8, 8))
        max_val = 0
        marcadores = ['o', '+', '*', '>', 'x', '^', 'v', '<', '|', '_', 's', 'D', 'p', 'h', 'H']

        for index, (df, legenda) in enumerate(dados_graficos):
            marcador = marcadores[index % len(marcadores)]
            plt.scatter(df["Zreal"], -df["Zimag"], marker=marcador, label=legenda)
            max_val = max(max_val, df["Zreal"].max(), df["Zimag"].max())

            # Adicionar rótulos apenas nos últimos pontos se ativado
            if exibir_rotulos:
                ultimo_ponto = df.iloc[-1]  # Última linha da tabela
                freq_texto = frequencias.get(legenda, "Sem frequência informada")  # Busca a frequência inserida pelo usuário
                plt.annotate(freq_texto, 
                             (ultimo_ponto["Zreal"], -ultimo_ponto["Zimag"]),
                             fontsize=9, ha='right', color='red')

        plt.xlim(0, max_val * 1.1)
        plt.ylim(0, max_val * 1.1)
        plt.gca().set_aspect('equal', adjustable='box')
        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.xlabel("Z real (ohm.cm^2)")
        plt.ylabel("-Z imag (ohm.cm^2)")
        plt.legend()

        plt.title(titulo)
        plt.savefig(img_bytes, format="png", dpi=300)
        plt.close()

        # Adicionar ao ZIP
        zipf.writestr(f"{titulo}.png", img_bytes.getvalue())

        # Salvar no histórico
        historico_path = os.path.join(HISTORICO_DIR, f"{titulo}.png")
        with open(historico_path, "wb") as f:
            f.write(img_bytes.getvalue())

        # Exibir pré-visualização no Streamlit, se ativado
        if mostrar_preview:
            st.image(img_bytes.getvalue(), caption=titulo, use_column_width=True)

    except Exception as e:
        st.error(f"Erro ao gerar gráfico combinado: {e}")

# Interface Streamlit
st.set_page_config(page_title="Gerador de Gráficos", page_icon="📊")

st.title("Gerador de Gráficos a partir de Arquivos Excel")
st.write("Faça upload de um ou mais arquivos `.xlsx` para gerar gráficos automaticamente.")

# Upload de arquivos
uploaded_files = st.file_uploader("Selecione os arquivos Excel", type=["xlsx"], accept_multiple_files=True)

# Nome da pasta de saída
pasta_saida = st.text_input("Nome da pasta de saída", "Graficos_Gerados")

# Escolha de gráficos
gerar_combinado = st.checkbox("Gerar gráfico combinado (todos os arquivos juntos)")

# Opção para exibir rótulos nos últimos pontos
exibir_rotulos = st.checkbox("Exibir rótulos nos últimos pontos com a frequência inserida")

# Dicionário para armazenar frequências inseridas pelo usuário
frequencias = {}

# Solicita que o usuário digite a frequência para cada arquivo
if uploaded_files and exibir_rotulos:
    for uploaded_file in uploaded_files:
        frequencias[uploaded_file.name] = st.text_input(f"Digite a frequência para {uploaded_file.name}:", "")

# Opção para mostrar pré-visualização do gráfico combinado
mostrar_preview = st.checkbox("Mostrar pré-visualização do gráfico combinado antes de baixar")

# Processamento dos arquivos
if uploaded_files and pasta_saida and gerar_combinado:
    arquivos_processados = []

    # Criar arquivo ZIP temporário para salvar os gráficos
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        dados_graficos = []

        for uploaded_file in uploaded_files:
            df = carregar_planilha(uploaded_file)
            if df is not None:
                titulo = f"{uploaded_file.name.replace('.xlsx', '')}_grafico_combinado"
                dados_graficos.append((df, titulo))

        if gerar_combinado and dados_graficos:
            gerar_grafico_combinado(dados_graficos, f"{pasta_saida}_combinado", zipf, exibir_rotulos, frequencias, mostrar_preview)

    # Finalizar ZIP
    zip_buffer.seek(0)

    # Criar botão para baixar a pasta ZIP
    st.download_button(
        label="Baixar Pasta com os Gráficos",
        data=zip_buffer,
        file_name=f"{pasta_saida}.zip",
        mime="application/zip"
    )

    st.success("Gráficos gerados! Baixe a pasta compactada acima.")

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
