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

# Função para exibir gráficos no Streamlit (pré-visualização)
def exibir_grafico_no_streamlit(df, titulo, exibir_rotulos, rotulo_pontos, mostrar_legenda):
    """Exibe um gráfico individualmente no Streamlit."""
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(df["Zreal"], df["Zimag"], marker='o', linestyle='-', label=titulo)

    if exibir_rotulos and not df.empty:
        ax.annotate(rotulo_pontos, (df["Zreal"].iloc[-1], df["Zimag"].iloc[-1]), textcoords="offset points", xytext=(5,5), ha='right')

    ax.set_xlabel("Z Real")
    ax.set_ylabel("Z Imaginário")
    ax.set_title(titulo)

    if mostrar_legenda:
        ax.legend()

    ax.grid(True)
    st.pyplot(fig)

# Função para gerar gráficos e salvar no ZIP
def gerar_grafico_combinado(dados_graficos, nome_arquivo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda):
    """Gera um gráfico combinado de vários arquivos e salva no arquivo ZIP."""
    plt.figure(figsize=(8, 6))

    for df, titulo in dados_graficos:
        plt.plot(df["Zreal"], df["Zimag"], marker='o', linestyle='-', label=titulo)

        if exibir_rotulos:
            plt.annotate(rotulo_pontos, (df["Zreal"].iloc[-1], df["Zimag"].iloc[-1]), textcoords="offset points", xytext=(5,5), ha='right')

    plt.xlabel("Z Real")
    plt.ylabel("Z Imaginário")
    plt.title("Gráfico Combinado")
    
    if mostrar_legenda:
        plt.legend()

    plt.grid(True)

    # Salvar imagem em memória antes de adicionar ao zip
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format="png")
    img_buffer.seek(0)

    zipf.writestr(f"{nome_arquivo}.png", img_buffer.read())
    plt.close()

# Interface Streamlit
st.set_page_config(page_title="Gerador de Gráficos", page_icon="📊")

st.title("Gerador de Gráficos Z real x Z imaginário")
st.write("Faça upload de um ou mais arquivos `.xlsx` e gere gráficos automaticamente.")

# Upload de arquivos
uploaded_files = st.file_uploader("Selecione os arquivos Excel", type=["xlsx"], accept_multiple_files=True)

# Entrada para o fator de área
fator_area = st.number_input("Insira o fator de área para multiplicação dos valores:", min_value=0.0001, value=1.0)

# Nome da pasta de saída
pasta_saida = st.text_input("Nome da pasta de saída", "Graficos_Gerados")

# Escolha de gráficos
gerar_combinado = st.checkbox("Gerar gráfico combinado com todos os arquivos juntos")
gerar_individual = st.checkbox("Gerar gráficos individuais para cada arquivo")

# Configuração de rótulos e legenda
exibir_rotulos = st.toggle("Exibir frequência nos últimos pontos")
mostrar_legenda = st.checkbox("Mostrar legenda no gráfico", value=True)
rotulo_pontos = ""

if exibir_rotulos:
    rotulo_pontos = st.text_input("Digite a frequência para o último ponto de todos os gráficos:")

# Processamento dos arquivos
if uploaded_files and pasta_saida:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        dados_graficos = []

        for uploaded_file in uploaded_files:
            df = carregar_planilha(uploaded_file)
            if df is not None:
                df[["Zreal", "Zimag"]] *= fator_area  # Aplicação do fator de multiplicação
                titulo = f"{uploaded_file.name.replace('.xlsx', '')}_grafico"
                dados_graficos.append((df, titulo))

                if gerar_individual:
                    # Pré-visualizar gráfico no Streamlit antes de salvar no ZIP
                    with st.expander(f"📊 Pré-visualização: {titulo}"):
                        exibir_grafico_no_streamlit(df, titulo, exibir_rotulos, rotulo_pontos, mostrar_legenda)

                    # Gerar e salvar no ZIP
                    gerar_grafico_combinado([(df, titulo)], titulo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda)

        if gerar_combinado and dados_graficos:
            gerar_grafico_combinado(dados_graficos, f"{pasta_saida}_combinado", zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda)

    zip_buffer.seek(0)

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

# Botão para limpar histórico
st.subheader("🗑️ Gerenciamento do Histórico")
if st.button("Limpar Histórico de Gráficos", key="clear_history_button"):
    for arq in os.listdir(HISTORICO_DIR):
        os.remove(os.path.join(HISTORICO_DIR, arq))
    st.rerun()
