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
def gerar_grafico_combinado(dados_graficos, titulo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda):
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

            # Adicionar rótulo apenas no último ponto de cada conjunto de dados
            if exibir_rotulos and rotulo_pontos:
                ultimo_ponto = df.iloc[-1]  # Última linha da tabela
                plt.annotate(rotulo_pontos, 
                             (ultimo_ponto["Zreal"], -ultimo_ponto["Zimag"]),
                             fontsize=9, ha='right', color='black')

        plt.xlim(0, max_val * 1.1)
        plt.ylim(0, max_val * 1.1)
        plt.gca().set_aspect('equal', adjustable='box')
        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.xlabel("Z real (ohm.cm^2)")
        plt.ylabel("-Z imag (ohm.cm^2)")
        if mostrar_legenda:
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

        return img_bytes.getvalue()

    except Exception as e:
        st.error(f"Erro ao gerar gráfico combinado: {e}")
        return None

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
gerar_individual = st.checkbox("Gerar gráficos individuais para cada arquivo")

# Configuração do toggle e entrada de texto
exibir_rotulos = st.toggle("Exibir rótulos nos últimos pontos")
mostrar_legenda = st.checkbox("Mostrar legenda no gráfico", value=True)
rotulo_pontos = ""

if exibir_rotulos:
    rotulo_pontos = st.text_input("Digite o rótulo para o último ponto de todos os gráficos:")

# Checkbox para visualizar em duas colunas
visualizar_duas_colunas = st.checkbox("Visualizar gráficos em duas colunas")

# Processamento dos arquivos
if uploaded_files and pasta_saida:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        dados_graficos = []
        imagens = []

        for uploaded_file in uploaded_files:
            df = carregar_planilha(uploaded_file)
            if df is not None:
                titulo = f"{uploaded_file.name.replace('.xlsx', '')}_grafico"
                dados_graficos.append((df, titulo))

                if gerar_individual:
                    img = gerar_grafico_combinado([(df, titulo)], titulo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda)
                    if img:
                        imagens.append((titulo, img))

        if gerar_combinado and dados_graficos:
            img = gerar_grafico_combinado(dados_graficos, f"{pasta_saida}_combinado", zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda)
            if img:
                imagens.append((f"{pasta_saida}_combinado", img))

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

    # Exibir gráficos de acordo com a escolha do usuário
    if imagens:
        if visualizar_duas_colunas:
            col1, col2 = st.columns(2)
            for i, (titulo, img) in enumerate(imagens):
                if i % 2 == 0:
                    col1.image(img, caption=titulo, use_container_width=True)
                else:
                    col2.image(img, caption=titulo, use_container_width=True)
        else:
            for titulo, img in imagens:
                st.image(img, caption=titulo, use_container_width=True)

