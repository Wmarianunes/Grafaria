import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import zipfile
from io import BytesIO

# Criar diretório temporário para histórico de gráficos
HISTORICO_DIR = "historico_graficos"
os.makedirs(HISTORICO_DIR, exist_ok=True)

def carregar_planilha(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, skiprows=5, usecols="A:B")
        df.columns = ["Zreal", "Zimag"]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        return None

def gerar_grafico_combinado(dados_graficos, titulo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda, multiplicador, otimizar_espaco):
    try:
        img_bytes = BytesIO()

        # Determinar os limites máximos e mínimos para ajustar a altura dinamicamente
        max_zreal, max_zimag = 0, 0
        min_zimag = float("inf")

        for df, _ in dados_graficos:
            df = df.copy()
            df["Zreal"] *= multiplicador
            df["Zimag"] *= multiplicador
            max_zreal = max(max_zreal, df["Zreal"].max())
            max_zimag = max(max_zimag, df["Zimag"].max())
            min_zimag = min(min_zimag, df["Zimag"].min())

        largura = 8  # Mantemos a largura fixa
        altura = 8   # Altura padrão caso não precise otimização

        if otimizar_espaco:
            # Ajustar a altura dinamicamente com base no intervalo de Zimag
            altura = max(4, (max_zimag - min_zimag) / max_zreal * largura)  # Mantém a proporção

        plt.figure(figsize=(largura, altura))
        marcadores = ['o', '+', '*', '>', 'x', '^', 'v', '<', '|', '_', 's', 'D', 'p', 'h', 'H']

        for index, (df, legenda) in enumerate(dados_graficos):
            marcador = marcadores[index % len(marcadores)]
            plt.scatter(df["Zreal"], -df["Zimag"], marker=marcador, label=legenda)

            if exibir_rotulos and rotulo_pontos:
                ultimo_ponto = df.iloc[-1]
                plt.annotate(rotulo_pontos,
                             (ultimo_ponto["Zreal"], -ultimo_ponto["Zimag"]),
                             fontsize=9, ha='right', color='black')

        plt.xlim(0, max_zreal * 1.1)
        plt.ylim(-max_zimag * 1.1, -min_zimag * 1.1)  # Ajusta o Y para evitar espaços desnecessários
        plt.gca().set_aspect('auto', adjustable='datalim')
        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.xlabel("Z real (ohm.cm^2)")
        plt.ylabel("-Z imag (ohm.cm^2)")
        if mostrar_legenda:
            plt.legend()
        plt.title(titulo)
        plt.savefig(img_bytes, format="png", dpi=300, bbox_inches='tight')
        plt.close()
        zipf.writestr(f"{titulo}.png", img_bytes.getvalue())
        return img_bytes.getvalue()
    except Exception as e:
        st.error(f"Erro ao gerar gráfico combinado: {e}")
        return None

st.set_page_config(page_title="Gerador de Gráficos", page_icon="📊")

st.title("Gerador de Gráficos Z real X Z imaginário")
st.write("Faça upload de um ou mais arquivos `.xlsx` gerados na aba Conversor para criar gráficos automaticamente.")

uploaded_files = st.file_uploader("Selecione os arquivos Excel", type=["xlsx"], accept_multiple_files=True)

pasta_saida = st.text_input("Nome da pasta de saída", "Graficos_Gerados")

multiplicador = st.number_input("Área do corpo de prova:", min_value=0.1, value=1.0, step=0.1)

gerar_combinado = st.checkbox("Gerar gráfico combinando todos os arquivos juntos")
gerar_individual = st.checkbox("Gerar gráficos individuais para cada arquivo")

exibir_rotulos = st.toggle("Exibir valor da frequência nos últimos pontos")
mostrar_legenda = st.checkbox("Mostrar legenda no gráfico", value=True)
rotulo_pontos = ""
if exibir_rotulos:
    rotulo_pontos = st.text_input("Digite a frequência para o último ponto:")

duas_colunas = st.checkbox("Exibir gráficos em duas colunas")
otimizar_espaco = st.checkbox("Otimizar espaço do gráfico")

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
                    img = gerar_grafico_combinado([(df, titulo)], titulo, zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda, multiplicador, otimizar_espaco)
                    if img:
                        imagens.append((titulo, img))

        if gerar_combinado and dados_graficos:
            img = gerar_grafico_combinado(dados_graficos, f"{pasta_saida}_combinado", zipf, exibir_rotulos, rotulo_pontos, mostrar_legenda, multiplicador, otimizar_espaco)
            if img:
                imagens.append((f"{pasta_saida}_combinado", img))

    zip_buffer.seek(0)
    st.download_button(
        label="Baixar Pasta com os Gráficos",
        data=zip_buffer,
        file_name=f"{pasta_saida}.zip",
        mime="application/zip"
    )
    st.success("Gráficos gerados! Baixe a pasta compactada acima.")

    if imagens:
        if duas_colunas:
            col1, col2 = st.columns(2)
            for i, (titulo, img) in enumerate(imagens):
                if i % 2 == 0:
                    col1.image(img, caption=titulo, use_container_width=True)
                else:
                    col2.image(img, caption=titulo, use_container_width=True)
        else:
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

