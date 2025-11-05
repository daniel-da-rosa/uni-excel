import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Compilador de Excel", layout="wide")

st.title("Compilador de Arquivos Excel")
st.markdown("---")

# ==========================================================
# 🛑 MUDANÇA ESSENCIAL: INJEÇÃO DE CSS PARA TRADUÇÃO DO WIDGET
# ==========================================================

st.markdown("""
<style>
/* 1. MUDAR "Drag and drop files here" */
[data-testid="stFileUploaderDropzone"] > div {
    /* Esconde o texto original, mas mantém o espaço do ícone */
    visibility: hidden;
    height: 0px; 
    padding-top: 50px; /* Ajusta o padding para centralizar o texto novo */
}

/* 2. INSERIR TEXTO TRADUZIDO */
[data-testid="stFileUploaderDropzone"] > div:before {
    visibility: visible;
    display: block;
    content: "Arraste e solte seus arquivos aqui"; /* <--- SEU NOVO TEXTO AQUI */
    height: 0px;
    font-size: 18px; /* Opcional: Ajuste o tamanho da fonte */
    color: #FAFAFA; /* Cor do texto (mudar se o tema não for escuro) */
    position: relative;
    top: -45px; /* Ajuste a posição vertical */
}

/* 3. MUDAR O TEXTO DO LIMITE (Limit 200MB...) */
[data-testid="stFileUploaderFileStatusBar"] [data-testid="stText"] {
    /* Seleciona o elemento que contém o texto de limite */
    visibility: hidden;
    height: 0px; 
}
[data-testid="stFileUploaderFileStatusBar"] [data-testid="stText"]:before {
    visibility: visible;
    display: block;
    content: "Limite 200MB por arquivo • XLSX"; /* <--- SEU NOVO TEXTO DE LIMITE AQUI */
    position: relative;
    top: -5px;
    height: 0px;
    font-size: 14px;
}
</style>
""", unsafe_allow_html=True)
# ==========================================================
# FIM DA INJEÇÃO DE CSS
# ==========================================================


# 1. Widget de Upload de Arquivos
uploaded_files = st.file_uploader(
    "📂 **Faça o Upload dos Arquivos Excel (.xlsx)**", 
    type=["xlsx"],
    accept_multiple_files=True,
    key="file_uploader_custom" # Adicione um key, é sempre bom para widgets
)

# Constante para as colunas
COLUNAS_SELECIONADAS = 'A:E'

# Verifica se há arquivos e adiciona o botão de compilação
if uploaded_files:
    # ... (O restante do seu código de compilação continua o mesmo) ...
    if st.button(f"Compilar Colunas {COLUNAS_SELECIONADAS} da Última Aba"):
        st.info("Iniciando a compilação dos arquivos...")

        # 3. Inicializa o buffer de memória
        output = BytesIO()

        try:
            # Inicia o escritor do Excel
            with pd.ExcelWriter(output, engine='openpyxl') as writer:

                # Itera sobre cada arquivo enviado
                for file in uploaded_files:
                    
                    # CORREÇÃO CRÍTICA 1: Resetar o ponteiro antes de inspecionar
                    file.seek(0)
                    
                    # 1. Obter o nome da última aba
                    with pd.ExcelFile(file, engine='openpyxl') as xls:
                        sheet_names = xls.sheet_names

                    if not sheet_names:
                        st.warning(f"⚠️ Arquivo '{file.name}' ignorado: Não foram encontradas planilhas.")
                        continue 

                    last_sheet_name = sheet_names[-1]

                    # CORREÇÃO CRÍTICA 2: Resetar o ponteiro ANTES de ler os dados
                    file.seek(0)

                    # 2. Ler apenas a última planilha e SOMENTE as colunas A a E
                    df = pd.read_excel(
                        file, 
                        sheet_name=last_sheet_name, 
                        usecols=COLUNAS_SELECIONADAS, 
                        engine='openpyxl'
                    )

                    # O nome da aba de destino no arquivo compilado
                    base_name = file.name.replace(".xlsx", "")
                    sheet_name_output = f"{base_name} ({last_sheet_name})"[:31]

                    # Escreve o DataFrame como uma nova aba
                    df.to_excel(writer, sheet_name=sheet_name_output, index=False)
                    st.success(f"✅ Arquivo '{file.name}' - Colunas {COLUNAS_SELECIONADAS} da aba '{last_sheet_name}' compiladas em '{sheet_name_output}'")
                    

        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo(s): {e}")

        # 4. Move o ponteiro para o início para o download
        output.seek(0)

        st.success("🎉 Compilação concluída! Faça o download abaixo:")

        # 5. BOTÃO DE DOWNLOAD
        st.download_button(
            label="Baixar Arquivo Excel Compilado",
            data=output,
            file_name="Arquivos_Compilados_A_E.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )