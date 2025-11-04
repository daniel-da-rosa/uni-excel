import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Compilador de Excel", layout="wide")

st.title("Compilador de Arquivos Excel")
st.markdown("---")

# 1. Widget de Upload de Arquivos
uploaded_files = st.file_uploader(
    "Escolha seus arquivos Excel (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True, 
    help="Você pode selecionar vários arquivos de uma vez."
)

# Verifica se há arquivos e adiciona o botão de compilação
if uploaded_files:
    # 2. BOTÃO DE COMPILAÇÃO
    if st.button("Compilar em um Único Arquivo Excel"):
        st.info("Iniciando a compilação dos arquivos...")
        
        # 3. CORREÇÃO ESSENCIAL: Inicializa o buffer de memória
        output = BytesIO()
        
        try:
            # Inicia o escritor do Excel
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                # Itera sobre cada arquivo enviado
                for file in uploaded_files:
                    # O nome da aba será o nome do arquivo, sem a extensão .xlsx
                    # Limitado a 31 caracteres, que é o máximo do Excel
                    sheet_name = file.name.replace(".xlsx", "")[:31] 
                    
                    # Lê o arquivo Excel na memória
                    df = pd.read_excel(file)
                    
                    # Escreve o DataFrame como uma nova aba
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    st.success(f"✅ Arquivo '{file.name}' compilado na aba '{sheet_name}'")
                    
        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo(s): {e}")
            
        # 4. CORREÇÃO ESSENCIAL: Move o ponteiro para o início para o download
        output.seek(0)
        
        st.success("🎉 Compilação concluída! Faça o download abaixo:")
        
        # 5. BOTÃO DE DOWNLOAD
        st.download_button(
            label="Baixar Arquivo Excel Compilado",
            data=output,
            file_name="Arquivos_Compilados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )