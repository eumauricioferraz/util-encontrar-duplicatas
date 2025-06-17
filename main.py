import streamlit as st
import pandas as pd
from fuzzywuzzy import process, fuzz
import unicodedata
import re

# --- Funções de Apoio (invariáveis) ---

def preprocess_name(name: str) -> str:
    """
    Normaliza uma string: remove acentos, converte para minúsculas,
    remove caracteres especiais e espaços extras.
    """
    if not isinstance(name, str):
        return ""
    nfkd_form = unicodedata.normalize('NFKD', name)
    name = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    name = name.lower()
    name = re.sub(r'[^a-z0-9\s]', '', name)
    return name.strip()

def find_matches(df1: pd.DataFrame, df2: pd.DataFrame, col1: str, col2: str, threshold: int) -> pd.DataFrame:
    """
    Encontra correspondências aproximadas entre duas colunas de DataFrames.
    """
    source_list = df1[col1].dropna().astype(str).tolist()
    target_list = df2[col2].dropna().astype(str).tolist()
    results = []
    progress_bar = st.progress(0, text="Analisando...")
    total_items = len(source_list)

    for i, query_name in enumerate(source_list):
        best_match = process.extractOne(
            query_name,
            target_list,
            processor=preprocess_name,
            scorer=fuzz.token_set_ratio
        )
        match_name, score = (best_match[0], best_match[1]) if best_match else (None, 0)

        if score >= threshold:
            results.append({
                'Valor de Entrada (Planilha 1)': query_name,
                'Melhor Correspondência (Planilha 2)': match_name,
                'Pontuação de Similaridade (%)': score
            })
        progress_bar.progress((i + 1) / total_items, text=f"Analisando item {i+1} de {total_items}")

    return pd.DataFrame(results)

# --- Interface do Streamlit ---

st.set_page_config(layout="wide")
st.title("🔎 Ferramenta para Comparação de Dados")
st.markdown("Compare dados entre **planilhas específicas** de dois arquivos Excel.")

# --- Configurações na Barra Lateral ---
with st.sidebar:
    st.header("⚙️ Configurações")

    uploaded_file1 = st.file_uploader("1. Carregue o arquivo de Origem", type=["xlsx", "xls"])
    uploaded_file2 = st.file_uploader("2. Carregue o arquivo Alvo", type=["xlsx", "xls"])

    df1, df2 = None, None
    col1_selection, col2_selection = None, None
    
    if uploaded_file1:
        try:
            xls1 = pd.ExcelFile(uploaded_file1)
            sheet_names1 = xls1.sheet_names
            # Adicionada uma 'key' única para o seletor da planilha 1
            selected_sheet1 = st.selectbox(
                f"Selecione a planilha do arquivo '{uploaded_file1.name}'",
                sheet_names1,
                key='sheet_selector_1' 
            )
            df1 = pd.read_excel(uploaded_file1, sheet_name=selected_sheet1)
            # Adicionada uma 'key' única para o seletor de coluna 1
            col1_selection = st.selectbox(
                "Selecione a coluna da Planilha 1",
                df1.columns,
                key='column_selector_1'
            )
        except Exception as e:
            st.error(f"Erro ao ler o arquivo 1: {e}")

    if uploaded_file2:
        try:
            xls2 = pd.ExcelFile(uploaded_file2)
            sheet_names2 = xls2.sheet_names
            # Adicionada uma 'key' única para o seletor da planilha 2
            selected_sheet2 = st.selectbox(
                f"Selecione a planilha do arquivo '{uploaded_file2.name}'",
                sheet_names2,
                key='sheet_selector_2'
            )
            df2 = pd.read_excel(uploaded_file2, sheet_name=selected_sheet2)
            # Adicionada uma 'key' única para o seletor de coluna 2
            col2_selection = st.selectbox(
                "Selecione a coluna da Planilha 2",
                df2.columns,
                key='column_selector_2'
            )
        except Exception as e:
            st.error(f"Erro ao ler o arquivo 2: {e}")
    
    if df1 is not None and df2 is not None:
        similarity_threshold = st.slider(
            "Limiar de Similaridade (%)",
            min_value=0, max_value=100, value=80,
            help="Define o quão rigorosa a correspondência deve ser. 100% significa idêntico."
        )
        run_button = st.button("Encontrar Correspondências")

# --- Área Principal (lógica de execução inalterada) ---
if 'run_button' in locals() and run_button:
    if all([df1 is not None, df2 is not None, col1_selection, col2_selection]):
        with st.spinner("Analisando dados... Por favor, aguarde."):
            results_df = find_matches(df1, df2, col1_selection, col2_selection, similarity_threshold)

        st.success(f"Análise concluída! Encontradas {len(results_df)} correspondências.")

        if not results_df.empty:
            st.dataframe(results_df)
            csv = results_df.to_csv(index=False).encode('utf-8')
            st.download_button(
               label="Baixar resultados como CSV",
               data=csv,
               file_name='resultados_comparacao.csv',
               mime='text/csv',
            )
        else:
            st.warning("Nenhuma correspondência encontrada com o limiar de similaridade definido.")
    else:
        st.error("Por favor, carregue ambos os arquivos e selecione as planilhas e colunas na barra lateral.")
else:
    st.info("Aguardando o carregamento dos arquivos e a configuração na barra lateral para iniciar a análise.")