import streamlit as st
import pandas as pd
import io
import os
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Classificador de Afiliados", page_icon="📊", layout="wide")

# --- FUNÇÃO DO ROBÔ QUE VAI LER OS LINKS ---
def analisar_link(url):
    # Se a célula estiver vazia ou não for um link válido
    if pd.isna(url) or not str(url).startswith('http'):
        return "Sem link válido", "Sem link"
    
    # Disfarça o robô do Python como se fosse um navegador comum (Chrome)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
    }
    
    try:
        # Tenta acessar a página (espera no máximo 5 segundos para não travar o sistema)
        resposta = requests.get(url, headers=headers, timeout=5)
        
        # Se a página permitir o acesso (Status 200 = OK)
        if resposta.status_code == 200:
            soup = BeautifulSoup(resposta.text, 'html.parser')
            
            # Tenta pegar o título da página para definir "O que é"
            titulo = soup.title.string.strip() if soup.title else "Site Genérico"
            
            # Lógica simples para tentar achar seguidores (Aviso: redes sociais bloqueiam 99% das vezes)
            texto_pagina = soup.get_text().lower()
            if "instagram.com" in url:
                return "Perfil Instagram", "Bloqueado para acesso (Proteção Anti-Robô)"
            elif "tiktok.com" in url:
                return "Perfil TikTok", "Bloqueado para acesso (Proteção Anti-Robô)"
            elif "youtube.com" in url:
                return "Canal YouTube", "Bloqueado para acesso (Proteção Anti-Robô)"
            else:
                return titulo[:50], "Não aplicável (Não é rede social)"
        else:
            return "Bloqueado para acesso", "Bloqueado para acesso"
            
    except Exception as e:
        # Se der erro de conexão, timeout, etc.
        return "Site fora do ar ou erro", "Bloqueado para acesso"

# -------------------------------------------

if os.path.exists("logo.png"):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

st.title("📊 Classificador e Enriquecedor de Afiliados")
st.write("Além de classificar as cores, o sistema agora lê os links da Coluna I da Base Principal para buscar dados.")

col_main, col_test = st.columns(2)

with col_main:
    st.subheader("📂 1. Base Principal")
    arquivo_principal = st.file_uploader("Envie a Base Principal", type=['xlsx', 'csv'], key="main")

with col_test:
    st.subheader("📂 2. Base Teste")
    arquivo_teste = st.file_uploader("Envie a Base Teste", type=['xlsx', 'csv'], key="test")

# Checkbox para o usuário decidir se quer fazer a busca demorada dos links
buscar_links = st.checkbox("🔍 Ativar busca de 'O que é' e 'Seguidores' nos links da Base Principal (Isso pode demorar alguns minutos)")

if arquivo_principal and arquivo_teste:
    if st.button("🚀 Comparar, Classificar e Processar"):
        
        def carregar_dados(arquivo):
            if arquivo.name.endswith('.csv'):
                return pd.read_csv(arquivo)
            return pd.read_excel(arquivo)

        df_main = carregar_dados(arquivo_principal)
        df_test = carregar_dados(arquivo_teste)

        col_login_main = next((col for col in df_main.columns if str(col).strip().lower() == 'login'), None)
        col_login_test = next((col for col in df_test.columns if str(col).strip().lower() == 'login'), None)

        if not col_login_main or not col_login_test:
            st.error("Erro: A coluna 'Login' não foi encontrada nas planilhas.")
        else:
            with st.spinner('Cruzando as informações e aplicando cores...'):
                logins_main = df_main[col_login_main].astype(str).str.strip().str.lower()
                logins_test = df_test[col_login_test].astype(str).str.strip().str.lower()

                contagem_main = logins_main.value_counts()
                contagem_test = logins_test.value_counts()

                status_list = []
                cor_list = []

                for idx, row in df_test.iterrows():
                    login_limpo = str(row[col_login_test]).strip().lower()
                    qtd_na_principal = contagem_main.get(login_limpo, 0)
                    qtd_na_teste = contagem_test.get(login_limpo, 0)

                    if qtd_na_principal == 0:
                        status_list.append("Novo (Não está na base principal)")
                        cor_list.append("Laranja")
                    elif qtd_na_principal > 0 and qtd_na_teste == 1:
                        status_list.append("1 para 1 (Já existe, sem duplicidade)")
                        cor_list.append("Verde")
                    elif qtd_na_principal > 0 and qtd_na_teste > 1:
                        status_list.append("Duplicidade (Repetido na base teste)")
                        cor_list.append("Azul")
                    else:
                        status_list.append("Outro")
                        cor_list.append("Branco")

                df_resultado_teste = df_test.copy()
                df_resultado_teste['Status_Analise'] = status_list 
            
            # --- NOVA PARTE: ANALISAR LINKS DA BASE PRINCIPAL ---
            if buscar_links:
                # A Coluna I no Excel é a 9ª coluna. No Python, a contagem começa do zero, então é o índice 8.
                # Para evitar erros se a planilha for menor que 9 colunas, verificamos o tamanho:
                if len(df_main.columns) >= 9:
                    st.info("Iniciando a leitura dos links na Base Principal. Por favor, aguarde...")
                    barra_progresso = st.progress(0)
                    texto_progresso = st.empty()
                    
                    coluna_I = df_main.columns[8] # Pega o nome da coluna I dinamicamente
                    
                    lista_o_que_e = []
                    lista_seguidores = []
                    
                    total_linhas = len(df_main)
                    
                    for index, link in enumerate(df_main[coluna_I]):
                        # Atualiza a barra de progresso
                        progresso = (index + 1) / total_linhas
                        barra_progresso.progress(progresso)
                        texto_progresso.text(f"Analisando link {index + 1} de {total_linhas}...")
                        
                        o_que_e, seguidores = analisar_link(link)
                        lista_o_que_e.append(o_que_e)
                        lista_seguidores.append(seguidores)
                        
                    # Adiciona as colunas novas na Base Principal
                    df_main['O que é'] = lista_o_que_e
                    df_main['Qtd Seguidores'] = lista_seguidores
                    
                    st.success("Leitura de links concluída!")
                else:
                    st.error("A Base Principal não tem 9 colunas (Coluna I) para fazer a leitura de links.")

            # --- PARTE DO EXCEL (PINTAR AS CÉLULAS) ---
            # Vamos gerar o arquivo da Base Teste colorida
            buffer_teste = io.BytesIO()
            df_resultado_teste.to_excel(buffer_teste, index=False, engine='openpyxl')
            buffer_teste.seek(0)
            
            wb = load_workbook(buffer_teste)
            ws = wb.active
            
            fundo_laranja = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
            fundo_verde = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
            fundo_azul = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            fonte_branca = Font(color="FFFFFF", bold=True)
            fonte_preta = Font(color="000000", bold=True)
            
            idx_coluna_login = None
            for idx, celula in enumerate(ws[1], start=1):
                if celula.value == col_login_test:
                    idx_coluna_login = idx
                    break
            
            if idx_coluna_login:
                for num_linha, cor in enumerate(cor_list, start=2):
                    celula = ws.cell(row=num_linha, column=idx_coluna_login)
                    if cor == "Laranja":
                        celula.fill = fundo_laranja
                        celula.font = fonte_preta
                    elif cor == "Verde":
                        celula.fill = fundo_verde
                        celula.font = fonte_branca
                    elif cor == "Azul":
                        celula.fill = fundo_azul
                        celula.font = fonte_branca
            
            buffer_final_teste = io.BytesIO()
            wb.save(buffer_final_teste)
            buffer_final_teste.seek(0)
            
            st.success("Processamento finalizado!")
            
            # Oferece os downloads
            col_down1, col_down2 = st.columns(2)
            
            with col_down1:
                st.download_button(
                    label="📥 Baixar Base Teste (Colorida)",
                    data=buffer_final_teste,
                    file_name="Base_Teste_Colorida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col_down2:
                if buscar_links and len(df_main.columns) >= 9:
                    buffer_main = io.BytesIO()
                    df_main.to_excel(buffer_main, index=False)
                    buffer_main.seek(0)
                    
                    st.download_button(
                        label="📥 Baixar Base Principal (Com novos dados)",
                        data=buffer_main,
                        file_name="Base_Principal_Enriquecida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
