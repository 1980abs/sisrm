import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Classificador de Afiliados", page_icon="📊", layout="wide")

# Adiciona a logo se existir
if os.path.exists("logo.png"):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

st.title("📊 Classificador de Afiliados")
st.write("Faça o upload das duas planilhas separadamente para classificar a **Base Teste** com as cores solicitadas.")

# Cria duas colunas para o upload ficar organizado lado a lado
col_main, col_test = st.columns(2)

with col_main:
    st.subheader("📂 1. Base Principal")
    st.write("*(Ex: Planilha com 315 afiliados)*")
    arquivo_principal = st.file_uploader("Envie a Base Principal", type=['xlsx', 'csv'], key="main")

with col_test:
    st.subheader("📂 2. Base Teste")
    st.write("*(Ex: Planilha com 55 afiliados)*")
    arquivo_teste = st.file_uploader("Envie a Base Teste", type=['xlsx', 'csv'], key="test")

# Só roda a análise quando os dois arquivos forem enviados
if arquivo_principal and arquivo_teste:
    if st.button("🚀 Comparar e Classificar"):
        with st.spinner('Cruzando as informações...'):
            
            # Função para ler csv ou excel
            def carregar_dados(arquivo):
                if arquivo.name.endswith('.csv'):
                    return pd.read_csv(arquivo)
                return pd.read_excel(arquivo)

            df_main = carregar_dados(arquivo_principal)
            df_test = carregar_dados(arquivo_teste)

            # Acha as colunas 'Login' nas duas planilhas
            col_login_main = next((col for col in df_main.columns if str(col).strip().lower() == 'login'), None)
            col_login_test = next((col for col in df_test.columns if str(col).strip().lower() == 'login'), None)

            if not col_login_main or not col_login_test:
                st.error("Erro: A coluna 'Login' não foi encontrada em uma (ou ambas) as planilhas.")
            else:
                # Limpa espaços e transforma em minúsculo para garantir que a comparação seja perfeita
                logins_main = df_main[col_login_main].astype(str).str.strip().str.lower()
                logins_test = df_test[col_login_test].astype(str).str.strip().str.lower()

                # Conta quantas vezes cada login aparece em cada planilha
                contagem_main = logins_main.value_counts()
                contagem_test = logins_test.value_counts()

                status_list = []
                cor_list = []

                # Percorre cada linha da Base Teste aplicando as regras explicadas
                for idx, row in df_test.iterrows():
                    login_limpo = str(row[col_login_test]).strip().lower()
                    
                    qtd_na_principal = contagem_main.get(login_limpo, 0)
                    qtd_na_teste = contagem_test.get(login_limpo, 0)

                    # REGRA 1 (Ex: Losartana): Não existe na Base Principal -> Laranja
                    if qtd_na_principal == 0:
                        status_list.append("Novo (Não está na base principal)")
                        cor_list.append("Laranja")
                        
                    # REGRA 2 (Ex: Trebor00): Existe na Principal e aparece só 1 vez na Teste -> Verde
                    elif qtd_na_principal > 0 and qtd_na_teste == 1:
                        status_list.append("1 para 1 (Já existe, sem duplicidade)")
                        cor_list.append("Verde")
                        
                    # REGRA 3: Existe na Principal e aparece mais de 1 vez na Teste -> Azul
                    elif qtd_na_principal > 0 and qtd_na_teste > 1:
                        status_list.append("Duplicidade (Repetido na base teste)")
                        cor_list.append("Azul")
                        
                    else:
                        status_list.append("Outro")
                        cor_list.append("Branco")

                # Cria a planilha final baseada na Teste
                df_resultado = df_test.copy()
                df_resultado['Status_Analise'] = status_list # Adiciona coluna de filtro
                
                # Resumo visual na tela
                st.success("Análise concluída com sucesso!")
                
                col_res1, col_res2, col_res3 = st.columns(3)
                col_res1.warning(f"🔸 Novos (Laranja): **{cor_list.count('Laranja')}**")
                col_res2.success(f"🟩 1 para 1 (Verde): **{cor_list.count('Verde')}**")
                col_res3.info(f"🟦 Duplicidades (Azul): **{cor_list.count('Azul')}**")
                
                st.write("Prévia dos Dados:")
                st.dataframe(df_resultado.head(10))
                
                # --- PARTE DO EXCEL (PINTAR AS CÉLULAS) ---
                buffer = io.BytesIO()
                df_resultado.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                
                wb = load_workbook(buffer)
                ws = wb.active
                
                # Definindo as cores do Excel
                fundo_laranja = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")
                fundo_verde = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                fundo_azul = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
                
                fonte_branca = Font(color="FFFFFF", bold=True)
                fonte_preta = Font(color="000000", bold=True)
                
                # Acha qual é a coluna de Login no Excel gerado
                idx_coluna_login = None
                for idx, celula in enumerate(ws[1], start=1):
                    if celula.value == col_login_test:
                        idx_coluna_login = idx
                        break
                
                # Aplica as cores na coluna de Login
                if idx_coluna_login:
                    for num_linha, cor in enumerate(cor_list, start=2):
                        celula = ws.cell(row=num_linha, column=idx_coluna_login)
                        
                        if cor == "Laranja":
                            celula.fill = fundo_laranja
                            celula.font = fonte_preta # Fonte preta dá mais leitura no laranja
                        elif cor == "Verde":
                            celula.fill = fundo_verde
                            celula.font = fonte_branca
                        elif cor == "Azul":
                            celula.fill = fundo_azul
                            celula.font = fonte_branca
                
                # Salva o arquivo final
                buffer_final = io.BytesIO()
                wb.save(buffer_final)
                buffer_final.seek(0)
                
                st.download_button(
                    label="📥 Baixar Planilha Classificada",
                    data=buffer_final,
                    file_name="Base_Teste_Classificada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
