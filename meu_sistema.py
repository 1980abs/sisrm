import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Sistema de Afiliados", page_icon="📊", layout="centered")

if "chave_uploader" not in st.session_state:
    st.session_state.chave_uploader = 0

def limpar_dados():
    st.session_state.chave_uploader += 1

if os.path.exists("logo.png"):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

st.title("📊 Comparador de Planilhas - Afiliados")
st.write("Faça upload das planilhas. Registros repetidos ficam em **Azul** e registros novos em **Laranja**.")

st.button("🔄 Limpar Dados / Iniciar Nova Análise", on_click=limpar_dados)

arquivos_enviados = st.file_uploader(
    "Arraste e solte as planilhas aqui", 
    type=['xlsx', 'csv'], 
    accept_multiple_files=True,
    key=str(st.session_state.chave_uploader)
)

if arquivos_enviados:
    st.success(f"{len(arquivos_enviados)} arquivo(s) carregado(s) com sucesso!")
    
    if st.button("🚀 Analisar e Juntar Planilhas"):
        with st.spinner('Analisando os dados...'):
            lista_planilhas = []
            
            for arquivo in arquivos_enviados:
                if arquivo.name.endswith('.csv'):
                    df = pd.read_csv(arquivo)
                else:
                    df = pd.read_excel(arquivo)
                
                df['Origem_Arquivo'] = arquivo.name 
                lista_planilhas.append(df)
            
            planilha_completa = pd.concat(lista_planilhas, ignore_index=True)
            planilha_completa = planilha_completa.dropna(axis=1, how='all')
            
            coluna_login = None
            for col in planilha_completa.columns:
                if str(col).strip().lower() == 'login':
                    coluna_login = col
                    break
            
            if not coluna_login:
                st.error("Erro: Não encontrei nenhuma coluna chamada 'Login' nas planilhas.")
            else:
                # Limpa espaços das pontas
                planilha_completa[coluna_login] = planilha_completa[coluna_login].astype(str).str.strip()
                
                # Identifica duplicados ignorando maiúsculas/minúsculas
                chave_comparativa = planilha_completa[coluna_login].str.lower()
                planilha_completa['Eh_Duplicado'] = chave_comparativa.duplicated(keep=False)
                
                # Ordena agrupando duplicados no topo
                planilha_completa = planilha_completa.sort_values(
                    by=['Eh_Duplicado', coluna_login], 
                    ascending=[False, True]
                ).reset_index(drop=True)
                
                # Recalcula a máscara após reordenar a tabela
                chave_ordenada = planilha_completa[coluna_login].str.lower()
                duplicados_mascara = chave_ordenada.duplicated(keep=False)
                
                qtd_duplicados = duplicados_mascara.sum()
                qtd_novos = len(planilha_completa) - qtd_duplicados
                
                planilha_completa = planilha_completa.drop(columns=['Eh_Duplicado'])
                
                st.warning(f"Análise concluída! Encontramos **{qtd_duplicados}** registros repetidos (Azul) e **{qtd_novos}** registros novos (Laranja).")
                st.dataframe(planilha_completa.head(15))
                
                # Estilização no Excel
                buffer = io.BytesIO()
                planilha_completa.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                
                wb = load_workbook(buffer)
                ws = wb.active
                
                # Definição das cores
                fundo_azul = PatternFill(start_color="00008B", end_color="00008B", fill_type="solid")
                fundo_laranja = PatternFill(start_color="FF7F00", end_color="FF7F00", fill_type="solid")
                fonte_branca = Font(color="FFFFFF", bold=True)
                
                idx_coluna_login = None
                for idx, celula in enumerate(ws[1], start=1):
                    if celula.value == coluna_login:
                        idx_coluna_login = idx
                        break
                        
                if idx_coluna_login:
                    for num_linha, eh_duplicado in enumerate(duplicados_mascara, start=2):
                        celula = ws.cell(row=num_linha, column=idx_coluna_login)
                        celula.font = fonte_branca
                        if eh_duplicado:
                            celula.fill = fundo_azul
                        else:
                            celula.fill = fundo_laranja
                            
                buffer_final = io.BytesIO()
                wb.save(buffer_final)
                buffer_final.seek(0)
                
                st.success("Tudo pronto! Sua planilha colorida já pode ser baixada.")
                st.download_button(
                    label="📥 Baixar Planilha Final Formatada",
                    data=buffer_final,
                    file_name="Planilha_Comparada_Organizada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
