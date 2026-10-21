import io
import os
import re
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Sistema de Afiliados", page_icon="📊", layout="wide"
)

if os.path.exists("logo.png"):
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.image("logo.png", use_container_width=True)

st.title("📊 Gestor e Analista de Afiliados")

# --- CRIANDO AS ABAS INDEPENDENTES ---
tab_comparador, tab_enriquecedor, tab_gigante = st.tabs([
    "🔀 1. Comparar Duplicados",
    "🔍 2. Enriquecer Links (Web Scraping)",
    "📑 3. Relatório Avançado (Base Gigante/Rápido)",
])

# ==============================================================================
# FUNÇÕES GERAIS
# ==============================================================================
def identificar_tipo_trafego(url_str):
  url = str(url_str).lower()
  if any(x in url for x in ['t.me', 'telegram']): return 'Telegram'
  if any(x in url for x in ['instagram.com', 'instagr.am']): return 'Instagram'
  if any(x in url for x in ['youtube.com', 'youtu.be']): return 'YouTube'
  if any(x in url for x in ['wa.me', 'whatsapp']): return 'WhatsApp'
  if 'tiktok.com' in url: return 'TikTok'
  if any(x in url for x in ['facebook.com', 'fb.com']): return 'Facebook'
  if 'kwai.com' in url: return 'Kwai'
  if any(x in url for x in ['shopee', 'amazon', 'aliexpress', 'magalu']): return 'E-Commerce'
  if 'http' in url: return 'Site Próprio / Landing Page'
  return 'Não Identificado'

def identificar_segmentacao(texto_busca):
  texto = str(texto_busca).lower()
  if re.search(r'bet|cassino|aviator|slots|tigrinho|fortune|stake|blaze|poker|aposta', texto): return 'APOSTA'
  if re.search(r'invest|trader|acoes|banco|renda|orcamento|finance|crypto|bitcoin|forex', texto): return 'FINANÇAS'
  if re.search(r'curso|mentoria|aula|ebook|edtech|concurso|faculdade|treinamento|escola', texto): return 'EDUCAÇÃO'
  if re.search(r'emagrecer|treino|diet|saude|whey|keto|fitness|nutri|estetica|corpo', texto): return 'SAÚDE / FITNESS'
  if re.search(r'skin|make|unha|cabelo|maquiagem|cilios|sobrancelha|skincare', texto): return 'BELEZA / ESTÉTICA'
  if re.search(r'cupom|promo|desconto|oferta', texto): return 'CUPOM / OFERTAS'
  if re.search(r'achado|achadinho', texto): return 'ACHADINHOS'
  if re.search(r'shopee|amazon|magalu|aliexpress', texto): return 'E-COMMERCE'
  return 'OUTROS / NÃO IDENTIFICADO'

def analisar_link_completo(url):
  if pd.isna(url) or not str(url).startswith("http"):
    return "Não Identificado", "Sem Link Válido", "Não Identificado", "Sem Link"
  url_str = str(url).strip()
  tipo_trafego = identificar_tipo_trafego(url_str)
  headers = {"User-Agent": "Mozilla/5.0"}
  try:
    resposta = requests.get(url_str, headers=headers, timeout=5)
    if resposta.status_code == 200:
      soup = BeautifulSoup(resposta.text, "html.parser")
      titulo = soup.title.string.strip() if soup.title else "Página Web"
      meta_desc = ""
      desc_tag = soup.find("meta", attrs={"name": "description"})
      if desc_tag and desc_tag.get("content"): meta_desc = desc_tag["content"]
      texto_completo = f"{url_str} {titulo} {meta_desc}"
      segmentacao = identificar_segmentacao(texto_completo)
      if tipo_trafego in ["Instagram", "TikTok", "YouTube", "Facebook", "Telegram"]:
        return tipo_trafego, "Bloqueado para acesso", segmentacao, f"Perfil/Canal de {tipo_trafego}"
      else:
        return tipo_trafego, "Não aplicável (Site)", segmentacao, titulo[:60]
    else:
      return tipo_trafego, "Bloqueado para acesso", identificar_segmentacao(url_str), "Bloqueado para acesso"
  except Exception:
    return tipo_trafego, "Bloqueado para acesso", identificar_segmentacao(url_str), "Site Indisponível / Erro"

def analisar_link_rapido(url):
  url_str = str(url).strip()
  tipo_trafego = identificar_tipo_trafego(url_str)
  segmentacao = identificar_segmentacao(url_str)
  
  if tipo_trafego in ["Site Próprio / Landing Page", "Não Identificado", "E-Commerce"]:
    oquee = f"Página Web focada em {segmentacao.title()}"
    seguidores = "NÃO APLICÁVEL (SITE)"
  else:
    oquee = f"Canal de {tipo_trafego.title()} focado em {segmentacao.title()}"
    seguidores = "PÁGINA PRIVADA / REQUER API"
    
  return tipo_trafego, seguidores, segmentacao, oquee

# ==============================================================================
# ABA 1 e 2 (Mantidas das versões anteriores)
# ==============================================================================
with tab_comparador:
  st.write("Compare a **Base Teste** contra a **Base Principal** para identificar duplicados.")
  col_main, col_test = st.columns(2)
  with col_main: arq_principal = st.file_uploader("1. Base Principal", type=["xlsx", "csv"], key="c_main")
  with col_test: arq_teste = st.file_uploader("2. Base Teste", type=["xlsx", "csv"], key="c_test")
  
  if arq_principal and arq_teste and st.button("🚀 Comparar"):
      st.info("Comparador em manutenção nesta visualização simplificada.")

with tab_enriquecedor:
  st.write("Envie a **Base Principal** para varrer os links (Lento/Scraping).")
  arq_unico = st.file_uploader("Base Principal", type=["xlsx", "csv"], key="e_main")

# ==============================================================================
# ABA 3: RELATÓRIO AVANÇADO (BASE GIGANTE - MÚLTIPLAS ABAS)
# ==============================================================================
with tab_gigante:
  st.header("Processador de Bases Gigantes")
  st.write("1. Une todas as abas, analisa os links e cria as 4 colunas (Segmento, Categoria, Descrição, Seguidores).")
  st.write("2. Gera **DOIS** relatórios formatados: um com links classificados e outro com os sem link (para contato).")

  arq_gigante = st.file_uploader("Envie a Base Completa (.xlsx)", type=["xlsx"], key="u_gigante")

  if arq_gigante:
    if st.button("🚀 Processar Base e Gerar Relatórios"):
      with st.spinner("A ler abas e a classificar... isto pode demorar 1-2 minutos."):
        
        todas_as_abas = pd.read_excel(arq_gigante, sheet_name=None)
        lista_df = []
        for nome_aba, df_aba in todas_as_abas.items():
          df_aba["Mes_Origem (Aba)"] = nome_aba
          lista_df.append(df_aba)

        df_completo = pd.concat(lista_df, ignore_index=True)
        df_completo = df_completo.dropna(how='all')
        
        col_fonte = next((c for c in df_completo.columns if 'fonte' in str(c).lower()), df_completo.columns[4])
        df_completo[col_fonte] = df_completo[col_fonte].astype(str).str.strip()

        termos_excluir = ["erro", "sem link", "não cadastrada", "nao cadastrada", "-", "nan", "", "nenhuma", "null"]
        mascara_lixo = df_completo[col_fonte].str.lower().isin(termos_excluir)
        mascara_sem_http = ~df_completo[col_fonte].str.contains("http|www|\.com|\.br|@", case=False, na=False)

        # Base 2: Sem Link / Erros
        df_sem_link = df_completo[mascara_lixo | mascara_sem_http].copy()
        
        # Base 1: Com Links Válidos
        df_limpo = df_completo[~(mascara_lixo | mascara_sem_http)].copy()

        st.info(f"Análise: {len(df_limpo)} links válidos e {len(df_sem_link)} registos sem link para contactar.")

        list_trafego, list_seguidores, list_segmentacao, list_oquee = [], [], [], []

        for link in df_limpo[col_fonte]:
          trafego, seg, segm, oquee = analisar_link_rapido(link)
          list_trafego.append(trafego)
          list_seguidores.append(seg)
          list_segmentacao.append(segm)
          list_oquee.append(oquee)

        df_limpo["Segmento"] = list_segmentacao
        df_limpo["Categoria"] = list_trafego
        df_limpo["Breve Descrição"] = list_oquee
        df_limpo["Qtd Seguidores"] = list_seguidores

        df_final = df_limpo.sort_values(by=["Segmento", "Categoria"], ascending=True)

        st.success("Tudo pronto! Pode descarregar os seus dois relatórios:")

        # --- EXPORTAR BASE VÁLIDA ---
        buf_final = io.BytesIO()
        df_final.to_excel(buf_final, index=False, engine='openpyxl')
        buf_final.seek(0)
        
        wb = load_workbook(buf_final)
        ws = wb.active
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        ws.auto_filter.ref = ws.dimensions
        
        buf_formatado = io.BytesIO()
        wb.save(buf_formatado)
        buf_formatado.seek(0)

        st.download_button(
            label="✅ 1. Baixar Relatório com Links Classificados (9 Colunas)",
            data=buf_formatado, file_name="1_Relatorio_Links_Classificados.xlsx"
        )

        # --- EXPORTAR BASE SEM LINKS ---
        buf_erros = io.BytesIO()
        df_sem_link.to_excel(buf_erros, index=False, engine='openpyxl')
        buf_erros.seek(0)
        
        wb2 = load_workbook(buf_erros)
        ws2 = wb2.active
        header_fill_erro = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        for cell in ws2[1]:
            cell.fill = header_fill_erro
            cell.font = header_font
        ws2.auto_filter.ref = ws2.dimensions
        
        buf_erros_formatado = io.BytesIO()
        wb2.save(buf_erros_formatado)
        buf_erros_formatado.seek(0)

        st.download_button(
            label="⚠️ 2. Baixar Relatório de Contatos Pendentes (Sem Link)",
            data=buf_erros_formatado, file_name="2_Relatorio_Para_Contatar_Sem_Link.xlsx"
        )
