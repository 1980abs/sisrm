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
    "🔍 2. Enriquecer Links",
    "📑 3. Relatório Avançado (Base Gigante)",
])


# ==============================================================================
# FUNÇÕES GERAIS DE IDENTIFICAÇÃO E ANÁLISE
# ==============================================================================
def identificar_tipo_trafego(url_str):
  url = str(url_str).lower()
  if any(x in url for x in ['t.me', 'telegram']): return "Telegram"
  elif any(x in url for x in ['instagram.com', 'instagr.am']): return "Instagram"
  elif any(x in url for x in ['youtube.com', 'youtu.be']): return "YouTube"
  elif any(x in url for x in ['wa.me', 'whatsapp']): return "WhatsApp"
  elif 'tiktok.com' in url: return "TikTok"
  elif any(x in url for x in ['facebook.com', 'fb.com']): return "Facebook"
  elif 'kwai.com' in url: return "Kwai"
  elif any(x in url for x in ['shopee', 'amazon', 'aliexpress', 'magalu']): return "E-Commerce"
  elif 'http' in url: return "Site Próprio / Landing Page"
  else: return "Não Identificado"


def identificar_segmentacao(texto_busca):
  texto = str(texto_busca).lower()
  if re.search(r'bet|cassino|aviator|slots|tigrinho|fortune|stake|blaze|poker|aposta', texto): return "APOSTA"
  elif re.search(r'invest|trader|acoes|banco|renda|orcamento|finance|crypto|bitcoin|forex', texto): return "FINANÇAS"
  elif re.search(r'curso|mentoria|aula|ebook|edtech|concurso|faculdade|treinamento|escola', texto): return "EDUCAÇÃO"
  elif re.search(r'emagrecer|treino|diet|saude|whey|keto|fitness|nutri|estetica|corpo', texto): return "SAÚDE / FITNESS"
  elif re.search(r'skin|make|unha|cabelo|maquiagem|cilios|sobrancelha|skincare', texto): return "BELEZA / ESTÉTICA"
  elif re.search(r'cupom|promo|desconto|oferta', texto): return "CUPOM / OFERTAS"
  elif re.search(r'achado|achadinho', texto): return "ACHADINHOS"
  elif re.search(r'shopee|amazon|magalu|aliexpress', texto): return "E-COMMERCE"
  else: return "OUTROS / NÃO IDENTIFICADO"


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
      if desc_tag and desc_tag.get("content"):
        meta_desc = desc_tag["content"]

      texto_completo = f"{url_str} {titulo} {meta_desc}"
      segmentacao = identificar_segmentacao(texto_completo)

      if tipo_trafego in ["Instagram", "TikTok", "YouTube", "Facebook", "Telegram"]:
        qtd_seguidores = "Bloqueado para acesso"
        o_que_e = f"Perfil/Canal de {tipo_trafego}"
      else:
        qtd_seguidores = "Não aplicável (Site Próprio)"
        o_que_e = titulo[:60]

      return tipo_trafego, qtd_seguidores, segmentacao, o_que_e
    else:
      segmentacao = identificar_segmentacao(url_str)
      return tipo_trafego, "Bloqueado para acesso", segmentacao, "Bloqueado para acesso"
  except Exception:
    segmentacao = identificar_segmentacao(url_str)
    return tipo_trafego, "Bloqueado para acesso", segmentacao, "Site Indisponível / Erro"


def analisar_link_rapido(url):
  if pd.isna(url) or not str(url).startswith("http"):
    return "Não Identificado", "Sem Link Válido", "Não Identificado", "Sem Link"

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
# ABA 1: COMPARADOR DE DUPLICADOS (Baseadas em 2 arquivos)
# ==============================================================================
with tab_comparador:
  st.write("Compare a **Base Teste** contra a **Base Principal** para identificar duplicados.")

  col_main, col_test = st.columns(2)
  with col_main:
    arq_principal = st.file_uploader("1. Envie a Base Principal", type=["xlsx", "csv"], key="comp_main")
  with col_test:
    arq_teste = st.file_uploader("2. Envie a Base Teste", type=["xlsx", "csv"], key="comp_test")

  if arq_principal and arq_teste:
    if st.button("🚀 Comparar Duplicidades", key="btn_comparar"):

      def ler_arquivo(arq):
        return pd.read_csv(arq) if arq.name.endswith(".csv") else pd.read_excel(arq)

      df_m = ler_arquivo(arq_principal)
      df_t = ler_arquivo(arq_teste)

      col_login_m = next((c for c in df_m.columns if str(c).strip().lower() == "login"), None)
      col_login_t = next((c for c in df_t.columns if str(c).strip().lower() == "login"), None)

      if not col_login_m or not col_login_t:
        st.error("Erro: A coluna 'Login' precisa existir em ambas as planilhas.")
      else:
        logins_m = df_m[col_login_m].astype(str).str.strip().str.lower()
        logins_t = df_t[col_login_t].astype(str).str.strip().str.lower()

        cnt_m = logins_m.value_counts()
        cnt_t = logins_t.value_counts()

        status_list, cor_list = [], []

        for _, row in df_t.iterrows():
          login_limpo = str(row[col_login_t]).strip().lower()
          qtd_p = cnt_m.get(login_limpo, 0)
          qtd_t = cnt_t.get(login_limpo, 0)

          if qtd_p == 0:
            status_list.append("Novo (Não existe na principal)")
            cor_list.append("Laranja")
          elif qtd_p > 0 and qtd_t == 1:
            status_list.append("1 para 1 (Existe, sem duplicidade na teste)")
            cor_list.append("Verde")
          elif qtd_p > 0 and qtd_t > 1:
            status_list.append("Duplicidade (Repetido na base teste)")
            cor_list.append("Azul")

        df_res = df_t.copy()
        df_res["Status_Analise"] = status_list

        buffer = io.BytesIO()
        df_res.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        wb = load_workbook(buffer)
        ws = wb.active

        f_laranja = PatternFill(start_color="FF9900", fill_type="solid")
        f_verde = PatternFill(start_color="00B050", fill_type="solid")
        f_azul = PatternFill(start_color="0070C0", fill_type="solid")
        font_w = Font(color="FFFFFF", bold=True)
        font_b = Font(color="000000", bold=True)

        idx_col = None
        for idx, cel in enumerate(ws[1], start=1):
          if cel.value == col_login_t:
            idx_col = idx
            break

        if idx_col:
          for n_linha, cor in enumerate(cor_list, start=2):
            c = ws.cell(row=n_linha, column=idx_col)
            if cor == "Laranja":
              c.fill, c.font = f_laranja, font_b
            elif cor == "Verde":
              c.fill, c.font = f_verde, font_w
            elif cor == "Azul":
              c.fill, c.font = f_azul, font_w

        buf_final = io.BytesIO()
        wb.save(buf_final)
        buf_final.seek(0)

        st.success("Comparação concluída!")
        st.download_button(
            "📥 Baixar Base Teste Classificada",
            data=buf_final,
            file_name="Base_Teste_Classificada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ==============================================================================
# ABA 2: ENRIQUECEDOR DE LINKS (Acesso Direto à Web)
# ==============================================================================
with tab_enriquecedor:
  st.write("Envie a **Base Principal** para varrer os links e preencher automaticamente as novas colunas.")

  arq_unico = st.file_uploader("Envie a Base Principal", type=["xlsx", "csv"], key="enri_main")

  if arq_unico:
    df_enri = pd.read_csv(arq_unico) if arq_unico.name.endswith(".csv") else pd.read_excel(arq_unico)

    st.subheader("Configuração das Colunas")
    colunas_disponiveis = list(df_enri.columns)
    indice_padrao = 8 if len(colunas_disponiveis) >= 9 else 0
    coluna_link_selecionada = st.selectbox(
        "Selecione em qual coluna estão os links/redes sociais:",
        options=colunas_disponiveis,
        index=indice_padrao,
    )

    if st.button("🔍 Iniciar Enriquecimento dos Links"):
      st.info("Varrendo os links. Aguarde o término do processamento...")

      barra = st.progress(0)
      status_txt = st.empty()

      list_trafego, list_seguidores, list_segmentacao, list_o_que_e = [], [], [], []
      total = len(df_enri)

      for idx, link in enumerate(df_enri[coluna_link_selecionada]):
        barra.progress((idx + 1) / total)
        status_txt.text(f"Processando link {idx + 1} de {total}...")

        trafego, seg, segm, oquee = analisar_link_completo(link)
        list_trafego.append(trafego)
        list_seguidores.append(seg)
        list_segmentacao.append(segm)
        list_o_que_e.append(oquee)

      df_enri["Tipo de Tráfego"] = list_trafego
      df_enri["Qtd Seguidores"] = list_seguidores
      df_enri["Segmentação"] = list_segmentacao
      df_enri["O que é"] = list_o_que_e

      st.success("Enriquecimento de dados concluído!")
      st.dataframe(df_enri[["Tipo de Tráfego", "Qtd Seguidores", "Segmentação", "O que é"]].head(10))

      buf_enri = io.BytesIO()
      df_enri.to_excel(buf_enri, index=False)
      buf_enri.seek(0)

      st.download_button(
          "📥 Baixar Base Principal Enriquecida",
          data=buf_enri,
          file_name="Base_Principal_Enriquecida.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )


# ==============================================================================
# ABA 3: RELATÓRIO AVANÇADO (BASE GIGANTE - OTIMIZADA PARA MEMÓRIA)
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

        # --- EXPORTAR BASE VÁLIDA (SEM PINTAR PARA POUPAR MEMÓRIA) ---
        buf_final = io.BytesIO()
        df_final.to_excel(buf_final, index=False, engine='openpyxl')
        buf_final.seek(0)

        st.download_button(
            label="✅ 1. Baixar Relatório com Links Classificados (9 Colunas)",
            data=buf_final, 
            file_name="1_Relatorio_Links_Classificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- EXPORTAR BASE SEM LINKS (SEM PINTAR PARA POUPAR MEMÓRIA) ---
        buf_erros = io.BytesIO()
        df_sem_link.to_excel(buf_erros, index=False, engine='openpyxl')
        buf_erros.seek(0)

        st.download_button(
            label="⚠️ 2. Baixar Relatório de Contatos Pendentes (Sem Link)",
            data=buf_erros, 
            file_name="2_Relatorio_Para_Contatar_Sem_Link.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
