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
tab_comparador, tab_enriquecedor, tab_gigante, tab_copia_cola = st.tabs([
    "🔀 1. Comparar Duplicados",
    "🔍 2. Enriquecer Links",
    "📑 3. Relatório Avançado (Base Gigante)",
    "📋 4. Copiar e Colar (Rápido)"
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
  if re.search(r'\b(bet|cassino|aviator|slots|tigrinho|fortune|stake|blaze|poker|aposta|apostas)\b', texto): return "APOSTA"
  elif re.search(r'\b(investimento|investir|trader|acoes|ações|banco|renda|orcamento|finanças|finance|crypto|bitcoin|forex)\b', texto): return "FINANÇAS"
  elif re.search(r'\b(curso|cursos|mentoria|aula|aulas|ebook|edtech|concurso|faculdade|treinamento|escola)\b', texto): return "EDUCAÇÃO"
  elif re.search(r'\b(viagem|viajar|turismo|passagem|hospedagem|hotel|milhas)\b', texto): return "VIAGEM"
  elif re.search(r'\b(emagrecer|treino|dieta|diet|saude|saúde|whey|keto|fitness|nutri|estetica|estética|corpo)\b', texto): return "SAÚDE / FITNESS"
  elif re.search(r'\b(skin|make|unha|cabelo|maquiagem|cilios|sobrancelha|skincare)\b', texto): return "BELEZA / ESTÉTICA"
  elif re.search(r'\b(cupom|cupons|promo|promoção|promocao|desconto|oferta|ofertas)\b', texto): return "CUPOM / OFERTAS"
  elif re.search(r'\b(achado|achadinho|achados|achadinhos)\b', texto): return "ACHADINHOS"
  elif re.search(r'\b(shopee|amazon|magalu|aliexpress|ecommerce)\b', texto): return "E-COMMERCE"
  else: return "NÃO IDENTIFICADO"


def analisar_link_completo(url):
  if pd.isna(url) or not str(url).startswith("http"):
    return "Não Identificado", "Sem Link Válido", "NÃO IDENTIFICADO", "Sem Link"

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
        qtd_seguidores = "Bloqueado para acesso (Página Privada)"
        o_que_e = f"Perfil/Canal de {tipo_trafego}" if segmentacao == "NÃO IDENTIFICADO" else f"Canal de {tipo_trafego} focado em {segmentacao.title()}"
      else:
        qtd_seguidores = "Não aplicável (Site Próprio)"
        o_que_e = titulo[:60]

      return tipo_trafego, qtd_seguidores, segmentacao, o_que_e
    else:
      segmentacao = identificar_segmentacao(url_str)
      return tipo_trafego, "Acesso Restrito", segmentacao, "Página não acessível"
  except Exception:
    segmentacao = identificar_segmentacao(url_str)
    return tipo_trafego, "Erro de Conexão", segmentacao, "Site Indisponível / Erro"


# ==============================================================================
# ABA 1: COMPARADOR DE DUPLICADOS
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
        cnt_m, cnt_t = logins_m.value_counts(), logins_t.value_counts()
        status_list, cor_list = [], []

        for _, row in df_t.iterrows():
          login_limpo = str(row[col_login_t]).strip().lower()
          qtd_p, qtd_t = cnt_m.get(login_limpo, 0), cnt_t.get(login_limpo, 0)

          if qtd_p == 0: status_list.append("Novo (Não existe na principal)"); cor_list.append("Laranja")
          elif qtd_p > 0 and qtd_t == 1: status_list.append("1 para 1 (Existe, sem duplicidade)"); cor_list.append("Verde")
          elif qtd_p > 0 and qtd_t > 1: status_list.append("Duplicidade (Repetido)"); cor_list.append("Azul")

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

        idx_col = next((idx for idx, cel in enumerate(ws[1], start=1) if cel.value == col_login_t), None)
        if idx_col:
          for n_linha, cor in enumerate(cor_list, start=2):
            c = ws.cell(row=n_linha, column=idx_col)
            if cor == "Laranja": c.fill, c.font = f_laranja, font_b
            elif cor == "Verde": c.fill, c.font = f_verde, font_w
            elif cor == "Azul": c.fill, c.font = f_azul, font_w

        buf_final = io.BytesIO()
        wb.save(buf_final)
        buf_final.seek(0)
        st.success("Comparação concluída!")
        st.download_button("📥 Baixar Base Teste Classificada", data=buf_final, file_name="Base_Teste_Classificada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==============================================================================
# ABA 2: ENRIQUECEDOR DE LINKS
# ==============================================================================
with tab_enriquecedor:
  st.write("Envie a **Base Principal** para varrer os links e preencher automaticamente as novas colunas.")
  arq_unico = st.file_uploader("Envie a Base Principal", type=["xlsx", "csv"], key="enri_main")
  
  if arq_unico:
    df_enri = pd.read_csv(arq_unico) if arq_unico.name.endswith(".csv") else pd.read_excel(arq_unico)
    coluna_link_selecionada = st.selectbox("Selecione a coluna dos links:", options=list(df_enri.columns))

    if st.button("🔍 Iniciar Enriquecimento dos Links"):
      barra = st.progress(0)
      status_txt = st.empty()
      list_trafego, list_seguidores, list_segmentacao, list_o_que_e = [], [], [], []
      total = len(df_enri)

      for idx, link in enumerate(df_enri[coluna_link_selecionada]):
        barra.progress((idx + 1) / total)
        status_txt.text(f"Processando link {idx + 1} de {total}...")
        trafego, seg, segm, oquee = analisar_link_completo(link)
        list_trafego.append(trafego); list_seguidores.append(seg); list_segmentacao.append(segm); list_o_que_e.append(oquee)

      df_enri["Tipo de Tráfego"], df_enri["Qtd Seguidores"], df_enri["Segmentação"], df_enri["O que é"] = list_trafego, list_seguidores, list_segmentacao, list_o_que_e
      st.success("Enriquecimento de dados concluído!")
      
      buf_enri = io.BytesIO()
      df_enri.to_excel(buf_enri, index=False)
      buf_enri.seek(0)
      st.download_button("📥 Baixar Base Enriquecida", data=buf_enri, file_name="Base_Principal_Enriquecida.xlsx")


# ==============================================================================
# ABA 3: RELATÓRIO AVANÇADO (BASE GIGANTE)
# ==============================================================================
with tab_gigante:
  st.header("Processador de Bases Gigantes")
  arq_gigante = st.file_uploader("Envie a Base Completa (.xlsx)", type=["xlsx"], key="u_gigante")

  if arq_gigante:
    if st.button("🚀 Iniciar Processamento Profundo (Scraping)", key="btn_gigante"):
      status_leitura = st.info("A ler e juntar as abas. Aguarde...")
      todas_as_abas = pd.read_excel(arq_gigante, sheet_name=None)
      lista_df = [df.assign(**{"Mes_Origem (Aba)": nome}) for nome, df in todas_as_abas.items()]
      df_completo = pd.concat(lista_df, ignore_index=True).dropna(how='all')
      
      col_fonte = next((c for c in df_completo.columns if 'fonte' in str(c).lower()), df_completo.columns[4])
      df_completo[col_fonte] = df_completo[col_fonte].astype(str).str.strip()

      termos_excluir = ["erro", "sem link", "não cadastrada", "nao cadastrada", "-", "nan", "", "nenhuma", "null"]
      mascara_lixo = df_completo[col_fonte].str.lower().isin(termos_excluir)
      mascara_sem_http = ~df_completo[col_fonte].str.contains("http|www|\.com|\.br|@", case=False, na=False)

      df_sem_link = df_completo[mascara_lixo | mascara_sem_http].copy()
      df_limpo = df_completo[~(mascara_lixo | mascara_sem_http)].copy()
      
      status_leitura.empty()
      st.info(f"A visitar {len(df_limpo)} links válidos na internet...")

      list_t, list_seg, list_segm, list_oq = [], [], [], []
      barra_progresso = st.progress(0)
      texto_progresso = st.empty()
      total_validos = len(df_limpo)

      for idx, link in enumerate(df_limpo[col_fonte]):
        barra_progresso.progress((idx + 1) / total_validos)
        texto_progresso.text(f"Classificando link {idx + 1} de {total_validos}...")
        t, s, sm, oq = analisar_link_completo(link)
        list_t.append(t); list_seg.append(s); list_segm.append(sm); list_oq.append(oq)

      df_limpo["Segmento"], df_limpo["Categoria"], df_limpo["Breve Descrição"], df_limpo["Qtd Seguidores"] = list_segm, list_t, list_oq, list_seg
      df_final = df_limpo.sort_values(by=["Segmento", "Categoria"], ascending=True)
      st.success("🎉 Processamento Profundo concluído!")

      buf_final = io.BytesIO()
      df_final.to_excel(buf_final, index=False, engine='openpyxl')
      buf_final.seek(0)
      st.download_button("✅ 1. Baixar Relatório Classificados", data=buf_final, file_name="Relatorio_Links_Classificados.xlsx")

      csv_erros = df_sem_link.to_csv(index=False, sep=';', encoding='utf-8-sig')
      st.download_button("⚠️ 2. Baixar Contatos Pendentes", data=csv_erros, file_name="Contatar_Sem_Link.csv", mime="text/csv")


# ==============================================================================
# ABA 4: COPIAR E COLAR DADOS (RÁPIDO)
# ==============================================================================
with tab_copia_cola:
  st.header("Processador Rápido (Copiar e Colar)")
  st.write("Copie as colunas diretamente da sua planilha (ex: coluna de Nomes e coluna de Links) e cole no quadro abaixo.")
  
  dados_colados = st.text_area("Cole os dados aqui (pode colar múltiplas colunas):", height=250)

  if st.button("🚀 Processar Dados Colados"):
    if not dados_colados.strip():
        st.warning("Por favor, cole alguns dados no quadro antes de processar.")
    else:
        # Lê os dados colados do Excel (que usa tabulações para separar as colunas)
        try:
            df_colado = pd.read_csv(io.StringIO(dados_colados), sep='\t', header=None)
            
            # Tenta encontrar automaticamente qual coluna contém os links "http"
            col_fonte = None
            for col in df_colado.columns:
                if df_colado[col].astype(str).str.contains("http", case=False, na=False).any():
                    col_fonte = col
                    break
            
            if col_fonte is None:
                col_fonte = df_colado.columns[-1] # Se não achar, assume a última coluna
            
            # Adiciona nomes genéricos às colunas para o Excel final ficar bonito
            novos_nomes = [f"Coluna {i+1}" for i in range(len(df_colado.columns))]
            novos_nomes[col_fonte] = "Link (Fonte)"
            df_colado.columns = novos_nomes
            
            df_colado["Link (Fonte)"] = df_colado["Link (Fonte)"].astype(str).str.strip()

            termos_excluir = ["erro", "sem link", "não cadastrada", "nao cadastrada", "-", "nan", "", "nenhuma", "null"]
            mascara_lixo = df_colado["Link (Fonte)"].str.lower().isin(termos_excluir)
            mascara_sem_http = ~df_colado["Link (Fonte)"].str.contains("http|www|\.com|\.br|@", case=False, na=False)

            df_sem_link = df_colado[mascara_lixo | mascara_sem_http].copy()
            df_limpo = df_colado[~(mascara_lixo | mascara_sem_http)].copy()

            st.info(f"Encontrados {len(df_limpo)} links para analisar e {len(df_sem_link)} sem link.")

            list_t, list_seg, list_segm, list_oq = [], [], [], []
            barra_progresso = st.progress(0)
            texto_progresso = st.empty()
            total_validos = len(df_limpo)

            for idx, link in enumerate(df_limpo["Link (Fonte)"]):
                if total_validos > 0: barra_progresso.progress((idx + 1) / total_validos)
                texto_progresso.text(f"A analisar o link {idx + 1} de {total_validos}...")
                
                # Executa a mesma inteligência da Aba 3
                t, s, sm, oq = analisar_link_completo(link)
                list_t.append(t); list_seg.append(s); list_segm.append(sm); list_oq.append(oq)

            if total_validos > 0:
                df_limpo["Segmento"] = list_segm
                df_limpo["Categoria"] = list_t
                df_limpo["Breve Descrição"] = list_oq
                df_limpo["Qtd Seguidores"] = list_seg

            df_final = df_limpo.sort_values(by=["Segmento", "Categoria"], ascending=True) if total_validos > 0 else df_limpo
            
            st.success("🎉 Processamento dos dados colados concluído!")

            col_btn1, col_btn2 = st.columns(2)
            
            # --- EXPORTAR BASE VÁLIDA ---
            buf_final = io.BytesIO()
            df_final.to_excel(buf_final, index=False, engine='openpyxl')
            buf_final.seek(0)
            with col_btn1:
                st.download_button("✅ 1. Baixar Classificados (Excel)", data=buf_final, file_name="CopiaCola_Classificados.xlsx")

            # --- EXPORTAR BASE SEM LINKS ---
            csv_erros = df_sem_link.to_csv(index=False, sep=';', encoding='utf-8-sig')
            with col_btn2:
                st.download_button("⚠️ 2. Baixar Contatos Pendentes (CSV)", data=csv_erros, file_name="CopiaCola_Sem_Link.csv", mime="text/csv")
                
        except Exception as e:
            st.error(f"Erro ao ler os dados colados. Certifique-se de que os copiou corretamente do Excel. Erro: {e}")
