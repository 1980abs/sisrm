# ==============================================================================
# ABA 3: RELATÓRIO AVANÇADO (BASE GIGANTE - MÚLTIPLAS ABAS)
# ==============================================================================
with tab_gigante:
  st.header("Processador de Bases Gigantes (Excel com Múltiplas Abas)")
  st.write("1. Une todas as abas mensais numa só base.")
  st.write("2. Analisa os links criando 4 colunas (Segmento, Categoria, Descrição, Seguidores).")
  st.write("3. Gera DOIS relatórios: um com links classificados e outro com os erros/sem link para contactar.")

  arq_gigante = st.file_uploader(
      "Envie a Base de Afiliados Completa (.xlsx)",
      type=["xlsx"],
      key="upload_gigante",
  )

  if arq_gigante:
    if st.button("🚀 Processar Base e Gerar Relatórios"):
      with st.spinner("A processar todas as abas e a classificar a base... isto pode demorar 1-2 minutos."):
        
        # Lê todas as abas
        todas_as_abas = pd.read_excel(arq_gigante, sheet_name=None)
        lista_df = []
        for nome_aba, df_aba in todas_as_abas.items():
          df_aba["Mes_Origem (Aba)"] = nome_aba
          lista_df.append(df_aba)

        df_completo = pd.concat(lista_df, ignore_index=True)
        df_completo = df_completo.dropna(how='all') # Remove linhas totalmente vazias
        
        # Identifica a coluna Fonte
        col_fonte = next((c for c in df_completo.columns if 'fonte' in str(c).lower()), df_completo.columns[4])
        df_completo[col_fonte] = df_completo[col_fonte].astype(str).str.strip()

        # Separar quem tem link válido de quem não tem
        termos_excluir = ["erro", "sem link", "não cadastrada", "nao cadastrada", "-", "nan", "", "nenhuma"]
        mascara_lixo = df_completo[col_fonte].str.lower().isin(termos_excluir)
        mascara_sem_http = ~df_completo[col_fonte].str.contains("http|www|\.com|\.br|@", case=False, na=False)

        # Base 2: RELATÓRIO DE ERROS / SEM LINK (Para entrar em contato)
        df_sem_link = df_completo[mascara_lixo | mascara_sem_http].copy()
        
        # Base 1: RELATÓRIO COM LINKS VÁLIDOS (Para classificar)
        df_limpo = df_completo[~(mascara_lixo | mascara_sem_http)].copy()

        st.info(f"Análise concluída: {len(df_limpo)} links válidos encontrados e {len(df_sem_link)} registos sem link para contactar.")

        # Classificação da Base 1
        list_trafego, list_seguidores, list_segmentacao, list_oquee = [], [], [], []

        for link in df_limpo[col_fonte]:
          trafego, seg, segm, oquee = analisar_link_rapido(link)
          list_trafego.append(trafego)
          
          # Ajuste do texto para "Página Privada" caso seja rede social
          if trafego == "Site Próprio / Landing Page":
              list_seguidores.append("NÃO APLICÁVEL (SITE)")
          else:
              list_seguidores.append("PÁGINA PRIVADA / REQUER API")
              
          list_segmentacao.append(segm)
          
          # Ajuste da descrição
          if trafego == "Site Próprio / Landing Page":
              list_oquee.append(f"Página Web focada em {segm.title()}")
          else:
              list_oquee.append(f"Canal de {trafego.title()} focado em {segm.title()}")

        # Adiciona as 4 colunas solicitadas
        df_limpo["Segmento"] = list_segmentacao
        df_limpo["Categoria"] = list_trafego
        df_limpo["Breve Descrição"] = list_oquee
        df_limpo["Qtd Seguidores"] = list_seguidores

        # Ordenar por Segmento e Categoria
        df_final = df_limpo.sort_values(by=["Segmento", "Categoria"], ascending=True)

        st.success("Tudo pronto! Pode descarregar os seus dois relatórios abaixo:")

        # ==========================================
        # BOTÃO 1: Relatório Principal Formatado
        # ==========================================
        buf_final = io.BytesIO()
        # Salvamos com o openpyxl para formatar as cores
        df_final.to_excel(buf_final, index=False, engine='openpyxl')
        buf_final.seek(0)
        
        # Formatação do cabeçalho
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
            data=buf_formatado,
            file_name="1_Relatorio_Links_Classificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ==========================================
        # BOTÃO 2: Relatório Sem Links (Contatos)
        # ==========================================
        buf_erros = io.BytesIO()
        df_sem_link.to_excel(buf_erros, index=False, engine='openpyxl')
        buf_erros.seek(0)
        
        wb2 = load_workbook(buf_erros)
        ws2 = wb2.active
        header_fill_erro = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid") # Vermelho
        for cell in ws2[1]:
            cell.fill = header_fill_erro
            cell.font = header_font
        ws2.auto_filter.ref = ws2.dimensions
        
        buf_erros_formatado = io.BytesIO()
        wb2.save(buf_erros_formatado)
        buf_erros_formatado.seek(0)

        st.download_button(
            label="⚠️ 2. Baixar Relatório de Contatos Pendentes (Sem Link)",
            data=buf_erros_formatado,
            file_name="2_Relatorio_Para_Contatar_Sem_Link.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
