import os
import io
import json
import datetime as dt
import streamlit as st
import openpyxl
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# CONFIGURAÇÃO DE ACESSO
# ==============================================================================
ID_PASTA_DRIVE = "1YXz96rcj7IgthzYHK6uyei_KrC5q9cqb"

def obter_serviço_drive():
    creds = Credentials(
        token=None,
        refresh_token=st.secrets["gcp_oauth"]["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=st.secrets["gcp_oauth"]["client_id"],
        client_secret=st.secrets["gcp_oauth"]["client_secret"]
    )
    if not creds.valid:
        creds.refresh(Request())
    return build('drive', 'v3', credentials=creds)

def calcular_pontos(g_m_real, g_v_real, g_m_palpite, g_v_palpite):
    if g_m_real is None or g_v_real is None or g_m_palpite is None or g_v_palpite is None: return 0
    try:
        g_m_real, g_v_real = int(g_m_real), int(g_v_real)
        g_m_palpite, g_v_palpite = int(g_m_palpite), int(g_v_palpite)
    except (ValueError, TypeError): return 0
    
    res_real = 1 if g_m_real > g_v_real else (-1 if g_m_real < g_v_real else 0)
    res_palpite = 1 if g_m_palpite > g_v_palpite else (-1 if g_m_palpite < g_v_palpite else 0)
    if res_real != res_palpite: return 0
    if g_m_real == g_m_palpite and g_v_real == g_v_palpite: return 7
    if res_real == 0: return 3
    if g_m_real == g_m_palpite or g_v_real == g_v_palpite: return 4
    return 2

def gerar_tabela_html(participantes, titulo, subtitulo, cor_destaque=None, posicoes_destaque=None):
    if posicoes_destaque is None: posicoes_destaque = []
    
    html = f"<div style='background: #0F172A; padding: 15px; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.35); text-align: center; margin-bottom: 20px; font-family: sans-serif;'>"
    html += f"<h4 style='margin-bottom: 0px; color: #F8FAFC;'>{titulo}</h4>"
    html += f"<p style='margin-top: 0px; font-size: 14px; font-style: italic; color: #CBD5E1;'>{subtitulo}</p>"
    
    html += "<table style='width: 100%; border-collapse: collapse; text-align: center; font-size: 14px; color: #F8FAFC; margin-top: 15px;'>"
    html += "<thead style='background-color: #111827; border-bottom: 2px solid #334155;'>"
    html += "<tr><th style='padding: 8px;'>Pos</th><th style='padding: 8px;'>Var</th><th style='text-align: left; padding: 8px;'>Nome</th><th style='padding: 8px;'>Pts</th><th style='padding: 8px; color:#94A3B8;'>Dif</th></tr>"
    html += "</thead><tbody>"

    for p in participantes:
        pos_str = p['posicao'].replace("º Lugar", "").strip()
        is_highlight = (pos_str.isdigit() and int(pos_str) in posicoes_destaque)
        bg_color = cor_destaque if is_highlight else "#1E293B"
        
        html += f"<tr style='background-color: {bg_color}; border-bottom: 1px solid #334155;'>"
        html += f"<td style='padding: 6px; font-weight: bold;'>{p['posicao'].replace(' Lugar', '')}</td>"
        html += f"<td style='padding: 6px; font-size: 13px;'>{p.get('var_html', '➖')}</td>"
        html += f"<td style='text-align: left; padding: 6px;'>{p.get('nome_exibicao', p['nome'])}</td>"
        html += f"<td style='padding: 6px; font-weight: bold;'>{p['pontos']}</td>"
        html += f"<td style='padding: 6px; font-size: 12px; color:#94A3B8;'>{p.get('dif', '-')}</td>"
        html += "</tr>"

    html += "</tbody></table></div>"
    return html

# ==============================================================================
# CONFIGURAÇÃO VISUAL E GESTÃO DE ESTADO
# ==============================================================================
st.set_page_config(page_title="Bolão Copa 2026 (THE)", page_icon="⚽", layout="wide")

if 'ranking_processado' not in st.session_state:
    st.session_state.ranking_processado = []

st.title("🏆 Apuração do Bolão da Copa 2026 (THE) - V2")
st.write("Clique no botão abaixo para processar os palpites no Drive e atualizar o ranking em tempo real.")

if st.button("🚀 Atualizar Classificação", type="primary"):
    with st.spinner("Lendo planilhas e histórico no Google Drive..."):
        try:
            service = obter_serviço_drive()
            query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
            resultados = service.files().list(q=query, fields="files(id, name)").execute()
            mapa_arquivos = {arq['name']: arq['id'] for arq in resultados.get('files', [])}
            
            if 'Arquivo_de_controle.xlsx' not in mapa_arquivos:
                st.error("Erro: 'Arquivo_de_controle.xlsx' não encontrado na pasta do Drive.")
                st.stop()
                
            # --- LER HISTÓRICO JSON (OPÇÃO B) ---
            fuso_br = dt.timezone(dt.timedelta(hours=-3))
            hoje_str = dt.datetime.now(fuso_br).strftime('%Y-%m-%d')
            dados_historico = {"data_referencia": "", "posicoes": {}}
            
            nome_json = "historico_bolao.json"
            if nome_json in mapa_arquivos:
                id_json = mapa_arquivos[nome_json]
                req_json = service.files().get_media(fileId=id_json)
                bytes_json = io.BytesIO()
                baix_json = MediaIoBaseDownload(bytes_json, req_json)
                while not baix_json.next_chunk()[1]: pass
                bytes_json.seek(0)
                try: dados_historico = json.loads(bytes_json.read().decode('utf-8'))
                except: pass
            
            posicoes_antigas = dados_historico.get("posicoes", {})
            data_hist = dados_historico.get("data_referencia", "")
            nova_foto_diaria = (hoje_str != data_hist)
            # ------------------------------------

            # LER ARQUIVO DE CONTROLE
            id_controle = mapa_arquivos['Arquivo_de_controle.xlsx']
            requisicao = service.files().get_media(fileId=id_controle)
            bytes_controle = io.BytesIO()
            baixador = MediaIoBaseDownload(bytes_controle, requisicao)
            concluido = False
            while not concluido: _, concluido = baixador.next_chunk()
                
            bytes_controle.seek(0)
            wb_leitura = openpyxl.load_workbook(bytes_controle, data_only=True)
            ws_fase1_leitura = wb_leitura['FASE 1']
            
            resultados_reais = {}
            for r in range(3, 75):
                resultados_reais[r] = (ws_fase1_leitura[f'G{r}'].value, ws_fase1_leitura[f'I{r}'].value)
                
            ws_tabela_leitura = wb_leitura['TABELA']
            lista_ranking = []
            
            # PROCESSAR PLANILHAS
            for row_tabela in range(2, 101):
                celula_nome = ws_tabela_leitura[f'A{row_tabela}'].value
                if celula_nome is None or str(celula_nome).strip() == "" or str(celula_nome).strip().lower() == "participante": continue
                    
                nome = str(celula_nome).strip()
                arquivo_palpite = f"{nome}.xlsx"
                
                if arquivo_palpite not in mapa_arquivos:
                    lista_ranking.append({'nome': nome, 'pontos': 0})
                    continue
                    
                id_part = mapa_arquivos[arquivo_palpite]
                req_part = service.files().get_media(fileId=id_part)
                bytes_part = io.BytesIO()
                baixador_part = MediaIoBaseDownload(bytes_part, req_part)
                concluido_part = False
                while not concluido_part: _, concluido_part = baixador_part.next_chunk()
                    
                bytes_part.seek(0)
                wb_part = openpyxl.load_workbook(bytes_part, data_only=True)
                ws_part_fase1 = wb_part['FASE 1']
                
                pontos_totais = 0
                for r in range(3, 75):
                    g_m_real, g_v_real = resultados_reais[r]
                    if g_m_real is None or g_v_real is None: continue
                    pontos_totais += calcular_pontos(g_m_real, g_v_real, ws_part_fase1[f'G{r}'].value, ws_part_fase1[f'I{r}'].value)
                    
                wb_part.close()
                lista_ranking.append({'nome': nome, 'pontos': pontos_totais})
                
            wb_leitura.close()
            lista_ranking.sort(key=lambda x: (-x['pontos'], x['nome'].lower()))
            
            posicao_atual = 1
            pontos_anteriores = None
            posicoes_para_json = {}
            
            for idx, participante in enumerate(lista_ranking):
                if participante['pontos'] != pontos_anteriores:
                    posicao_atual = idx + 1
                pontos_anteriores = participante['pontos']
                participante['posicao'] = f"{posicao_atual}º Lugar"
                
                # Medalhas Automáticas
                emoji = ""
                if posicao_atual == 1: emoji = "🥇 "
                elif posicao_atual == 2: emoji = "🥈 "
                elif posicao_atual == 3: emoji = "🥉 "
                participante['nome_exibicao'] = f"{emoji}{participante['nome']}"
                
                # Distância para o próximo
                if idx < len(lista_ranking) - 1:
                    pts_below = lista_ranking[idx+1]['pontos']
                    dif = participante['pontos'] - pts_below
                    participante['dif'] = f"+{dif}" if dif > 0 else "="
                else:
                    participante['dif'] = "-"

                # Variação de Posição via JSON Histórico
                pos_antiga = posicoes_antigas.get(participante['nome'], posicao_atual)
                variacao = pos_antiga - posicao_atual
                
                if variacao > 0:
                    participante['var_html'] = f"<span style='color: #22C55E;'>▲ {variacao}</span>"
                elif variacao < 0:
                    participante['var_html'] = f"<span style='color: #EF4444;'>▼ {abs(variacao)}</span>"
                else:
                    participante['var_html'] = "<span style='color: #94A3B8;'>➖</span>"

                # Guarda a posição para a foto diária se precisar atualizar
                posicoes_para_json[participante['nome']] = posicao_atual

            # ATUALIZAR JSON NO DRIVE SE FOR UM NOVO DIA
            if nova_foto_diaria:
                novo_historico = {
                    "data_referencia": hoje_str,
                    "posicoes": posicoes_para_json
                }
                json_bytes = io.BytesIO(json.dumps(novo_historico).encode('utf-8'))
                media_json = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=True)
                
                if nome_json in mapa_arquivos:
                    service.files().update(fileId=mapa_arquivos[nome_json], media_body=media_json).execute()
                else:
                    service.files().create(body={'name': nome_json, 'parents': [ID_PASTA_DRIVE]}, media_body=media_json).execute()
                
                # Como é o primeiro update do dia, reseta visualmente as variações para 0
                for p in lista_ranking: p['var_html'] = "<span style='color: #94A3B8;'>➖</span>"

            # GRAVAR NO EXCEL (Planilha de Controle)
            bytes_controle.seek(0)
            wb_gravar = openpyxl.load_workbook(bytes_controle, data_only=False)
            if 'Dados Pessoais' in wb_gravar.sheetnames: wb_gravar.remove(wb_gravar['Dados Pessoais'])
            ws_tabela_gravar = wb_gravar['TABELA']
            for r in range(1, 101): ws_tabela_gravar[f'A{r}'].value, ws_tabela_gravar[f'B{r}'].value, ws_tabela_gravar[f'C{r}'].value = None, None, None
            ws_tabela_gravar['A1'].value, ws_tabela_gravar['B1'].value, ws_tabela_gravar['C1'].value = "Participante", "Pontos", "Posição"
            
            for idx, participante in enumerate(lista_ranking):
                linha_atual = 2 + idx
                ws_tabela_gravar[f'A{linha_atual}'].value = participante['nome']
                ws_tabela_gravar[f'B{linha_atual}'].value = participante['pontos']
                ws_tabela_gravar[f'C{linha_atual}'].value = participante['posicao']
                
            st.session_state.ranking_processado = lista_ranking.copy()
            
            arquivo_saida_bytes = io.BytesIO()
            wb_gravar.save(arquivo_saida_bytes)
            arquivo_saida_bytes.seek(0)
            wb_gravar.close()
            
            nome_saida = 'BOLÃO DA COPA DO MUNDO 2026 (THE).xlsx'
            mimetype_excel = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            media_upload = MediaIoBaseUpload(arquivo_saida_bytes, mimetype=mimetype_excel, resumable=True)
            if nome_saida in mapa_arquivos: service.files().update(fileId=mapa_arquivos[nome_saida], media_body=media_upload).execute()
            else: service.files().create(body={'name': nome_saida, 'parents': [ID_PASTA_DRIVE]}, media_body=media_upload).execute()
            
            st.markdown("<div style='background: #064E3B; color: #D1FAE5; border-left: 4px solid #10B981; padding: 12px; border-radius: 4px; margin-bottom: 20px;'><strong>Tabela atualizada com sucesso no Google Drive!</strong></div>", unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento: {e}")

st.divider()

# ==============================================================================
# NAVEGAÇÃO EM ABAS (UI)
# ==============================================================================
aba_ranking, aba_palpites = st.tabs(["🏆 Classificação e Resenha", "👀 Espiar Palpites da Rodada"])

with aba_ranking:
    if not st.session_state.ranking_processado:
        st.info("👆 Clique no botão azul lá em cima para buscar os dados no Google Drive e gerar a classificação atualizada.")
    else:
        dados = st.session_state.ranking_processado
        N = len(dados)
        
        # --- Card de Resumo do Líder Geral ---
        if dados:
            lider = dados[0]
            st.markdown(f"""
            <div style="background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%); border-left: 4px solid #F59E0B; padding: 15px; border-radius: 6px; margin-bottom: 25px;">
                <h4 style="color: #FBBF24; margin: 0 0 5px 0;">🏆 Líder Geral</h4>
                <p style="color: #F8FAFC; margin: 0; font-size: 16px;"><strong>{lider['nome']}</strong> segue no topo da tabela com <strong>{lider['pontos']} pontos</strong>!</p>
            </div>
            """, unsafe_allow_html=True)
        # -----------------------------------------------

        idx_prof = 5 if N >= 5 else N
        while idx_prof < N and dados[idx_prof]['pontos'] == dados[idx_prof - 1]['pontos']: idx_prof += 1
            
        idx_lant = max(N - 3, idx_prof) 
        while idx_lant > idx_prof and dados[idx_lant - 1]['pontos'] == dados[idx_lant]['pontos']: idx_lant -= 1
            
        alvo_amad = idx_prof + 10
        idx_amad = min(alvo_amad, idx_lant) 
        while idx_amad < idx_lant and dados[idx_amad]['pontos'] == dados[idx_amad - 1]['pontos']: idx_amad += 1
            
        profissionais, amadores = dados[:idx_prof], dados[idx_prof:idx_amad]
        peladeiros, lanterna = dados[idx_amad:idx_lant], dados[idx_lant:]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if profissionais: st.markdown(gerar_tabela_html(profissionais, "Profissionais", "- Elite do Pitaco -", "#166534", [1, 2, 3, 4, 5]), unsafe_allow_html=True)
        with col2:
            if amadores: st.markdown(gerar_tabela_html(amadores, "Amadores", "- Os que ainda sonham -"), unsafe_allow_html=True)
        with col3:
            if peladeiros: st.markdown(gerar_tabela_html(peladeiros, "Peladeiros", "- Especialistas em Errar -"), unsafe_allow_html=True)
            
        st.write("---")
        col_vazia1, col_lanterna, col_vazia2 = st.columns([1, 2, 1])
        with col_lanterna:
            if lanterna:
                posicoes_lanterna = [int(p['posicao'].replace("º Lugar", "").strip()) for p in lanterna]
                st.markdown(gerar_tabela_html(lanterna, "Prêmio Espírito Coletivo", "- Bastava Apostar ao Contrário -", "#991B1B", posicoes_lanterna), unsafe_allow_html=True)

with aba_palpites:
    st.write("O código da Sessão 2 (Espiar Palpites) continua o mesmo que alinhamos anteriormente.")
