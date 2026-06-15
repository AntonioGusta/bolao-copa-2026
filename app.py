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

@st.cache_data(ttl=dt.timedelta(hours=24), show_spinner=False)
def carregar_palpites_em_cache(file_id):
    service = obter_serviço_drive()
    req = service.files().get_media(fileId=file_id)
    bytes_io = io.BytesIO()
    baixador = MediaIoBaseDownload(bytes_io, req)
    while not baixador.next_chunk()[1]: pass
    bytes_io.seek(0)
    
    wb = openpyxl.load_workbook(bytes_io, data_only=True)
    ws = wb['FASE 1']
    palpites = {}
    for r in range(3, 75):
        palpites[r] = (ws[f'G{r}'].value, ws[f'I{r}'].value)
    wb.close()
    return palpites

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

if 'ranking_processado' not in st.session_state: st.session_state.ranking_processado = []
if 'resumo_ontem' not in st.session_state: st.session_state.resumo_ontem = {}

st.title("🏆 Apuração do Bolão da Copa 2026 (THE) - V2")
st.write("Clique no botão abaixo para processar os palpites e atualizar o ranking em tempo real.")

if st.button("🚀 Atualizar Classificação", type="primary"):
    with st.spinner("Sincronizando com o Drive..."):
        try:
            service = obter_serviço_drive()
            query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
            resultados = service.files().list(q=query, fields="files(id, name)").execute()
            mapa_arquivos = {arq['name']: arq['id'] for arq in resultados.get('files', [])}
            
            if 'Arquivo_de_controle.xlsx' not in mapa_arquivos:
                st.error("Erro: 'Arquivo_de_controle.xlsx' não encontrado na pasta do Drive.")
                st.stop()
                
            # ==================================================================
            # 1. PREPARANDO O BANCO DE HISTÓRICO
            # ==================================================================
            fuso_br = dt.timezone(dt.timedelta(hours=-3))
            hoje_date = dt.datetime.now(fuso_br).date()
            ontem_date = hoje_date - dt.timedelta(days=1)
            anteontem_date = hoje_date - dt.timedelta(days=2)
            
            hoje_str = hoje_date.strftime('%Y-%m-%d')
            dados_historico = {"historico": {}}
            
            nome_json = "historico_bolao.json"
            if nome_json in mapa_arquivos:
                id_json = mapa_arquivos[nome_json]
                req_json = service.files().get_media(fileId=id_json)
                bytes_json = io.BytesIO()
                baix_json = MediaIoBaseDownload(bytes_json, req_json)
                while not baix_json.next_chunk()[1]: pass
                bytes_json.seek(0)
                try: 
                    conteudo_lido = json.loads(bytes_json.read().decode('utf-8'))
                    if "historico" in conteudo_lido: dados_historico = conteudo_lido
                except: pass
            
            linha_do_tempo = dados_historico.get("historico", {})
            nova_foto_diaria = (hoje_str not in linha_do_tempo)

            # LER ARQUIVO DE CONTROLE
            id_controle = mapa_arquivos['Arquivo_de_controle.xlsx']
            requisicao = service.files().get_media(fileId=id_controle)
            bytes_controle = io.BytesIO()
            baixador = MediaIoBaseDownload(bytes_controle, requisicao)
            while not baixador.next_chunk()[1]: pass
                
            bytes_controle.seek(0)
            wb_leitura = openpyxl.load_workbook(bytes_controle, data_only=True)
            ws_fase1_leitura = wb_leitura['FASE 1']
            
            resultados_reais = {}
            datas_jogos = {}
            for r in range(3, 75):
                resultados_reais[r] = (ws_fase1_leitura[f'G{r}'].value, ws_fase1_leitura[f'I{r}'].value)
                
                # Coleta as datas reais dos jogos (Coluna C) para a Máquina do Tempo
                val_data = ws_fase1_leitura[f'C{r}'].value
                d_jogo = None
                if isinstance(val_data, dt.datetime):
                    d_jogo = val_data.date()
                elif val_data:
                    txt = str(val_data).strip()
                    if "/" in txt:
                        try: d_jogo = dt.date(2026, int(txt.split('/')[1]), int(txt.split('/')[0]))
                        except: pass
                datas_jogos[r] = d_jogo
                
            ws_tabela_leitura = wb_leitura['TABELA']
            lista_ranking = []
            
            # Listas para guardar a simulação do passado
            lista_anteontem = []
            lista_ontem = []
            
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
                palpites = carregar_palpites_em_cache(id_part)
                
                pts_tot = 0
                pts_ant = 0
                pts_ont = 0
                
                for r in range(3, 75):
                    g_m_real, g_v_real = resultados_reais[r]
                    if g_m_real is None or g_v_real is None: continue
                    
                    p_casa, p_fora = palpites.get(r, (None, None))
                    pts = calcular_pontos(g_m_real, g_v_real, p_casa, p_fora)
                    pts_tot += pts
                    
                    d_jogo = datas_jogos.get(r)
                    if d_jogo:
                        if d_jogo <= anteontem_date: pts_ant += pts
                        if d_jogo <= ontem_date: pts_ont += pts
                        
                lista_ranking.append({'nome': nome, 'pontos': pts_tot})
                lista_anteontem.append({'nome': nome, 'pontos': pts_ant})
                lista_ontem.append({'nome': nome, 'pontos': pts_ont})
                
            wb_leitura.close()
            
            # ==================================================================
            # O "TIME MACHINE" (RECONSTRUÇÃO EXCLUSIVA DO HISTÓRICO DE ONTEM)
            # ==================================================================
            if len(linha_do_tempo) < 2:
                def ranquear(lista):
                    lista.sort(key=lambda x: (-x['pontos'], x['nome'].lower()))
                    pos_dict = {}
                    pos_atual = 1
                    pts_anteriores = None
                    for idx, p in enumerate(lista):
                        if p['pontos'] != pts_anteriores: pos_atual = idx + 1
                        pts_anteriores = p['pontos']
                        pos_dict[p['nome']] = pos_atual
                    return pos_dict
                
                # Recriar as fotos passadas e injetar no JSON permanentemente
                linha_do_tempo[anteontem_date.strftime('%Y-%m-%d')] = ranquear(lista_anteontem)
                linha_do_tempo[ontem_date.strftime('%Y-%m-%d')] = ranquear(lista_ontem)
                nova_foto_diaria = True # Força o update do JSON
            # ==================================================================
            
            datas_passadas = sorted([d for d in linha_do_tempo.keys() if d < hoje_str])
            pos_fim_ontem = linha_do_tempo[datas_passadas[-1]] if len(datas_passadas) >= 1 else {}
            pos_inicio_ontem = linha_do_tempo[datas_passadas[-2]] if len(datas_passadas) >= 2 else pos_fim_ontem
            data_ontem_str = datas_passadas[-1] if len(datas_passadas) >= 1 else None

            # Cálculo de quem subiu e caiu ONDE IMPORTA (Ontem!)
            var_ontem_cards = {}
            for n, p_fin in pos_fim_ontem.items():
                p_ini = pos_inicio_ontem.get(n, p_fin)
                var_ontem_cards[n] = p_ini - p_fin
                
            m_subida = max(var_ontem_cards.values()) if var_ontem_cards else 0
            m_queda = min(var_ontem_cards.values()) if var_ontem_cards else 0
            
            h_sub = [n for n, v in var_ontem_cards.items() if v == m_subida] if m_subida > 0 else []
            v_que = [n for n, v in var_ontem_cards.items() if v == m_queda] if m_queda < 0 else []
            
            data_str_card = "Ontem"
            if data_ontem_str:
                dt_obj = dt.datetime.strptime(data_ontem_str, '%Y-%m-%d')
                data_str_card = f"{dt_obj.day:02d}/{dt_obj.month:02d}"

            st.session_state.resumo_ontem = {
                "data_str": data_str_card,
                "maior_subida": {"nomes": h_sub, "valor": m_subida},
                "maior_queda": {"nomes": v_que, "valor": abs(m_queda)}
            }

            lista_ranking.sort(key=lambda x: (-x['pontos'], x['nome'].lower()))
            
            posicao_atual = 1
            pontos_anteriores = None
            posicoes_para_json = {}
            
            for idx, participante in enumerate(lista_ranking):
                if participante['pontos'] != pontos_anteriores: posicao_atual = idx + 1
                pontos_anteriores = participante['pontos']
                participante['posicao'] = f"{posicao_atual}º Lugar"
                
                emoji = ""
                if posicao_atual == 1: emoji = "🥇 "
                elif posicao_atual == 2: emoji = "🥈 "
                elif posicao_atual == 3: emoji = "🥉 "
                participante['nome_exibicao'] = f"{emoji}{participante['nome']}"
                
                if idx < len(lista_ranking) - 1:
                    dif = participante['pontos'] - lista_ranking[idx+1]['pontos']
                    participante['dif'] = f"+{dif}" if dif > 0 else "="
                else: participante['dif'] = "-"

                # Variação da Tabela Ao Vivo (Fim de Ontem vs Agora)
                pos_antiga = pos_fim_ontem.get(participante['nome'], posicao_atual)
                variacao = pos_antiga - posicao_atual
                
                if variacao > 0: participante['var_html'] = f"<span style='color: #22C55E;'>▲ {variacao}</span>"
                elif variacao < 0: participante['var_html'] = f"<span style='color: #EF4444;'>▼ {abs(variacao)}</span>"
                else: participante['var_html'] = "<span style='color: #94A3B8;'>➖</span>"

                posicoes_para_json[participante['nome']] = posicao_atual

            # SALVAR SÉRIE TEMPORAL NO JSON
            if nova_foto_diaria:
                linha_do_tempo[hoje_str] = posicoes_para_json
                novo_historico = {"historico": linha_do_tempo}
                
                json_bytes = io.BytesIO(json.dumps(novo_historico).encode('utf-8'))
                media_json = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=True)
                
                if nome_json in mapa_arquivos: service.files().update(fileId=mapa_arquivos[nome_json], media_body=media_json).execute()
                else: service.files().create(body={'name': nome_json, 'parents': [ID_PASTA_DRIVE]}, media_body=media_json).execute()

            # GRAVAR EXCEL
            bytes_controle.seek(0)
            wb_gravar = openpyxl.load_workbook(bytes_controle, data_only=False)
            if 'Dados Pessoais' in wb_gravar.sheetnames: wb_gravar.remove(wb_gravar['Dados Pessoais'])
            ws_tabela_gravar = wb_gravar['TABELA']
            for r in range(1, 101): ws_tabela_gravar[f'A{r}'].value, ws_tabela_gravar[f'B{r}'].value, ws_tabela_gravar[f'C{r}'].value = None, None, None
            ws_tabela_gravar['A1'].value, ws_tabela_gravar['B1'].value, ws_tabela_gravar['C1'].value = "Participante", "Pontos", "Posição"
            
            for idx, p in enumerate(lista_ranking):
                linha_atual = 2 + idx
                ws_tabela_gravar[f'A{linha_atual}'].value = p['nome']
                ws_tabela_gravar[f'B{linha_atual}'].value = p['pontos']
                ws_tabela_gravar[f'C{linha_atual}'].value = p['posicao']
                
            st.session_state.ranking_processado = lista_ranking.copy()
            
            arquivo_saida_bytes = io.BytesIO()
            wb_gravar.save(arquivo_saida_bytes)
            arquivo_saida_bytes.seek(0)
            wb_gravar.close()
            
            nome_saida = 'BOLÃO DA COPA DO MUNDO 2026 (THE).xlsx'
            media_upload = MediaIoBaseUpload(arquivo_saida_bytes, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
            if nome_saida in mapa_arquivos: service.files().update(fileId=mapa_arquivos[nome_saida], media_body=media_upload).execute()
            else: service.files().create(body={'name': nome_saida, 'parents': [ID_PASTA_DRIVE]}, media_body=media_upload).execute()
            
            st.markdown("<div style='background: #064E3B; color: #D1FAE5; border-left: 4px solid #10B981; padding: 12px; border-radius: 4px; margin-bottom: 20px;'><strong>Tabela atualizada com sucesso!</strong></div>", unsafe_allow_html=True)
            
        except Exception as e: st.error(f"Ocorreu um erro no processamento: {e}")

st.divider()

# ==============================================================================
# NAVEGAÇÃO EM ABAS (UI)
# ==============================================================================
aba_ranking, aba_palpites = st.tabs(["🏆 Classificação e Resenha", "👀 Espiar Palpites da Rodada"])

with aba_ranking:
    if not st.session_state.ranking_processado: st.info("👆 Clique no botão azul lá em cima para buscar os dados.")
    else:
        dados = st.session_state.ranking_processado
        resumo = st.session_state.get('resumo_ontem', {})
        N = len(dados)
        
        # --- CARDS DE RESUMO (LÍDER, SUBIDA, QUEDA) ---
        if dados:
            lider = dados[0]
            col_c1, col_c2, col_c3 = st.columns(3)
            
            with col_c1:
                st.markdown(f"""
                <div style="background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%); border-left: 4px solid #F59E0B; padding: 15px; border-radius: 6px; margin-bottom: 25px; min-height: 105px;">
                    <h5 style="color: #FBBF24; margin: 0 0 5px 0;">🏆 Líder Geral</h5>
                    <p style="color: #F8FAFC; margin: 0; font-size: 14px;"><strong>{lider['nome']}</strong> segue no topo com <strong>{lider['pontos']} pts</strong>!</p>
                </div>
                """, unsafe_allow_html=True)
                
            with col_c2:
                if resumo and resumo.get('maior_subida', {}).get('nomes'):
                    nomes_sub = ", ".join(resumo['maior_subida']['nomes'])
                    val_sub = resumo['maior_subida']['valor']
                    data_str = resumo.get('data_str', 'Ontem')
                    st.markdown(f"""
                    <div style="background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%); border-left: 4px solid #22C55E; padding: 15px; border-radius: 6px; margin-bottom: 25px; min-height: 105px;">
                        <h5 style="color: #4ADE80; margin: 0 0 5px 0;">🚀 Maior Subida ({data_str})</h5>
                        <p style="color: #F8FAFC; margin: 0; font-size: 14px;"><strong>{nomes_sub}</strong> saltou <strong>+{val_sub} posições</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div style="background: #0F172A; border-left: 4px solid #334155; padding: 15px; border-radius: 6px; margin-bottom: 25px; min-height: 105px;">
                        <h5 style="color: #94A3B8; margin: 0 0 5px 0;">🚀 Maior Subida</h5>
                        <p style="color: #64748B; margin: 0; font-size: 14px;">Aguardando jogos...</p>
                    </div>
                    """, unsafe_allow_html=True)

            with col_c3:
                if resumo and resumo.get('maior_queda', {}).get('nomes'):
                    nomes_que = ", ".join(resumo['maior_queda']['nomes'])
                    val_que = resumo['maior_queda']['valor']
                    data_str = resumo.get('data_str', 'Ontem')
                    st.markdown(f"""
                    <div style="background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%); border-left: 4px solid #EF4444; padding: 15px; border-radius: 6px; margin-bottom: 25px; min-height: 105px;">
                        <h5 style="color: #F87171; margin: 0 0 5px 0;">📉 Maior Queda ({data_str})</h5>
                        <p style="color: #F8FAFC; margin: 0; font-size: 14px;"><strong>{nomes_que}</strong> caiu <strong>-{val_que} posições</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div style="background: #0F172A; border-left: 4px solid #334155; padding: 15px; border-radius: 6px; margin-bottom: 25px; min-height: 105px;">
                        <h5 style="color: #94A3B8; margin: 0 0 5px 0;">📉 Maior Queda</h5>
                        <p style="color: #64748B; margin: 0; font-size: 14px;">Aguardando jogos...</p>
                    </div>
                    """, unsafe_allow_html=True)
        # ----------------------------------------------

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
    st.subheader("👀 Espiar Palpites da Rodada")
    fuso_br = dt.timezone(dt.timedelta(hours=-3))
    hoje = dt.datetime.now(fuso_br)
    data_padrao = f"{hoje.day}/{hoje.month}"
    
    data_pesquisa = st.text_input("📅 Digite a data dos jogos que deseja ver (Ex: 12/6):", value=data_padrao)
    
    if st.button("🔍 Buscar Rodada", type="secondary"):
        with st.spinner("Analisando palpites (Utilizando cache ultra-rápido)..."):
            try:
                service = obter_serviço_drive()
                pesq = data_pesquisa.strip()
                if "/" in pesq:
                    try: pesq_limpa = f"{int(pesq.split('/')[0])}/{int(pesq.split('/')[1])}"
                    except: pesq_limpa = pesq
                else: pesq_limpa = pesq

                query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
                resultados = service.files().list(q=query, fields="files(id, name)").execute()
                mapa_arquivos = {arq['name']: arq['id'] for arq in resultados.get('files', [])}
                
                id_controle = mapa_arquivos['Arquivo_de_controle.xlsx']
                req_c = service.files().get_media(fileId=id_controle)
                bytes_c = io.BytesIO()
                baix_c = MediaIoBaseDownload(bytes_c, req_c)
                while not baix_c.next_chunk()[1]: pass
                bytes_c.seek(0)
                wb_leitura = openpyxl.load_workbook(bytes_c, data_only=True)
                
                ws_fase1 = wb_leitura['FASE 1']
                jogos_reais_do_dia = {}
                for r in range(3, 75):
                    val_data = ws_fase1[f'C{r}'].value
                    if val_data is None: continue
                    data_jogo = f"{val_data.day}/{val_data.month}" if isinstance(val_data, dt.datetime) else str(val_data).strip()
                    if "/" in data_jogo and not isinstance(val_data, dt.datetime):
                        try: data_jogo = f"{int(data_jogo.split('/')[0])}/{int(data_jogo.split('/')[1])}"
                        except: pass
                    
                    if data_jogo == pesq_limpa:
                        jogos_reais_do_dia[r] = {
                            'time_casa': str(ws_fase1[f'F{r}'].value).strip(),
                            'g_casa': ws_fase1[f'G{r}'].value,
                            'g_fora': ws_fase1[f'I{r}'].value,
                            'time_fora': str(ws_fase1[f'J{r}'].value).strip()
                        }

                if jogos_reais_do_dia:
                    st.markdown("### 🏟️ Resultados Oficiais")
                    cols_reais = st.columns(len(jogos_reais_do_dia))
                    for idx, (linha, jogo) in enumerate(jogos_reais_do_dia.items()):
                        gc = int(float(jogo['g_casa'])) if jogo['g_casa'] is not None else "-"
                        gf = int(float(jogo['g_fora'])) if jogo['g_fora'] is not None else "-"
                        with cols_reais[idx % len(cols_reais)]:
                            st.markdown(f"""
                            <div style="background:#1E293B; border:1px solid #334155; border-radius:8px; padding:12px; text-align:center; margin-bottom:15px;">
                                <div style="font-size: 16px; font-weight: bold; color: #F8FAFC;">
                                    {jogo['time_casa']} &nbsp;<span style="color:#60A5FA;">{gc} x {gf}</span>&nbsp; {jogo['time_fora']}
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                ws_tabela = wb_leitura['TABELA']
                participantes = []
                for r in range(2, 101):
                    nome = ws_tabela[f'A{r}'].value
                    if nome and str(nome).strip() != "" and str(nome).strip().lower() != "participante":
                        participantes.append(str(nome).strip())
                wb_leitura.close()
                
                dict_ranking = {p['nome']: p for p in st.session_state.ranking_processado} if st.session_state.ranking_processado else {}
                dados_painel = []
                melhor_pontuacao_dia = -1
                herois_do_dia = []

                for nome in participantes:
                    arq_palpite = f"{nome}.xlsx"
                    if arq_palpite in mapa_arquivos:
                        id_part = mapa_arquivos[arq_palpite]
                        palpites_usuario = carregar_palpites_em_cache(id_part)
                        
                        pontos_do_dia = 0
                        cards_html = ""
                        
                        for r, jogo_real in jogos_reais_do_dia.items():
                            palp_c, palp_f = palpites_usuario.get(r, (None, None))
                            
                            try: p_c_str = str(int(float(palp_c))) if palp_c is not None else "-"
                            except: p_c_str = "-"
                            try: p_f_str = str(int(float(palp_f))) if palp_f is not None else "-"
                            except: p_f_str = "-"
                            
                            cor_borda = "#334155" 
                            if jogo_real['g_casa'] is not None and jogo_real['g_fora'] is not None:
                                pts = calcular_pontos(jogo_real['g_casa'], jogo_real['g_fora'], palp_c, palp_f)
                                pontos_do_dia += pts
                                if pts == 7: cor_borda = "#166534" 
                                elif pts in [2, 3, 4]: cor_borda = "#2563EB" 
                                else: cor_borda = "#991B1B" 
                                
                            cards_html += f"""
                            <div style="background:#0F172A; border:1px solid {cor_borda}; border-radius:6px; padding:8px 12px; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;">
                               <span style="color:#CBD5E1; width:30%; text-align:right; font-size:14px;">{jogo_real['time_casa']}</span>
                               <span style="color:#F8FAFC; font-weight:bold; width:40%; text-align:center; font-size:15px;">{p_c_str} x {p_f_str}</span>
                               <span style="color:#CBD5E1; width:30%; text-align:left; font-size:14px;">{jogo_real['time_fora']}</span>
                            </div>
                            """
                        
                        info = dict_ranking.get(nome, {'pontos': '?', 'posicao': '-'})
                        pos_str = str(info['posicao']).replace("º Lugar", "").strip()
                        
                        emoji = "👤"
                        if pos_str == '1': emoji = "🥇"
                        elif pos_str == '2': emoji = "🥈"
                        elif pos_str == '3': emoji = "🥉"
                        
                        cabecalho = f"{emoji} {nome} | {info['pontos']} pts | {info['posicao']}"
                        
                        dados_painel.append({'cabecalho': cabecalho, 'cards_html': cards_html, 'pontos_dia': pontos_do_dia})
                        
                        if pontos_do_dia > melhor_pontuacao_dia:
                            melhor_pontuacao_dia = pontos_do_dia
                            herois_do_dia = [nome]
                        elif pontos_do_dia == melhor_pontuacao_dia and pontos_do_dia > 0:
                            herois_do_dia.append(nome)

                if jogos_reais_do_dia and melhor_pontuacao_dia > 0:
                    st.markdown("---")
                    st.markdown("### 🔥 Destaque da Rodada")
                    st.markdown(f"""
                    <div style="background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%); border-left: 4px solid #F59E0B; padding: 15px; border-radius: 6px; margin-bottom: 25px;">
                        <h4 style="color: #FBBF24; margin: 0 0 5px 0;">{", ".join(herois_do_dia)}</h4>
                        <p style="color: #F8FAFC; margin: 0; font-size: 14px;">Mito da rodada somando impressionantes <strong>{melhor_pontuacao_dia} pontos</strong> só nos jogos de hoje!</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("### 📋 Palpites da Galera")
                cols_grid = st.columns(3)
                for idx, painel in enumerate(dados_painel):
                    with cols_grid[idx % 3]:
                        with st.expander(painel['cabecalho']): st.markdown(painel['cards_html'], unsafe_allow_html=True)
                            
            except Exception as e: st.error(f"Ocorreu um erro na busca: {e}")
