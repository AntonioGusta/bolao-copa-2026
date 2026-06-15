import os
import io
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
    if g_m_real is None or g_v_real is None or g_m_palpite is None or g_v_palpite is None:
        return 0
    try:
        g_m_real, g_v_real = int(g_m_real), int(g_v_real)
        g_m_palpite, g_v_palpite = int(g_m_palpite), int(g_v_palpite)
    except (ValueError, TypeError):
        return 0
    res_real = 1 if g_m_real > g_v_real else (-1 if g_m_real < g_v_real else 0)
    res_palpite = 1 if g_m_palpite > g_v_palpite else (-1 if g_m_palpite < g_v_palpite else 0)
    if res_real != res_palpite: return 0
    if g_m_real == g_m_palpite and g_v_real == g_v_palpite: return 7
    if res_real == 0: return 3
    if g_m_real == g_m_palpite or g_v_real == g_v_palpite: return 4
    return 2

def gerar_tabela_html(participantes, titulo, subtitulo, cor_destaque=None, posicoes_destaque=None):
    """Gera um bloco HTML estilizado para as tabelas do bolão."""
    if posicoes_destaque is None:
        posicoes_destaque = []
        
    html = f"""
    <div style="text-align: center; margin-bottom: 20px; font-family: sans-serif;">
        <h4 style="margin-bottom: 0px; color: #1f1f1f;">{titulo}</h4>
        <p style="margin-top: 0px; font-size: 14px; font-style: italic; color: #555;">{subtitulo}</p>
        <table style="width: 100%; border-collapse: collapse; text-align: center; font-size: 14px;">
            <thead>
                <tr style="background-color: #f0f2f6; border-bottom: 2px solid #ccc;">
                    <th style="padding: 8px;">Pos</th>
                    <th style="text-align: left; padding: 8px;">Nome</th>
                    <th style="padding: 8px;">Pts</th>
                </tr>
            </thead>
            <tbody>
    """
    for p in participantes:
        pos_str = p['posicao'].replace("º Lugar", "").strip()
        bg_color = cor_destaque if (pos_str.isdigit() and int(pos_str) in posicoes_destaque) else "transparent"
        
        html += f"""
                <tr style="background-color: {bg_color}; border-bottom: 1px solid #eee;">
                    <td style="padding: 6px; font-weight: bold;">{p['posicao'].replace(' Lugar', '')}</td>
                    <td style="text-align: left; padding: 6px;">{p['nome']}</td>
                    <td style="padding: 6px; font-weight: bold;">{p['pontos']}</td>
                </tr>
        """
    html += """
            </tbody>
        </table>
    </div>
    """
    return html

# ==============================================================================
# CONFIGURAÇÃO VISUAL DA PÁGINA WEB E GESTÃO DE ESTADO
# ==============================================================================
st.set_page_config(page_title="Bolão Copa 2026 (THE)", page_icon="⚽", layout="wide")

# Inicializa o cache da tabela no session_state para ela não sumir ao mudar de aba
if 'ranking_processado' not in st.session_state:
    st.session_state.ranking_processado = []

st.title("🏆 Apuração do Bolão da Copa 2026 (THE) - V2")
st.write("Clique no botão abaixo para processar os palpites no Drive e atualizar o ranking em tempo real.")

# Botão principal processa os dados e salva na memória
if st.button("🚀 Atualizar Classificação", type="primary"):
    with st.spinner("Lendo planilhas no Google Drive... Isso pode levar alguns segundos."):
        try:
            service = obter_serviço_drive()
            query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
            resultados = service.files().list(q=query, fields="files(id, name)").execute()
            mapa_arquivos = {arq['name']: arq['id'] for arq in resultados.get('files', [])}
            
            if 'Arquivo_de_controle.xlsx' not in mapa_arquivos:
                st.error("Erro: 'Arquivo_de_controle.xlsx' não encontrado na pasta do Drive.")
                st.stop()
                
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
            
            # Leitura silenciosa dos participantes
            for row_tabela in range(2, 101):
                celula_nome = ws_tabela_leitura[f'A{row_tabela}'].value
                if celula_nome is None or str(celula_nome).strip() == "" or str(celula_nome).strip().lower() == "participante":
                    continue
                    
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
            
            # Lógica de posições com empates
            posicao_atual = 1
            pontos_anteriores = None
            
            bytes_controle.seek(0)
            wb_gravar = openpyxl.load_workbook(bytes_controle, data_only=False)
            if 'Dados Pessoais' in wb_gravar.sheetnames: wb_gravar.remove(wb_gravar['Dados Pessoais'])
            ws_tabela_gravar = wb_gravar['TABELA']
            
            for r in range(1, 101):
                ws_tabela_gravar[f'A{r}'].value, ws_tabela_gravar[f'B{r}'].value, ws_tabela_gravar[f'C{r}'].value = None, None, None
                
            ws_tabela_gravar['A1'].value, ws_tabela_gravar['B1'].value, ws_tabela_gravar['C1'].value = "Participante", "Pontos", "Posição"
            
            for idx, participante in enumerate(lista_ranking):
                linha_atual = 2 + idx
                if participante['pontos'] != pontos_anteriores:
                    posicao_atual = idx + 1
                pontos_anteriores = participante['pontos']
                
                # Registra a posição em string dentro do dicionário
                participante['posicao'] = f"{posicao_atual}º Lugar"
                
                ws_tabela_gravar[f'A{linha_atual}'].value = participante['nome']
                ws_tabela_gravar[f'B{linha_atual}'].value = participante['pontos']
                ws_tabela_gravar[f'C{linha_atual}'].value = participante['posicao']
                
            # Salva o ranking atualizado no estado da sessão
            st.session_state.ranking_processado = lista_ranking.copy()
            
            arquivo_saida_bytes = io.BytesIO()
            wb_gravar.save(arquivo_saida_bytes)
            arquivo_saida_bytes.seek(0)
            wb_gravar.close()
            
            nome_saida = 'BOLÃO DA COPA DO MUNDO 2026 (THE).xlsx'
            mimetype_excel = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            media_upload = MediaIoBaseUpload(arquivo_saida_bytes, mimetype=mimetype_excel, resumable=True)
            
            if nome_saida in mapa_arquivos:
                service.files().update(fileId=mapa_arquivos[nome_saida], media_body=media_upload).execute()
            else:
                metadados = {'name': nome_saida, 'parents': [ID_PASTA_DRIVE]}
                service.files().create(body=metadados, media_body=media_upload).execute()
                
            st.success("Tabela atualizada com sucesso!")
            
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento: {e}")

st.divider()

# ==============================================================================
# NAVEGAÇÃO EM ABAS (UI)
# ==============================================================================
aba_ranking, aba_palpites = st.tabs(["🏆 Classificação e Resenha", "👀 Espiar Palpites da Rodada"])

# --- ABA 1: RANKING ESTILIZADO ---
with aba_ranking:
    if not st.session_state.ranking_processado:
        st.info("👆 Clique no botão azul lá em cima para buscar os dados no Google Drive e gerar a classificação atualizada.")
    else:
        dados = st.session_state.ranking_processado
        
        # Fatiamento estratégico das categorias
        profissionais = dados[:5]
        amadores = dados[5:15]
        # Peladeiros engloba todo o resto, exceto os últimos 3 (para garantir que não falte ninguém na tela)
        peladeiros = dados[15:-3] 
        lanterna = dados[-3:]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            html_prof = gerar_tabela_html(
                profissionais, 
                "Profissionais", "- Elite do Pitaco -", 
                cor_destaque="#c8e6c9", # Verde claro
                posicoes_destaque=[1, 2, 3, 4, 5]
            )
            st.markdown(html_prof, unsafe_allow_html=True)
            
        with col2:
            html_amadores = gerar_tabela_html(
                amadores, 
                "Amadores", "- Os que ainda sonham -"
            )
            st.markdown(html_amadores, unsafe_allow_html=True)
            
        with col3:
            html_peladeiros = gerar_tabela_html(
                peladeiros, 
                "Peladeiros", "- Especialistas em Errar -"
            )
            st.markdown(html_peladeiros, unsafe_allow_html=True)
            
        st.write("---")
        col_vazia1, col_lanterna, col_vazia2 = st.columns([1, 2, 1])
        with col_lanterna:
            posicoes_lanterna = [int(p['posicao'].replace("º Lugar", "").strip()) for p in lanterna]
            html_lanterna = gerar_tabela_html(
                lanterna, 
                "Prêmio Espírito Coletivo", "- Bastava Apostar ao Contrário -",
                cor_destaque="#ffcdd2", # Vermelho claro
                posicoes_destaque=posicoes_lanterna
            )
            st.markdown(html_lanterna, unsafe_allow_html=True)

# --- ABA 2: PALPITES DO DIA ---
with aba_palpites:
    st.subheader("👀 Espiar Palpites da Rodada")
    
    fuso_br = dt.timezone(dt.timedelta(hours=-3))
    hoje = dt.datetime.now(fuso_br)
    data_padrao = f"{hoje.day}/{hoje.month}"
    
    data_pesquisa = st.text_input("📅 Digite a data dos jogos que deseja ver (Ex: 12/6):", value=data_padrao)
    
    if st.button("🔍 Ver Palpites do Dia", type="secondary"):
        with st.spinner("Buscando os palpites..."):
            try:
                service = obter_serviço_drive()
                pesq = data_pesquisa.strip()
                if "/" in pesq:
                    try: pesq_limpa = f"{int(pesq.split('/')[0])}/{int(pesq.split('/')[1])}"
                    except: pesq_limpa = pesq
                else:
                    pesq_limpa = pesq

                query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
                resultados = service.files().list(q=query, fields="files(id, name)").execute()
                mapa_arquivos = {arq['name']: arq['id'] for arq in resultados.get('files', [])}
                
                if 'Arquivo_de_controle.xlsx' not in mapa_arquivos:
                    st.error("Erro: 'Arquivo_de_controle.xlsx' não encontrado.")
                    st.stop()
                    
                id_controle = mapa_arquivos['Arquivo_de_controle.xlsx']
                req_c = service.files().get_media(fileId=id_controle)
                bytes_c = io.BytesIO()
                baix_c = MediaIoBaseDownload(bytes_c, req_c)
                concluido = False
                while not concluido: _, concluido = baix_c.next_chunk()
                    
                bytes_c.seek(0)
                wb_leitura = openpyxl.load_workbook(bytes_c, data_only=True)
                ws_tabela = wb_leitura['TABELA']
                
                participantes = []
                for r in range(2, 101):
                    nome = ws_tabela[f'A{r}'].value
                    if nome and str(nome).strip() != "" and str(nome).strip().lower() != "participante":
                        participantes.append(str(nome).strip())
                wb_leitura.close()
                
                st.write(f"**Buscando os palpites de todos para os jogos do dia: {pesq_limpa}...**")
                
                for nome in participantes:
                    arq_palpite = f"{nome}.xlsx"
                    if arq_palpite in mapa_arquivos:
                        id_p = mapa_arquivos[arq_palpite]
                        req_p = service.files().get_media(fileId=id_p)
                        bytes_p = io.BytesIO()
                        baix_p = MediaIoBaseDownload(bytes_p, req_p)
                        conc_p = False
                        while not conc_p: _, conc_p = baix_p.next_chunk()
                        
                        bytes_p.seek(0)
                        wb_p = openpyxl.load_workbook(bytes_p, data_only=True)
                        ws_p = wb_p['FASE 1']
                        
                        palpites_do_dia = []
                        COLUNA_DATA = 'C'  
                        
                        for r in range(3, 75):
                            val_data = ws_p[f'{COLUNA_DATA}{r}'].value
                            if val_data is None: continue
                            
                            if isinstance(val_data, dt.datetime):
                                data_jogo = f"{val_data.day}/{val_data.month}"
                            else:
                                txt = str(val_data).strip()
                                if "/" in txt:
                                    try: data_jogo = f"{int(txt.split('/')[0])}/{int(txt.split('/')[1])}"
                                    except: data_jogo = txt
                                else:
                                    data_jogo = txt
                            
                            if data_jogo == pesq_limpa:
                                time_casa = str(ws_p[f'F{r}'].value).strip()
                                gols_casa = ws_p[f'G{r}'].value
                                gols_fora = ws_p[f'I{r}'].value
                                time_fora = str(ws_p[f'J{r}'].value).strip()
                                
                                try: g_casa_str = str(int(float(gols_casa))) if gols_casa is not None and str(gols_casa).strip() != "" else "-"
                                except ValueError: g_casa_str = "-"
                                    
                                try: g_fora_str = str(int(float(gols_fora))) if gols_fora is not None and str(gols_fora).strip() != "" else "-"
                                except ValueError: g_fora_str = "-"
                                
                                palpites_do_dia.append(f"⚽ {time_casa} **{g_casa_str} x {g_fora_str}** {time_fora}")
                        
                        wb_p.close()
                        
                        if palpites_do_dia:
                            with st.expander(f"👤 Palpites de {nome}"):
                                for p in palpites_do_dia:
                                    st.markdown(p)
                                    
            except Exception as e:
                st.error(f"Ocorreu um erro na busca: {e}")
