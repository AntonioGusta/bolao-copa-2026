import os
import io
import streamlit as st
import openpyxl
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# CONFIGURAÇÃO DE ACESSO (COLE O ID DA SUA PASTA DO DRIVE AQUI)
# ==============================================================================
# Substitua o texto abaixo pelo ID real que você pegou na barra de endereços do Drive
ID_PASTA_DRIVE = "1YXz96rcj7IgthzYHK6uyei_KrC5q9cqb"

def obter_serviço_drive():
    # Conexão utilizando as chaves que vamos guardar nas configurações do Streamlit
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

# Configuração visual da página Web
st.set_page_config(page_title="Bolão Copa 2026 (THE)", page_icon="⚽", layout="centered")
st.title("🏆 Apuração do Bolão da Copa 2026 (THE)")
st.write("Clique no botão abaixo para processar os palpites e atualizar o ranking em tempo real.")

if st.button("🚀 Atualizar Classificação", type="primary"):
    try:
        service = obter_serviço_drive()
        
        # Buscar os arquivos dentro da pasta do Drive
        query = f"'{ID_PASTA_DRIVE}' in parents and trashed = false"
        resultados = service.files().list(q=query, fields="files(id, name)").execute()
        arquivos = resultados.get('files', [])
        mapa_arquivos = {arq['name']: arq['id'] for arq in arquivos}
        
        if 'Arquivo_de_controle.xlsx' not in mapa_arquivos:
            st.error("Erro: 'Arquivo_de_controle.xlsx' não encontrado na pasta do Drive.")
            st.stop()
            
        # Baixar o Arquivo de Controle para a memória do servidor
        id_controle = mapa_arquivos['Arquivo_de_controle.xlsx']
        requisicao = service.files().get_media(fileId=id_controle)
        bytes_controle = io.BytesIO()
        baixador = MediaIoBaseDownload(bytes_controle, requisicao)
        concluido = False
        while not concluido:
            _, concluido = baixador.next_chunk()
            
        bytes_controle.seek(0)
        wb_leitura = openpyxl.load_workbook(bytes_controle, data_only=True)
        
        ws_fase1_leitura = wb_leitura['FASE 1']
        resultados_reais = {}
        for r in range(3, 75):
            resultados_reais[r] = (ws_fase1_leitura[f'G{r}'].value, ws_fase1_leitura[f'I{r}'].value)
            
        ws_tabela_leitura = wb_leitura['TABELA']
        lista_ranking = []
        
        st.subheader("📦 Processamento de Planilhas:")
        caixa_logs = st.empty()
        texto_logs = ""
        
        # Varre a tabela buscando participantes a partir da linha 2
        for row_tabela in range(2, 101):
            celula_nome = ws_tabela_leitura[f'A{row_tabela}'].value
            if celula_nome is None or str(celula_nome).strip() == "" or str(celula_nome).strip().lower() == "participante":
                continue
                
            nome = str(celula_nome).strip()
            arquivo_palpite = f"{nome}.xlsx"
            
            texto_logs += f"Lendo {arquivo_palpite}...\n"
            caixa_logs.code(texto_logs)
            
            if arquivo_palpite not in mapa_arquivos:
                lista_ranking.append({'nome': nome, 'pontos': 0})
                texto_logs += f"  {nome} acumulou 0 pontos.\n........................................\n"
                caixa_logs.code(texto_logs)
                continue
                
            # Baixar o palpite do participante
            id_part = mapa_arquivos[arquivo_palpite]
            req_part = service.files().get_media(fileId=id_part)
            bytes_part = io.BytesIO()
            baixador_part = MediaIoBaseDownload(bytes_part, req_part)
            concluido_part = False
            while not concluido_part:
                _, concluido_part = baixador_part.next_chunk()
                
            bytes_part.seek(0)
            wb_part = openpyxl.load_workbook(bytes_part, data_only=True)
            ws_part_fase1 = wb_part['FASE 1']
            
            pontos_totais = 0
            for r in range(3, 75):
                g_m_real, g_v_real = resultados_reais[r]
                if g_m_real is None or g_v_real is None: continue
                pontos_totais += calcular_pontos(g_m_real, g_v_real, ws_part_fase1[f'G{r}'].value, ws_part_fase1[f'I{r}'].value)
                
            wb_part.close()
            texto_logs += f"  {nome} acumulou {pontos_totais} pontos.\n........................................\n"
            caixa_logs.code(texto_logs)
            lista_ranking.append({'nome': nome, 'pontos': points_totais if 'points_totais' in locals() else pontos_totais})
            
        wb_leitura.close()
        
        # Ordenação do ranking
        lista_ranking.sort(key=lambda x: x['pontos'], reverse=True)
        
        # Abrir o arquivo de controle original preservando fórmulas para gerar a saída
        bytes_controle.seek(0)
        wb_gravar = openpyxl.load_workbook(bytes_controle, data_only=False)
        if 'Dados Pessoais' in wb_gravar.sheetnames:
            wb_gravar.remove(wb_gravar['Dados Pessoais'])
            
        ws_tabela_gravar = wb_gravar['TABELA']
        for r in range(1, 101):
            ws_tabela_gravar[f'A{r}'].value = None
            ws_tabela_gravar[f'B{r}'].value = None
            ws_tabela_gravar[f'C{r}'].value = None
            
        ws_tabela_gravar['A1'].value = "Participante"
        ws_tabela_gravar['B1'].value = "Pontos"
        ws_tabela_gravar['C1'].value = "Posição"
        
        st.subheader("🏆 Classificação Atualizada:")
        
        for idx, participante in enumerate(lista_ranking):
            linha_atual = 2 + idx
            ws_tabela_gravar[f'A{linha_atual}'].value = participante['nome']
            ws_tabela_gravar[f'B{linha_atual}'].value = participante['pontos']
            ws_tabela_gravar[f'C{linha_atual}'].value = f"{idx+1}º Lugar"
            
            if idx == 0: emoji = "🥇 "
            elif idx == 1: emoji = "🥈 "
            elif idx == 2: emoji = "🥉 "
            else: emoji = "  "
                
            st.text(f"{emoji}{idx+1}º Lugar: {participante['nome'].ljust(16)} | Pontos: {participante['pontos']}")
            
        # Salvar o arquivo de resultados em memória temporária
        arquivo_saida_bytes = io.BytesIO()
        wb_gravar.save(arquivo_saida_bytes)
        arquivo_saida_bytes.seek(0)
        wb_gravar.close()
        
        # Enviar o arquivo pronto de volta para o Google Drive
        nome_saida = 'BOLÃO DA COPA DO MUNDO 2026 (THE).xlsx'
        mimetype_excel = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        media_upload = MediaIoBaseUpload(arquivo_saida_bytes, mimetype=mimetype_excel, resumable=True)
        
        if nome_saida in mapa_arquivos:
            service.files().update(fileId=mapa_arquivos[nome_saida], media_body=media_upload).execute()
        else:
            metadados = {'name': nome_saida, 'parents': [ID_PASTA_DRIVE]}
            service.files().create(body=metadados, media_body=media_upload).execute()
            
        st.success(f"Planilha '{nome_saida}' atualizada com sucesso no seu Google Drive!")
        
    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")
