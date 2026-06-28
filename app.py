def gerar_tabela_html(participantes, titulo, subtitulo, cor_destaque=None, posicoes_destaque=None):
    if posicoes_destaque is None: posicoes_destaque = []
    
    html = f"<div style='background: #0F172A; padding: 15px; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.35); text-align: center; margin-bottom: 20px; font-family: sans-serif;'>"
    html += f"<h4 style='margin-bottom: 0px; color: #F8FAFC;'>{titulo}</h4>"
    html += f"<p style='margin-top: 0px; font-size: 14px; font-style: italic; color: #CBD5E1;'>{subtitulo}</p>"
    
    html += "<table style='width: 100%; border-collapse: collapse; text-align: center; font-size: 13px; color: #F8FAFC; margin-top: 15px;'>"
    html += "<thead style='background-color: #111827; border-bottom: 2px solid #334155;'>"
    # --- NOVO CABEÇALHO ---
    html += "<tr><th style='padding: 8px;'>Pos</th><th style='padding: 8px;'>Var</th><th style='text-align: left; padding: 8px;'>Nome</th><th style='padding: 8px; color:#94A3B8;'>Fase Grupos</th><th style='padding: 8px; color:#94A3B8;'>Pré Oitavas</th><th style='padding: 8px; font-size: 15px;'>Total</th><th style='padding: 8px; color:#64748B;'>Dif</th></tr>"
    html += "</thead><tbody>"

    for p in participantes:
        pos_str = p['posicao'].replace("º Lugar", "").strip()
        is_highlight = (pos_str.isdigit() and int(pos_str) in posicoes_destaque)
        bg_color = cor_destaque if is_highlight else "#1E293B"
        
        html += f"<tr style='background-color: {bg_color}; border-bottom: 1px solid #334155;'>"
        html += f"<td style='padding: 6px; font-weight: bold;'>{p['posicao'].replace(' Lugar', '')}</td>"
        html += f"<td style='padding: 6px; font-size: 13px;'>{p.get('var_html', '➖')}</td>"
        html += f"<td style='text-align: left; padding: 6px;'>{p.get('nome_exibicao', p['nome'])}</td>"
        
        # --- NOVAS COLUNAS DE PONTOS ---
        html += f"<td style='padding: 6px; color:#CBD5E1;'>{p.get('pts_grupos', 0)}</td>"
        html += f"<td style='padding: 6px; color:#CBD5E1;'>{p.get('pts_pre_oitavas', 0)}</td>"
        html += f"<td style='padding: 6px; font-weight: bold; font-size: 15px; color:#F59E0B;'>{p['pontos']}</td>"
        
        html += f"<td style='padding: 6px; font-size: 12px; color:#64748B;'>{p.get('dif', '-')}</td>"
        html += "</tr>"

    html += "</tbody></table></div>"
    return html
