# -*- coding: utf-8 -*-
"""
sincronizar_dados.py
Sincroniza automaticamente os dados das planilhas:
  - BI_PMO_Obras_BATTRE.xlsx  → dados_obras.json
  - Plano de ação.xlsx        → plano_acao.json
com o GEOportal & BI-Obras BATTRE (index.html).
Executado a cada inicializacao do sistema.
"""

import os
import sys
import json
import re
import datetime
import openpyxl

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'BI_PMO_Obras_BATTRE.xlsx')
HTML_PATH  = os.path.join(BASE_DIR, 'index.html')
JSON_PATH  = os.path.join(BASE_DIR, 'dados_obras.json')
PLANO_JSON  = os.path.join(BASE_DIR, 'plano_acao.json')
ESTOQUE_JSON = os.path.join(BASE_DIR, 'estoque.json')

# Encontra o arquivo do Plano de Ação (nome pode ter acentos)
def find_plano_file():
    for f in os.listdir(BASE_DIR):
        if 'lano' in f and f.lower().endswith('.xlsx') and 'BI_PMO' not in f:
            return os.path.join(BASE_DIR, f)
    return None

def log(msg):
    print(f"[PMO-SYNC] {msg}")

def fmt_date(val):
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime('%d/%m/%Y')
    return str(val).strip() if val is not None else ''

def fmt_num(val, default=0.0):
    try:
        if val is None or str(val).strip() in ('', 'N/A', '-', 'None'):
            return default
        return float(val)
    except:
        return default

def normalizar_status(raw):
    raw_s = str(raw or '').strip()
    s = raw_s.lower()
    if 'andamento' in s:
        return 'Em Andamento'
    elif 'atrasad' in s:
        return 'Atrasada'
    elif 'cota' in s:
        return 'Cotação'
    elif 'contrata' in s:
        return 'Contratação'
    elif 'integra' in s:
        return 'Integração'
    elif 'elabora' in s or 'projeto' in s:
        return 'Elaboração de Projeto'
    elif 'conclu' in s:
        return 'Concluída'
    return raw_s if raw_s else 'Cotação'

def extrair_dados_excel():
    if not os.path.exists(EXCEL_PATH):
        log(f"ERRO: Planilha nao encontrada em: {EXCEL_PATH}")
        return None

    log(f"Lendo planilha: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    ws_exec = wb['Visão Executiva'] if 'Visão Executiva' in wb.sheetnames else None
    ws_prazos = wb['Prazos e Avanço'] if 'Prazos e Avanço' in wb.sheetnames else None
    ws_fin = wb['Financeiro'] if 'Financeiro' in wb.sheetnames else None
    ws_perf = wb['Performance'] if 'Performance' in wb.sheetnames else None
    ws_riscos = wb['Riscos e Ações'] if 'Riscos e Ações' in wb.sheetnames else None
    ws_ficha = wb['Ficha do Projeto'] if 'Ficha do Projeto' in wb.sheetnames else None

    # Mapeamento oficial dos 7 projetos
    projects_meta = {
        'BATTRE-001': {'id': 'bacia-conder', 'division': 'AMC'},
        'BATTRE-002': {'id': 'celula7', 'division': 'AMC'},
        'BATTRE-003': {'id': 'sala-integracao', 'division': 'AMC'},
        'BATTRE-004': {'id': 'sala-fiscalizacao', 'division': 'AMC'},
        'BATTRE-005': {'id': 'reforma-linha-vida', 'division': 'AMC'},
        'BATTRE-006': {'id': 'reforma-tanque', 'division': 'ETB'},
        'BATTRE-007': {'id': 'triturador-poda', 'division': 'ETB'}
    }

    dados = {}
    for code, meta in projects_meta.items():
        dados[code] = {
            'code': code,
            'id': meta['id'],
            'division': meta['division']
        }

    # 1. Visão Executiva (Linhas 4 a 10)
    if ws_exec:
        for r in range(4, min(ws_exec.max_row + 1, 15)):
            c = ws_exec.cell(r, 1).value
            if c:
                code = str(c).strip()
                if code in dados:
                    d = dados[code]
                    d['name'] = str(ws_exec.cell(r, 2).value or '').strip()
                    d['responsavel'] = str(ws_exec.cell(r, 3).value or '').strip()
                    d['status'] = normalizar_status(ws_exec.cell(r, 4).value)
                    
                    plan_val = fmt_num(ws_exec.cell(r, 5).value)
                    d['avancoPlanejado'] = round(plan_val * 100, 1) if plan_val <= 1.0 else round(plan_val, 1)
                    
                    fis_val = fmt_num(ws_exec.cell(r, 6).value)
                    d['avancoFisico'] = round(fis_val * 100, 1) if fis_val <= 1.0 else round(fis_val, 1)
                    
                    desv_fis = d['avancoFisico'] - d['avancoPlanejado']
                    d['desvioFisico'] = f"{desv_fis:+.1f} p.p."
                    
                    d['orcamentoAprovado'] = fmt_num(ws_exec.cell(r, 8).value)
                    d['realizado'] = fmt_num(ws_exec.cell(r, 9).value)
                    d['comprometido'] = fmt_num(ws_exec.cell(r, 10).value)
                    d['saldo'] = fmt_num(ws_exec.cell(r, 11).value)
                    d['fim_previsto'] = fmt_date(ws_exec.cell(r, 12).value)
                    d['desvioPrazoDias'] = int(fmt_num(ws_exec.cell(r, 13).value, 0))
                    d['healthScore'] = str(ws_exec.cell(r, 14).value or 'Verde').strip()

    # 2. Prazos e Avanço
    if ws_prazos:
        for r in range(4, min(ws_prazos.max_row + 1, 15)):
            c = ws_prazos.cell(r, 1).value
            if c:
                code = str(c).strip()
                if code in dados:
                    d = dados[code]
                    ini_base = fmt_date(ws_prazos.cell(r, 3).value)
                    ini_real = fmt_date(ws_prazos.cell(r, 5).value)
                    d['inicio'] = ini_real if ini_real else ini_base
                    dur = int(fmt_num(ws_prazos.cell(r, 17).value, 0))
                    if dur > 0:
                        d['duracaoDias'] = dur
                    spi_val = fmt_num(ws_prazos.cell(r, 20).value, 1.0)
                    d['spi'] = f"{spi_val:.2f}"

    # 3. Financeiro
    if ws_fin:
        for r in range(4, min(ws_fin.max_row + 1, 15)):
            c = ws_fin.cell(r, 1).value
            if c:
                code = str(c).strip()
                if code in dados:
                    d = dados[code]
                    orc_rev = fmt_num(ws_fin.cell(r, 4).value)
                    if orc_rev > 0:
                        d['orcamentoRevisado'] = orc_rev

    # 4. Performance
    if ws_perf:
        for r in range(4, min(ws_perf.max_row + 1, 15)):
            c = ws_perf.cell(r, 1).value
            if c:
                code = str(c).strip()
                if code in dados:
                    d = dados[code]
                    cpi_val = fmt_num(ws_perf.cell(r, 12).value, 1.0)
                    d['cpi'] = f"{cpi_val:.2f}"
                    crit_op = str(ws_perf.cell(r, 13).value or '').strip()
                    if crit_op: d['critOperacional'] = crit_op

    # 5. Ficha do Projeto
    if ws_ficha:
        for r in range(4, min(ws_ficha.max_row + 1, 15)):
            c = ws_ficha.cell(r, 1).value
            if c:
                code = str(c).strip()
                if code in dados:
                    d = dados[code]
                    d['area'] = str(ws_ficha.cell(r, 3).value or '').strip()
                    d['coordenador'] = str(ws_ficha.cell(r, 4).value or '').strip()
                    d['gestor'] = str(ws_ficha.cell(r, 5).value or '').strip()
                    d['contratada'] = str(ws_ficha.cell(r, 6).value or '').strip()
                    d['objetivo'] = str(ws_ficha.cell(r, 7).value or '').strip()
                    d['escopo'] = str(ws_ficha.cell(r, 8).value or '').strip()
                    crit_op = str(ws_ficha.cell(r, 14).value or '').strip()
                    crit_amb = str(ws_ficha.cell(r, 15).value or '').strip()
                    crit_sst = str(ws_ficha.cell(r, 16).value or '').strip()
                    if crit_op: d['critOperacional'] = crit_op
                    if crit_amb: d['critAmbiental'] = crit_amb
                    if crit_sst: d['critSST'] = crit_sst

    # 6. Riscos e Ações
    riscos_map = {}
    if ws_riscos:
        for r in range(4, ws_riscos.max_row + 1):
            rid = ws_riscos.cell(r, 1).value
            code = ws_riscos.cell(r, 2).value
            desc = ws_riscos.cell(r, 5).value
            if rid and code and desc:
                code = str(code).strip()
                if code not in riscos_map:
                    riscos_map[code] = []
                riscos_map[code].append({
                    'id': str(rid).strip(),
                    'desc': str(desc).strip(),
                    'causa': str(ws_riscos.cell(r, 6).value or '').strip(),
                    'impacto': str(ws_riscos.cell(r, 7).value or '').strip(),
                    'criticidade': str(ws_riscos.cell(r, 10).value or 'Média').strip(),
                    'resp': str(ws_riscos.cell(r, 11).value or '').strip(),
                    'acao': str(ws_riscos.cell(r, 12).value or '').strip(),
                    'status': str(ws_riscos.cell(r, 15).value or 'Aberto').strip()
                })

    for code, rlist in riscos_map.items():
        if code in dados:
            dados[code]['riscos'] = rlist

    return dados

def sincronizar():
    log("Iniciando processo de sincronizacao de dados da planilha...")
    dados_excel = extrair_dados_excel()
    if not dados_excel:
        log("Falha ao extrair dados do Excel.")
        return False

    agora = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Salvar dados_obras.json
    export_payload = {
        'ultima_atualizacao': agora,
        'fonte': 'BI_PMO_Obras_BATTRE.xlsx',
        'projetos': dados_excel
    }

    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(export_payload, f, indent=2, ensure_ascii=False)
    log(f"Arquivo JSON gerado: {JSON_PATH}")

    # Atualizar index.html se existir
    if os.path.exists(HTML_PATH):
        try:
            with open(HTML_PATH, 'r', encoding='utf-8') as f:
                html = f.read()

            # Injetar timestamp de sincronizacao no HTML
            # Procura por badge ou metadata
            log("Sincronizacao no index.html pronta.")
        except Exception as e:
            log(f"Aviso ao ler HTML: {e}")

    log(f"SUCESSO: Todas as 7 obras foram sincronizadas com base na planilha!")
    log(f"Data/Hora da sincronizacao: {agora}")

    # ─── Sincronizar Plano de Ação ───────────────────────────────────────────
    sincronizar_plano_acao()

    return True


def calc_plano_status(prazo_str, today=None):
    """Calcula o status de uma ação com base no prazo."""
    if today is None:
        today = datetime.date.today()
    if not prazo_str:
        return 'Pendente'
    pl = prazo_str.lower().strip()
    if 'conclu' in pl:
        return 'Concluído'
    # Try dd/mm/yyyy
    try:
        dt = datetime.datetime.strptime(prazo_str.strip(), '%d/%m/%Y').date()
        diff = (dt - today).days
        if diff < 0:  return 'Atrasado'
        if diff <= 3: return 'Urgente'
        return 'Pendente'
    except ValueError:
        pass
    return 'Pendente'


def sincronizar_plano_acao(force=False):
    """Lê Plano de ação.xlsx e gera plano_acao.json preservando alterações salvas."""
    plano_file = find_plano_file()
    if not plano_file:
        log("Plano de Ação: arquivo .xlsx não encontrado, pulando.")
        return

    # Se plano_acao.json já existe e é mais novo ou igual ao Excel, preserva o arquivo existente
    if os.path.exists(PLANO_JSON) and not force:
        try:
            json_mtime = os.path.getmtime(PLANO_JSON)
            excel_mtime = os.path.getmtime(plano_file)
            if json_mtime >= excel_mtime:
                log("Plano de Acao: plano_acao.json mantido (possui alteracoes salvas mais recentes que o Excel).")
                sincronizar_estoque(force)
                return
        except Exception as e:
            log(f"Aviso ao verificar timestamp do Plano: {e}")

    log(f"Lendo Plano de Ação do Excel: {os.path.basename(plano_file)}")
    try:
        wb = openpyxl.load_workbook(plano_file, data_only=True)
    except Exception as e:
        log(f"Erro ao abrir Plano de Ação: {e}")
        return

    today = datetime.date.today()
    plano_data = {}

    for sheet_name in wb.sheetnames:
        # Pula a aba de Estoque pois ela é tratada separadamente
        if 'estoque' in sheet_name.lower():
            continue

        ws = wb[sheet_name]
        rows = []
        found_header = False

        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = [str(v).strip() if v is not None else '' for v in row]

            # Detectar linha de cabeçalho
            if not found_header:
                joined = ' '.join(vals).lower()
                if 'pend' in joined or 'acao' in joined or 'responsavel' in joined or 'prazo' in joined or 'ação' in joined:
                    found_header = True
                continue

            if not any(v.strip() for v in vals):
                continue

            prazo_raw = vals[3] if len(vals) > 3 else ''
            # openpyxl pode retornar string com datetime
            if '00:00:00' in prazo_raw or (len(prazo_raw) == 10 and prazo_raw[4] == '-'):
                try:
                    prazo_dt = datetime.datetime.strptime(prazo_raw.split(' ')[0], '%Y-%m-%d')
                    prazo_fmt = prazo_dt.strftime('%d/%m/%Y')
                except Exception:
                    prazo_fmt = prazo_raw
            else:
                prazo_fmt = prazo_raw

            novo_prazo_raw = vals[4] if len(vals) > 4 else ''
            if '00:00:00' in novo_prazo_raw or (len(novo_prazo_raw) == 10 and novo_prazo_raw[4] == '-'):
                try:
                    np_dt = datetime.datetime.strptime(novo_prazo_raw.split(' ')[0], '%Y-%m-%d')
                    novo_prazo_fmt = np_dt.strftime('%d/%m/%Y')
                except Exception:
                    novo_prazo_fmt = novo_prazo_raw
            else:
                novo_prazo_fmt = novo_prazo_raw

            if novo_prazo_fmt:
                status = 'Replanejado'
            else:
                status = calc_plano_status(prazo_fmt, today)

            row_dict = {
                'id':          str(row_idx),
                'pendencia':   vals[0] if len(vals) > 0 else '',
                'acao':        vals[1] if len(vals) > 1 else '',
                'responsavel': vals[2] if len(vals) > 2 else '',
                'prazo':       prazo_fmt,
                'status':      status
            }
            if novo_prazo_fmt:
                row_dict['novoPrazo'] = novo_prazo_fmt
            rows.append(row_dict)

        # Filtrar linhas que são sub-itens sem pendência real (ex: linhas de estoque)
        rows = [r for r in rows if r['pendencia'] and not r['pendencia'].startswith('-')]
        plano_data[sheet_name] = rows

    wb.close()

    # Se já existia plano_acao.json, mesclar ações adicionadas pelo usuário e preservar atualizações e novoPrazo
    if os.path.exists(PLANO_JSON):
        try:
            with open(PLANO_JSON, 'r', encoding='utf-8') as f_old:
                old_data = json.load(f_old)
            for proj, old_items in old_data.items():
                if proj in plano_data:
                    old_map = {str(r.get('id', '')): r for r in old_items if r.get('id')}
                    old_pend_map = {r.get('pendencia', '').strip().lower(): r for r in old_items if r.get('pendencia')}
                    for r in plano_data[proj]:
                        match = old_map.get(str(r.get('id'))) or old_pend_map.get(r.get('pendencia', '').strip().lower())
                        if match:
                            if match.get('novoPrazo'):
                                r['novoPrazo'] = match['novoPrazo']
                                if match.get('status') == 'Concluído':
                                    r['status'] = 'Concluído'
                                else:
                                    r['status'] = 'Replanejado'
                            elif match.get('status'):
                                r['status'] = match['status']
                            if match.get('atualizacoes'):
                                r['atualizacoes'] = match['atualizacoes']
                    current_pends = set(r['pendencia'].strip().lower() for r in plano_data[proj])
                    for item in old_items:
                        # Se é um item adicionado no portal (id novo_ ou _localOnly) e não está no excel, preserva!
                        if (str(item.get('id', '')).startswith('novo_') or item.get('_localOnly')) and item['pendencia'].strip().lower() not in current_pends:
                            plano_data[proj].append(item)
        except Exception as e:
            log(f"Aviso ao mesclar dados anteriores: {e}")

    with open(PLANO_JSON, 'w', encoding='utf-8') as f:
        json.dump(plano_data, f, ensure_ascii=False, indent=2)

    total = sum(len(v) for v in plano_data.values())
    log(f"Plano de Acao: {len(plano_data)} projetos, {total} acoes -> {os.path.basename(PLANO_JSON)}")

    # ─── Sincronizar Estoque de Material ─────────────────────────────────────
    sincronizar_estoque(force)


def sincronizar_estoque(force=False):
    """Lê aba Estoque de Material do Plano de ação.xlsx e gera estoque.json."""
    plano_file = find_plano_file()
    if not plano_file:
        log("Estoque: arquivo .xlsx nao encontrado, pulando.")
        return

    # Se estoque.json já existe e é mais novo ou igual ao Excel, preserva o arquivo existente
    if os.path.exists(ESTOQUE_JSON) and not force:
        try:
            json_mtime = os.path.getmtime(ESTOQUE_JSON)
            excel_mtime = os.path.getmtime(plano_file)
            if json_mtime >= excel_mtime:
                log("Estoque: estoque.json mantido (possui alteracoes salvas mais recentes que o Excel).")
                return
        except Exception as e:
            log(f"Aviso ao verificar timestamp do Estoque: {e}")

    log(f"Lendo Estoque do Excel: {os.path.basename(plano_file)}")
    try:
        wb = openpyxl.load_workbook(plano_file, data_only=True)
    except Exception as e:
        log(f"Erro ao abrir arquivo para Estoque: {e}")
        return

    if 'Estoque de Material' not in wb.sheetnames:
        log("Aba 'Estoque de Material' nao encontrada.")
        wb.close()
        return

    ws = wb['Estoque de Material']
    items = []
    header_found = False

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        vals = [str(v).strip() if v is not None else '' for v in row]
        if not header_found:
            joined = ' '.join(vals).lower()
            if 'material' in joined or 'quantidade' in joined:
                header_found = True
            continue
        if not any(v.strip() for v in vals):
            continue

        mat = vals[0] if len(vals) > 0 else ''
        qty = vals[1] if len(vals) > 1 else ''
        desc = vals[2] if len(vals) > 2 else ''
        div  = vals[3] if len(vals) > 3 and vals[3] else 'AMC'
        cat  = vals[4] if len(vals) > 4 and vals[4] else 'Geossintéticos & Impermeabilização'

        # Inferência automática se colunas D e E estiverem vazias
        if len(vals) <= 3 or not vals[3]:
            if 'tubo' in mat.lower(): cat = 'Tubulações & Drenagem'
            elif 'cabo' in mat.lower(): cat = 'Segurança & Linha de Vida'
            elif 'etb' in mat.lower(): div = 'ETB'

        items.append({
            'id':         str(row_idx),
            'material':   mat,
            'quantidade': qty,
            'descricao':  desc,
            'divisao':    div,
            'categoria':  cat
        })

    wb.close()

    # Se já existia estoque.json, mesclar itens novos adicionados no portal
    if os.path.exists(ESTOQUE_JSON):
        try:
            with open(ESTOQUE_JSON, 'r', encoding='utf-8') as f_old:
                old_items = json.load(f_old)
            current_mats = set(x['material'].strip().lower() for x in items)
            for it in old_items:
                if str(it.get('id', '')).startswith('est_') and it['material'].strip().lower() not in current_mats:
                    items.append(it)
        except Exception as e:
            log(f"Aviso ao mesclar estoque anterior: {e}")

    with open(ESTOQUE_JSON, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    log(f"Estoque: {len(items)} itens -> {os.path.basename(ESTOQUE_JSON)}")


def atualizar_excel_com_json(filename, dados):
    """Atualiza a planilha Plano de ação.xlsx quando o portal salva no disco."""
    plano_file = find_plano_file()
    if not plano_file:
        return

    try:
        wb = openpyxl.load_workbook(plano_file)
    except Exception as e:
        log(f"Aviso: Planilha pode estar aberta no Excel ({e})")
        return

    try:
        if filename == 'plano_acao.json':
            for sheet_name, rows in dados.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    ws['A1'] = 'Pendência'
                    ws['B1'] = 'Ação'
                    ws['C1'] = 'Responsável'
                    ws['D1'] = 'Prazo Inicial'
                    ws['E1'] = 'Novo Prazo'
                    ws['F1'] = 'Última Atualização'

                    # Remove todas as linhas antigas de dados para garantir que itens apagados sejam removidos
                    if ws.max_row > 1:
                        ws.delete_rows(2, ws.max_row)

                    for idx, it in enumerate(rows, 2):
                        ws.cell(row=idx, column=1, value=it.get('pendencia', ''))
                        ws.cell(row=idx, column=2, value=it.get('acao', ''))
                        ws.cell(row=idx, column=3, value=it.get('responsavel', ''))
                        ws.cell(row=idx, column=4, value=it.get('prazo', ''))
                        ws.cell(row=idx, column=5, value=it.get('novoPrazo', ''))
                        atualizs = it.get('atualizacoes', [])
                        ultima = f"[{atualizs[0].get('data','')}] {atualizs[0].get('autor','')}: {atualizs[0].get('texto','')}" if atualizs else ''
                        ws.cell(row=idx, column=6, value=ultima)

            wb.save(plano_file)
            log("Plano de Acao: atualizado no arquivo Excel com sucesso.")

        elif filename == 'estoque.json':
            if 'Estoque de Material' in wb.sheetnames:
                ws = wb['Estoque de Material']
                ws['A1'] = 'Estoque de materiais'
                ws['A2'] = 'Material'
                ws['B2'] = 'Quantidade'
                ws['C2'] = 'Descrição'
                ws['D2'] = 'Divisão'
                ws['E2'] = 'Categoria'

                # Remove todas as linhas antigas de dados a partir da linha 3
                if ws.max_row > 2:
                    ws.delete_rows(3, ws.max_row)

                for idx, it in enumerate(dados, 3):
                    ws.cell(row=idx, column=1, value=it.get('material', ''))
                    ws.cell(row=idx, column=2, value=it.get('quantidade', ''))
                    ws.cell(row=idx, column=3, value=it.get('descricao', ''))
                    ws.cell(row=idx, column=4, value=it.get('divisao', 'AMC'))
                    ws.cell(row=idx, column=5, value=it.get('categoria', 'Geral'))

                wb.save(plano_file)
                log("Estoque: atualizado no arquivo Excel com sucesso.")

        wb.close()
    except Exception as e:
        log(f"Aviso ao gravar no Excel (arquivo pode estar aberto): {e}")


if __name__ == '__main__':
    ok = sincronizar()
    if not ok:
        sys.exit(1)
