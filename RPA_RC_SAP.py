# ----------------- CHECAGEM DE DEPENDÊNCIAS (Script) -----------------
import sys, subprocess, importlib.util

def ensure_installed(*pkgs):
    for p in pkgs:
        if importlib.util.find_spec(p) is None:
            print(f"Instalando {p}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", p])
    print("✓ Dependências checadas.")

ensure_installed("pywin32", "xlwings", "pandas")

# ----------------- IMPORTS -----------------
import win32com.client
import xlwings as xw
import time
import pandas as pd
import unicodedata
import re
import os

# ----------------- FUNÇÕES AUXILIARES -----------------
def normalizar_coluna(texto):
    if not texto:
        return ""
    texto = str(texto).strip().upper()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'\s+', '', texto)
    return texto

def valor_sap(v):
    """
    Converte valores para o formato que o SAP aceita (vírgula decimal, sem milhar),
    respeitando a sua lógica original.
    """
    if pd.isna(v) or str(v).strip() == "":
        return ""
    return str(v).replace(".", "#").replace(",", ".").replace("#", ",").replace(".", "")

def formatar_valor_br(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""
    try:
        valor_str = str(v).strip().replace(' ', '').replace(',', '.')
        valor = float(valor_str)
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return str(v)

def verificar_sessao(session):
    try:
        _ = session.findById("wnd[0]/tbar[0]/okcd").text
        return True
    except:
        return False

def aguardar_controle(session, id_controle, timeout=15):
    for _ in range(timeout):
        try:
            ctrl = session.findById(id_controle)
            return ctrl
        except:
            time.sleep(1)
    return None

# ----------------- LEITURA E TRATAMENTO PLANILHA -----------------
PATH_XLSM = r'C:\AUTOMACAO_RC_SAP\base_rc_sap.xlsm'
if not os.path.exists(PATH_XLSM):
    raise FileNotFoundError(f"Planilha não encontrada em:\n{PATH_XLSM}")

df_rc     = pd.read_excel(PATH_XLSM, sheet_name='sheet1', dtype=str)
df_texto  = pd.read_excel(PATH_XLSM, sheet_name='base_texto', dtype=str)

df_rc.columns    = [normalizar_coluna(c) for c in df_rc.columns]
df_texto.columns = [normalizar_coluna(c) for c in df_texto.columns]

col_rc_validas = ['ORDEM CIC', 'MATERIAL', 'QTDE', 'VALOR UNITARIO', 'TOTAL ITEM', 'DATA', 'LOJA', 'RC_SEQ', 'ARQUIVO', 'IDPADRAO', 'STATUS', 'NUMERO_RC']
col_idx = {c: df_rc.columns.get_loc(normalizar_coluna(c)) for c in col_rc_validas if normalizar_coluna(c) in df_rc.columns}

# Normalizações de dados
if 'LOJA' in df_rc.columns:
    df_rc['LOJA'] = df_rc['LOJA'].apply(lambda x: str(x).replace('.0', '').strip().zfill(4))
if 'QTDE' in df_rc.columns:
    # ✅ CORREÇÃO: preserva casas decimais quando o valor não é inteiro.
    # Antes usava int(float(x)), o que arredondava quantidades como 200.100,01 → 200100.
    def tratar_qtde(x):
        if pd.isnull(x) or str(x).strip() == '':
            return '0'
        # Normaliza separadores: remove milhar e identifica decimal
        valor_str = str(x).strip().replace(' ', '')
        # Detecta formato BR (ex: "200.100,01") vs formato padrão (ex: "200100.01")
        if ',' in valor_str:
            # Formato BR: ponto = milhar, vírgula = decimal
            valor_str = valor_str.replace('.', '').replace(',', '.')
        else:
            # Formato padrão: pode ter ponto como milhar ou decimal
            # Se houver mais de um ponto, todos são milhar (ex: "200.100")
            partes = valor_str.split('.')
            if len(partes) > 2:
                valor_str = valor_str.replace('.', '')
        f = float(valor_str)
        # Se for inteiro exato, retorna sem casas decimais
        if f == int(f):
            return str(int(f))
        else:
            # Retorna com VÍRGULA como decimal (formato SAP/BR)
            # ex: 4316.13 → "4316,13"
            decimal_str = f'{f:.10f}'.rstrip('0').rstrip('.')
            return decimal_str.replace('.', ',')
    df_rc['QTDE'] = df_rc['QTDE'].apply(tratar_qtde)
if 'DATA' in df_rc.columns:
    df_rc['DATA'] = df_rc['DATA'].apply(lambda x: str(x).strip() if pd.notnull(x) and str(x).strip() != '' else '0')

# ----------------- AGRUPAR POR RC_SEQ -----------------
dados = df_rc.values.tolist()
rqs_por_rc = {}
for i, linha in enumerate(dados, start=2):  # Excel: cabeçalho na linha 1
    rc_seq = str(linha[col_idx['RC_SEQ']]).strip()
    rqs_por_rc.setdefault(rc_seq, []).append((i, linha))

# ----------------- ABRIR PLANILHA -----------------
wb = xw.Book(PATH_XLSM)
ws = wb.sheets[0]

# ----------------- CONEXÃO SAP -----------------
SapGuiAuto  = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection  = application.Children[0]
session     = connection.Children[0]

# ----------------- PROCESSAR RC_SEQ -----------------
for rc_seq, itens in rqs_por_rc.items():
    print(f"\n🔷 Processando RC_SEQ {rc_seq} com {len(itens)} itens")

    # --- Condição: já processado? (U = 'Feito pela automação') ---
    excel_row_inicial  = itens[0][0]
    status_rc          = ws.range(f"U{excel_row_inicial}").value
    if status_rc and str(status_rc).strip().upper() == 'FEITO PELA AUTOMAÇÃO':
        print(f"✅ RC_SEQ {rc_seq} já processado (Status: {status_rc}). Pulando...")
        continue

    if not verificar_sessao(session):
        print("⚠️ Sessão desconectada. Reconectando...")
        SapGuiAuto  = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection  = application.Children[0]
        session     = connection.Children[0]

    # Abre ME51N
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # Buscar linha correspondente por IDPADRAO na base_texto
    id_padrao = str(itens[0][1][col_idx['IDPADRAO']]).strip()
    linha_correspondente = None
    for _, row in df_texto.iterrows():
        if str(row['IDPADRAO']).strip() == id_padrao:
            linha_correspondente = row
            break

    if linha_correspondente is None:
        print(f"⚠️ Nenhuma linha correspondente encontrada na base_texto para IDPADRAO {id_padrao}")
        continue

    # --- Definir tipo de requisição (SOBR/SOB1) ---
    combo_tipo_rc_id = ("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
                        "subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/"
                        "cmbMEREQ_TOPLINE-BSART")
    combo_tipo_rc = aguardar_controle(session, combo_tipo_rc_id, timeout=10)
    if not combo_tipo_rc:
        print(f"⚠️ Controle do tipo de requisição não encontrado para RC_SEQ {rc_seq}. Pulando...")
        continue
        
    projeto_rc = str(linha_correspondente['PROJETO']).strip().upper()
    if projeto_rc == 'LOJAS NOVAS':
        tipo_rc_key  = "SOB1"
        tipo_rc_desc = "Setor Obras Loja Nov (SOB1)"
    else:
        tipo_rc_key  = "SOBR"
        tipo_rc_desc = "Setor Obras (SOBR)"

    try:
        combo_tipo_rc.key = tipo_rc_key
        time.sleep(2)
        print(f"🏗️ Tipo de requisição alterado para '{tipo_rc_desc}'.")
    except Exception as e:
        print(f"⚠️ Erro ao alterar tipo de requisição para RC_SEQ {rc_seq}: {e}")
        continue

    # --- Preencher texto do cabeçalho (aba Textos) ---
    campos = [
        ("PROJETO",      linha_correspondente['PROJETO']),
        ("FILIAL",       linha_correspondente['FILIAL']),
        ("LOJA",         linha_correspondente['LOJA']),
        ("CNPJ",         linha_correspondente['CNPJ']),
        ("COORDENADOR",  linha_correspondente['COORDENADOR']),
        ("SERVIÇO",      linha_correspondente['SERVICO']),
        ("FORNECEDOR",   linha_correspondente['FORNECEDOR']),
        ("ORDEM",        linha_correspondente['ORDEM']),
        ("VALOR RC",     formatar_valor_br(linha_correspondente['VALORRC'])),
        ("ESCOPO",       linha_correspondente['ESCOPO'])
    ]
    texto_linha = "\n".join(f"{k}\t{v}" for k, v in campos)

    try:
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
                         "subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                         "subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1").select()
        time.sleep(1)
        sap_texto_id = ("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
                        "subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                        "subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/"
                        "ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/"
                        "subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")
        ctrl_texto = aguardar_controle(session, sap_texto_id)
        if ctrl_texto:
            ctrl_texto.text = texto_linha
            print("📝 Cabeçalho transposto preenchido com sucesso.")
        else:
            print("⚠️ Campo de texto do cabeçalho não encontrado.")
    except Exception as e:
        print(f"⚠️ Erro ao preencher texto do cabeçalho: {e}")

    # --- GRID de itens ---
    grid_id = ("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/"
               "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
               "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
    grid = aguardar_controle(session, grid_id, timeout=10)
    if not grid:
        print(f"❌ Grid de itens não encontrado para RC_SEQ {rc_seq}. Pulando...")
        continue

    for idx, linha in enumerate(itens):
        excel_row, dados = linha
        
        # Valores de item
        material                 = str(dados[col_idx['MATERIAL']]).strip()
        valor_unitario_original  = str(dados[col_idx['VALOR UNITARIO']])
        valor_total_item         = str(dados[col_idx['TOTAL ITEM']])
        valor_unitario_para_sap  = valor_unitario_original

        # Regra especial de negócio
        if material == '20103':
            print(f"⚠️ Material {material} detectado. Alterando VALOR UNITARIO de {valor_unitario_original} para 1.")
            valor_unitario_para_sap = '1'

        # Grid
        grid.modifyCell(idx, "KNTTP", "F")
        grid.modifyCell(idx, "MATNR", material)
        grid.modifyCell(idx, "MENGE", str(dados[col_idx['QTDE']]))
        grid.modifyCell(idx, "PREIS", valor_sap(valor_unitario_para_sap))
        grid.modifyCell(idx, "GSWRT", valor_sap(valor_total_item))
        grid.modifyCell(idx, "EEIND", str(dados[col_idx['DATA']]))
        grid.modifyCell(idx, "NAME1", str(dados[col_idx['LOJA']]))
        grid.setCurrentCell(idx, "NAME1")
        grid.pressEnter()
        time.sleep(1)
        
        # Contabilização (Ordem CIC)
        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                             "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                             "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
                             "tabsREQ_ITEM_DETAIL/tabpTABREQDT7/"
                             "ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/"
                             "subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/"
                             "subKONTBLOCK:SAPLKACB:2100/ctxtCOBL-AUFNR").text = str(dados[col_idx['ORDEM CIC']])
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
        except Exception:
            pass
        
        # Avaliação (quando preço unitário > 1)
        try:
            valor_unitario_str_aval = valor_unitario_para_sap.replace(",", ".")
            if float(valor_unitario_str_aval) > 1:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                                 "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                 "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
                                 "tabsREQ_ITEM_DETAIL/tabpTABREQDT6").select()
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                campo_preco_aval = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/"
                                                    "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                                                    "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
                                                    "tabsREQ_ITEM_DETAIL/tabpTABREQDT6/"
                                                    "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3320/txtMEREQ3320-PREIS")
                campo_preco_aval.text = valor_sap(valor_unitario_para_sap)
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
        except Exception:
            pass
        
        # Volta para Dados gerais do item
        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/"
                             "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/"
                             "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/"
                             "tabsREQ_ITEM_DETAIL/tabpTABREQDT1").select()
            time.sleep(1)
        except:
            pass

    # ------------------------------------------------------------------
    # 📎 ANEXAR MÚLTIPLOS DOCUMENTOS VIA GOS — ANTES DO SAVE (F11)
    # Lógica: dentro do mesmo RC_SEQ, agrupar por 'ARQUIVO' e anexar
    # todos os arquivos da pasta do fornecedor cujo nome INICIE com o prefixo.
    # Evita anexos duplicados no mesmo RC.
    # ------------------------------------------------------------------
    anexo_status = "N/D"
    anexo_nome   = ""
    try:
        if 'ARQUIVO' not in col_idx:
            raise KeyError("Coluna 'ARQUIVO' não encontrada na planilha (sheet1).")

        projeto    = str(linha_correspondente['PROJETO']).strip()
        fornecedor = str(linha_correspondente['FORNECEDOR']).strip()

        # Normalização específica
        if 'ALVO K' in fornecedor.upper() or 'ALVO - K' in fornecedor.upper():
            fornecedor = 'ALVO - K'

        CAMINHO_BASE  = r"C:\Users\126815\OneDrive - paguemenos.com.br\Área de Trabalho\ABERTURA DE REQUISIÇÃO"
        CAMINHO_ANEXO = os.path.join(CAMINHO_BASE, projeto, fornecedor)

        if not os.path.isdir(CAMINHO_ANEXO):
            raise FileNotFoundError(f"Pasta de anexo inexistente: {CAMINHO_ANEXO}")

        # Coletar todos os prefixos 'ARQUIVO' deste RC_SEQ (deduplicados)
        prefixos = []
        for _, linha_item in itens:
            pfx = str(linha_item[col_idx['ARQUIVO']]).strip()
            if pfx:
                prefixos.append(pfx.upper())
        prefixos_unicos = sorted(set(prefixos))

        if not prefixos_unicos:
            raise ValueError("Nenhum valor em 'ARQUIVO' nas linhas deste RC_SEQ.")

        # Arquivos candidatos na pasta do fornecedor
        candidatos = [
            f for f in os.listdir(CAMINHO_ANEXO)
            if os.path.isfile(os.path.join(CAMINHO_ANEXO, f))
        ]

        anexos_ok    = []
        anexos_falha = []
        ja_anexados  = set()  # evita anexar o mesmo arquivo mais de uma vez neste RC

        # Para cada prefixo de ARQUIVO, anexar TODOS os arquivos que iniciem por esse prefixo
        for pref in prefixos_unicos:
            arquivos_match = [
                f for f in candidatos
                if f.upper().startswith(pref + "#") or f.upper().startswith(pref)
            ]

            if not arquivos_match:
                anexos_falha.append(f"{pref}: nenhum arquivo correspondente na pasta")
                continue

            for fname in sorted(arquivos_match):
                key = fname.upper()
                if key in ja_anexados:
                    continue  # já anexado neste RC

                try:
                    print(f"📎 Anexando ({pref}) → {fname}")
                    session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
                    session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_PCATTA_CREA")
                    time.sleep(1)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text     = CAMINHO_ANEXO
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname
                    session.findById("wnd[1]").sendVKey(0)
                    time.sleep(2)

                    anexos_ok.append(fname)
                    ja_anexados.add(key)

                except Exception as e:
                    anexos_falha.append(f"{fname}: {e}")
                    # tenta fechar popup pendente (se houver)
                    try:
                        session.findById("wnd[1]/tbar[0]/btn[12]").press()  # Cancel
                    except:
                        pass

        # Resultado consolidado para Excel (colunas X e Y)
        if anexos_ok:
            anexo_status = f"OK ({len(anexos_ok)} arquivo(s))"
            preview = "; ".join(anexos_ok[:10])
            anexo_nome = preview if len(anexos_ok) <= 10 else preview + f"; ... (+{len(anexos_ok)-10})"
        else:
            anexo_status = "ERRO: nenhum anexo realizado"
            anexo_nome   = ""

        print(f"✅ Anexos realizados: {len(anexos_ok)} | ❌ Falhas: {len(anexos_falha)}")
        if anexos_falha:
            print("Detalhes (falhas): " + " | ".join(anexos_falha[:5]))

    except Exception as e:
        anexo_status = f"ERRO: {e}"
        anexo_nome   = ""
        print(f"⚠️ Falha no anexo (RC_SEQ {rc_seq}): {e}")

    # Registra status do anexo no Excel (X e Y) para todas as linhas do grupo
    for excel_row, _ in itens:
        ws.range(f"X{excel_row}").value = anexo_status
        ws.range(f"Y{excel_row}").value = anexo_nome

    # ------------------------------------------------------------------
    # 💾 SALVAR RC (F11) — DEPOIS DO ANEXO
    # ------------------------------------------------------------------
    print(f"💾 Salvando RC do RC_SEQ {rc_seq}...")
    session.findById("wnd[0]").sendVKey(11)
    time.sleep(2)
    
    # Popup eventual
    try:
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        time.sleep(1)
    except:
        pass  # Sem popup, segue

    # --- Capturar nº RC na barra de status ---
    rc_num_texto = session.findById("wnd[0]/sbar").Text  # ex.: "Requisição de compra criada sob nº 0010024898"
    print(f"🆔 Status bar: {rc_num_texto}")
    m_num = re.search(r"(\d{5,})", rc_num_texto or "")
    rc_num_puro = m_num.group(1) if m_num else ""

    # Atualiza Status e Nº RC no Excel (U e V)
    for excel_row, _ in itens:
        ws.range(f"U{excel_row}").value = "Feito pela automação"
        ws.range(f"V{excel_row}").value = rc_num_texto

# Salvar planilha
try:
    wb.save()
    # wb.close()  # mantém seu padrão
    print("\n✅ Planilha salva com sucesso.")
except Exception as e:
    print(f"\n❌ Erro ao salvar a planilha: {e}")
