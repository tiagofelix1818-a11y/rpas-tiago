# ================================================
# Monitor de e-mails Zendesk -> atualiza Excel
# Envia e-mail a Fornecedor e Coordenador
# (versão script Python pronta para execução)
# ================================================

import sys
import subprocess
import importlib.util
import os
import re
import pandas as pd
import xlwings as xw
import win32com.client as win32
from datetime import datetime, timedelta
from tqdm import tqdm


def ensure_installed(*pkgs):
    for p in pkgs:
        if importlib.util.find_spec(p) is None:
            print(f"Instalando {p}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", p])
    print("✓ Dependências checadas.")


ensure_installed("pywin32", "xlwings", "pandas", "openpyxl", "tqdm")

# ================================================================
# CONFIGURAÇÕES GLOBAIS
# ================================================================
EMAIL_CONTA = "tiagooliveira@pmenos.com.br"
PASTA_BASE = r"C:\\Users\\126815\\OneDrive - paguemenos.com.br\\ENGENHARIA - OBRAS - Documentos\\BANCO DE DADOS - POWER BI\\BI - OUTROS"
CAMINHO_BASE = os.path.join(PASTA_BASE, "BASE CONTROLE DE PAGAMENTOS.xlsx")
ABA_PAGAMENTO = "SOLICITAÇÃO DE PAGAMENTO"
ABA_EMAIL = "E-MAIL"
DIAS_FILTRO_EMAIL = 5  # janela de análise


# Cores para o terminal
class Cores:
    VERDE = '\033[92m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    AZUL = '\033[94m'
    RESET = '\033[0m'


# ================================================================
# E-mails dos COORDENADORES (sempre enviar ao responsável + CC Jordana)
# ================================================================
EMAILS_COORDENADORES = {
    'CAROL': 'carolinerocha@pmenos.com.br',
    'ANDRE': 'luizfrota@pmenos.com.br',
    'HENRIQUE': 'henriquesales@pmenos.com.br',
    # Adicione outros se necessário...
}
CC_FIXO = ['jordanasilva@pmenos.com.br']
# Para se copiar também, descomente:
# CC_FIXO.append('tiagooliveira@pmenos.com.br')


# ================================================================
# Regex (padrões robustos + fallbacks) — ORIGINAIS
# ================================================================
PADRAO_TICKET = re.compile(r"ticket\s*#?\s*(\d+)", re.IGNORECASE)
PADRAO_MIRO = re.compile(r"(?:[\u2022•·\-\s])?(?:n[ºo]?\s*)?miro[:\-#]*([0-9]{6,12})", re.IGNORECASE)
PADRAO_DATA = re.compile(
    r"(?:[\u2022•·\-\s])?(?:data\s+de\s+pagamento|data\s+pagamento|pagamento\s*(?:em)?)"
    r"[\s:\-]*"
    r"(\d{2}[./-]\d{2}[./-]\d{2,4})",
    re.IGNORECASE
)


def html_to_text(html: str) -> str:
    """Remove tags HTML simples e normaliza espaços/linhas."""
    if not html:
        return ""
    text = re.sub(r"<[^>]+>", " ", html)                 # remove tags
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    text = re.sub(r"[ \t]+", " ", text)                  # colapsa espaços
    text = re.sub(r"\s*\n\s*", "\n", text).strip()       # normaliza quebras
    return text


# Mantemos a função antiga (compatibilidade com e-mails antigos)
def parse_miro_and_payment(body_text: str):
    """
    Procura padrões de MIRO e Pagamento em qualquer lugar do texto,
    inclusive dentro do bloco 'Detalhes:' com bullets '•'.
    Aceita:
      • MIRO: 5106086423
      • Pagamento: 29/12/2025
    """
    if not body_text:
        return [], []
    miro_pattern = re.compile(r"(?:^\b)[•\-*]?\s*MIRO\s*[:\-]\s*([0-9]{8,12})", re.IGNORECASE | re.MULTILINE)
    date_pattern = re.compile(r"(?:^\b)[•\-*]?\s*Pagamento\s*[:\-]\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE | re.MULTILINE)
    miros = miro_pattern.findall(body_text)
    dates_str = date_pattern.findall(body_text)
    payment_dates = []
    for d in dates_str:
        try:
            payment_dates.append(datetime.strptime(d, "%d/%m/%Y").date())
        except ValueError:
            payment_dates.append(d)
    return miros, payment_dates


def matches_detalhes_block(body_text: str) -> bool:
    """Detecta a presença de um bloco 'Detalhes' seguido por MIRO e Pagamento."""
    pattern = re.compile(
        r"Detalhes\s*:[ \n\r].*MIRO\s*[:\-]\s*[0-9]{8,12}.*[ \n\r].*Pagamento\s*[:\-]\s*\d{2}/\d{2}/\d{4}",
        re.IGNORECASE | re.DOTALL
    )
    return bool(pattern.search(body_text))


# ================================================================
# NOVOS PADRÕES — e-mails simples com “Pagamento agendado … 27/02” + MIRO/MIGO
# ================================================================
PADRAO_MIRO_LINHA = re.compile(r"(?im)^[\s•\-\*\u2022]*MIRO\s*[:\-–—]?\s*([0-9]{6,12})\b")
PADRAO_MIGO_LINHA = re.compile(r"(?im)^[\s•\-\*\u2022]*MIGO\s*[:\-–—]?\s*([0-9]{6,12})\b")
PADRAO_PGTO_EXPLICITO = re.compile(
    r"(?i)\bPagamento(?:\s+agendado)?(?:\s+(?:para|em)\s+o?\s*dia)?\s*[:\-–—]?\s*(\d{2}[./-]\d{2}(?:[./-]\d{2,4})?)"
)
PADRAO_DATALHES_BLOCO = re.compile(
    r"(?is)Detalhes\s*:.*?MIRO\s*[:\-–—]\s*([0-9]{6,12}).*?Pagamento\s*[:\-–—]\s*(\d{2}/\d{2}/\d{2,4})"
)


def _parse_date_smart(ddmmyyyy_or_ddmm: str) -> datetime.date:
    s = ddmmyyyy_or_ddmm.strip().replace("-", "/").replace(".", "/")
    partes = s.split("/")
    if len(partes) == 2:
        dia, mes = partes
        ano = str(datetime.now().year)
    else:
        dia, mes, ano = partes[:3]
        if len(ano) == 2:
            ano = "20" + ano
    return datetime.strptime(f"{dia}/{mes}/{ano}", "%d/%m/%Y").date()


def parse_miro_migo_and_payment(body_text: str):
    if not body_text:
        return [], [], []

    # 1) Bloco 'Detalhes' tradicional
    m_bloco = PADRAO_DATALHES_BLOCO.search(body_text)
    if m_bloco:
        miro_num = m_bloco.group(1)
        data_str = m_bloco.group(2)
        try:
            data = _parse_date_smart(data_str)
        except Exception:
            data = None
        return ([miro_num] if miro_num else []), [], ([data] if data else [])

    # 2) Linhas soltas MIRO/MIGO + Pagamento
    miros = PADRAO_MIRO_LINHA.findall(body_text) or []
    migos = PADRAO_MIGO_LINHA.findall(body_text) or []

    datas = []
    for m in PADRAO_PGTO_EXPLICITO.finditer(body_text):
        raw = m.group(1)
        try:
            datas.append(_parse_date_smart(raw))
        except Exception:
            continue

    return miros, migos, datas


# ================================================================
# Utilitários
# ================================================================

def _formata_brl(valor) -> str:
    """Formata número para moeda BRL (R$)."""
    try:
        if pd.isna(valor):
            return ""
        if isinstance(valor, str):
            v = valor.replace('.', '').replace(',', '.')
            valor_float = float(v)
        else:
            valor_float = float(valor)
        return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(valor) if valor is not None else ""


def conectar_outlook():
    """Conecta-se ao Outlook e retorna a caixa de entrada e a pasta de processados."""
    print(f"{Cores.AZUL}📧 Conectando ao Outlook...{Cores.RESET}")
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        conta = session.Folders[EMAIL_CONTA]
        caixa_entrada = conta.Folders["Caixa de Entrada"]
        try:
            pasta_processados = conta.Folders["Processados"]
        except Exception:
            print(f"{Cores.AMARELO}Pasta 'Processados' não encontrada. Criando...{Cores.RESET}")
            pasta_processados = conta.Folders.Add("Processados")
        print(f"{Cores.VERDE}✅ Conectado com sucesso!{Cores.RESET}")
        return outlook, caixa_entrada, pasta_processados
    except Exception as e:
        print(f"{Cores.VERMELHO}❌ Erro fatal ao conectar ao Outlook: {e}{Cores.RESET}")
        return None, None, None


def filtrar_emails(caixa_entrada):
    """Filtra e-mails na caixa de entrada por data e 'ticket' no assunto."""
    print(f"{Cores.AZUL}📂 Lendo e-mails da Caixa de Entrada...{Cores.RESET}")
    data_limite = datetime.now() - timedelta(days=DIAS_FILTRO_EMAIL)
    itens = caixa_entrada.Items
    itens.Sort("[ReceivedTime]", True)

    # Variação de formato para Restrict (alguns ambientes exigem)
    formatos = ["%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M %p", "%m/%d/%Y %H:%M", "%m/%d/%Y %I:%M %p"]
    emails_filtrados = None
    for fmt in formatos:
        filtro = f"[ReceivedTime] >= '{data_limite.strftime(fmt)}'"
        try:
            resultado = itens.Restrict(filtro)
            if resultado.Count > 0:
                emails_filtrados = resultado
                break
        except:
            continue

    if not emails_filtrados or emails_filtrados.Count == 0:
        print(f"{Cores.AMARELO}ℹ️ Nenhum e-mail retornado pelo filtro de data.{Cores.RESET}")
        return []

    emails_validos = [email for email in emails_filtrados if re.search(r"ticket", (email.Subject or ""), re.IGNORECASE)]
    print(f"{Cores.AZUL}✉️ {len(emails_validos)} e-mails com 'ticket' no assunto encontrados.{Cores.RESET}")
    return emails_validos


def carregar_planilhas():
    """Carrega e normaliza as planilhas de pagamento e e-mail."""
    if not os.path.exists(CAMINHO_BASE):
        print(f"{Cores.VERMELHO}❌ Arquivo não encontrado: {CAMINHO_BASE}{Cores.RESET}")
        return None, None
    print(f"{Cores.AZUL}📊 Carregando dados do Excel...{Cores.RESET}")
    try:
        df_pagamento = pd.read_excel(CAMINHO_BASE, sheet_name=ABA_PAGAMENTO, engine="openpyxl")
        df_email = pd.read_excel(CAMINHO_BASE, sheet_name=ABA_EMAIL, engine="openpyxl")

        def normaliza_coluna(nome):
            return str(nome).strip().upper().replace("-", "").replace("_", "").replace(" ", "")

        df_pagamento.columns = [normaliza_coluna(col) for col in df_pagamento.columns]
        df_email.columns = [normaliza_coluna(col) for col in df_email.columns]

        if 'ID' not in df_pagamento.columns:
            print(f"{Cores.VERMELHO}❌ Coluna 'ID' não encontrada na aba '{ABA_PAGAMENTO}'.{Cores.RESET}")
            return None, None

        return df_pagamento, df_email
    except Exception as e:
        print(f"{Cores.VERMELHO}❌ Erro ao carregar o Excel: {e}{Cores.RESET}")
        return None, None


def enviar_email_fornecedor(outlook, linha_df, df_email, ticket_num, data_formatada, miro_num):
    """Busca o e-mail do fornecedor na aba E-MAIL e envia notificação de pagamento."""
    try:
        fornecedor = linha_df['FORNECEDOR'].iloc[0]
        assunto_planilha = linha_df['ASSUNTO'].iloc[0] if 'ASSUNTO' in linha_df.columns else ""
        nota_num = linha_df.get('NOTA', "Não Informada").iloc[0] if hasattr(linha_df, "get") else "Não Informada"

        email_fornecedor_df = df_email[df_email['FORNECEDOR'].str.lower() == str(fornecedor).lower()]
        if email_fornecedor_df.empty:
            return "sem_email"
        email_fornecedor = email_fornecedor_df['EMAIL'].iloc[0]

        mail = outlook.CreateItem(0)
        mail.To = email_fornecedor
        mail.Subject = f"Confirmação de Pagamento - Ticket {ticket_num}"
        mail.HTMLBody = f"""
        <p>Prezado(a) {fornecedor},</p>
        <p>Informamos que o pagamento referente ao <b>Ticket {ticket_num}</b> foi agendado/realizado.</p>
        <ul>
          <li><b>Assunto do Ticket:</b> {assunto_planilha}</li>
          <li><b>Data de Pagamento:</b> {data_formatada.strftime('%d/%m/%Y')}</li>
          <li><b>Número MIRO:</b> {miro_num or 'Não informado'}</li>
          <li><b>Número da Nota Fiscal:</b> {nota_num}</li>
        </ul>
        <p>Atenciosamente,<br>Equipe Engenharia - Pague Menos</p>
        """
        mail.Send()
        return "enviado"
    except Exception:
        return "erro_envio"


def enviar_email_coordenador(outlook, linha_df, ticket_num, data_formatada, miro_num):
    """
    Envia e-mail ao COORDENADOR responsável.
    Sempre CC Jordana. Não depende da aba E-MAIL.
    Inclui o campo PEDIDO e VALOR RC (quando houver).
    """
    try:
        coord = linha_df['COORDENADOR'].iloc[0] if 'COORDENADOR' in linha_df.columns else None
        if coord is None or str(coord).strip() == "":
            return "sem_coordenador"

        coord_key = str(coord).strip().upper()
        dest = EMAILS_COORDENADORES.get(coord_key)
        if not dest:
            return "coord_sem_email"

        fornecedor = linha_df['FORNECEDOR'].iloc[0] if 'FORNECEDOR' in linha_df.columns else ""
        assunto_planilha = linha_df['ASSUNTO'].iloc[0] if 'ASSUNTO' in linha_df.columns else ""
        nota_num = linha_df['NOTA'].iloc[0] if 'NOTA' in linha_df.columns else "Não Informada"
        valor_bi = linha_df['VALORBI'].iloc[0] if 'VALORBI' in linha_df.columns else None
        valor_rc = linha_df['VALORRC'].iloc[0] if 'VALORRC' in linha_df.columns else valor_bi

        pedido_val = ""
        if 'PEDIDO' in linha_df.columns:
            raw = linha_df['PEDIDO'].iloc[0]
            if pd.isna(raw):
                pedido_val = ""
            else:
                pedido_val = str(int(float(raw))) if isinstance(raw, (int, float)) else str(raw).strip()

        tabela_html = f"""
        <table style="border-collapse:collapse;font-family:Segoe UI, Arial,sans-serif;font-size:13px">
          <tr><th style="background:#0060A8;color:#fff;padding:6px 8px;text-align:left">Campo</th>
              <th style="background:#0060A8;color:#fff;padding:6px 8px;text-align:left">Valor</th></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Ticket</td><td style="padding:6px 8px;border-top:1px solid #eee">{ticket_num}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Assunto</td><td style="padding:6px 8px;border-top:1px solid #eee">{assunto_planilha}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Fornecedor</td><td style="padding:6px 8px;border-top:1px solid #eee">{fornecedor}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Nota Fiscal</td><td style="padding:6px 8px;border-top:1px solid #eee">{nota_num}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Data de Pagamento</td><td style="padding:6px 8px;border-top:1px solid #eee">{data_formatada.strftime('%d/%m/%Y')}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">MIRO</td><td style="padding:6px 8px;border-top:1px solid #eee">{miro_num or 'Não informado'}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Pedido</td><td style="padding:6px 8px;border-top:1px solid #eee">{pedido_val or 'Não informado'}</td></tr>
          <tr><td style="padding:6px 8px;border-top:1px solid #eee">Valor RC</td><td style="padding:6px 8px;border-top:1px solid #eee">{_formata_brl(valor_rc)}</td></tr>
        </table>
        """

        corpo_html = f"""
        <html>
        <body style="font-family:Segoe UI, Arial,sans-serif">
            <p>Olá, {coord_key.title()},</p>
            <p>Segue atualização do ticket com confirmação/agendamento de pagamento:</p>
            {tabela_html}
            <p>Qualquer dúvida, fico à disposição.</p>
        </body>
        </html>
        """

        mail = outlook.CreateItem(0)
        mail.To = dest
        mail.CC = "; ".join(CC_FIXO) if CC_FIXO else ""
        mail.Subject = f"[Ticket {ticket_num}] Pagamento confirmado – {fornecedor}"
        mail.HTMLBody = corpo_html
        mail.Send()
        return "enviado"
    except Exception:
        return "erro_envio"


def main():
    # Verifica arquivo base
    if not os.path.exists(CAMINHO_BASE):
        print(f"{Cores.VERMELHO}❌ Arquivo não encontrado: {CAMINHO_BASE}{Cores.RESET}")
        return

    # Conecta Outlook
    outlook_app, caixa_entrada, pasta_processados = conectar_outlook()
    if not caixa_entrada:
        return

    # Filtra e-mails
    emails_validos = filtrar_emails(caixa_entrada)
    if not emails_validos:
        return

    # Carrega planilhas
    df_pagamento, df_email = carregar_planilhas()
    if df_pagamento is None:
        return

    # Abre Excel via xlwings
    print(f"{Cores.AZUL}⚙️ Inicializando instância do Excel para atualização...{Cores.RESET}")
    try:
        app = xw.App(visible=False)
        wb = app.books.open(CAMINHO_BASE)
        ws_pag = wb.sheets[ABA_PAGAMENTO]
    except Exception as e:
        print(f"{Cores.VERMELHO}❌ Erro ao abrir o Excel via xlwings: {e}{Cores.RESET}")
        return

    # --- Processamento ---
    print(f"\n{Cores.AZUL}🚀 Processando e-mails...{Cores.RESET}")
    contadores = {
        "total": len(emails_validos),
        "atualizados": 0,
        "ja_preenchidos": 0,
        "nao_encontrados": 0,
        "sem_data": 0,
        "erros": 0,
        "emails_enviados_fornecedor": 0,
        "emails_enviados_coordenador": 0,
        "coord_sem_email": 0
    }
    logs = {
        "preenchidos": [],
        "nao_encontrados": [],
        "sem_data": [],
        "erros": [],
        "coord_sem_email": []
    }

    for email in tqdm(emails_validos, desc="Analisando e-mails", unit="email"):
        ticket_num = None
        try:
            assunto_email = email.Subject or ""
            html = getattr(email, "HTMLBody", None)
            corpo_texto = html_to_text(html or (email.Body or ""))

            # 1) Ticket (assunto/corpo)
            match_ticket = PADRAO_TICKET.search(assunto_email) or PADRAO_TICKET.search(corpo_texto)
            if not match_ticket:
                tqdm.write(f"{Cores.AMARELO}⏭️ E-mail ignorado (sem ticket): {assunto_email[:50]}{Cores.RESET}")
                continue
            ticket_num = int(match_ticket.group(1).strip())

            # 2) Ticket existe na base?
            linha_df = df_pagamento[df_pagamento['ID'] == ticket_num]
            if linha_df.empty:
                contadores["nao_encontrados"] += 1
                logs["nao_encontrados"].append(ticket_num)
                tqdm.write(f"{Cores.AMARELO}⚠️ Ticket {ticket_num} não encontrado na base (e-mail não movido){Cores.RESET}")
                continue

            # 3) Já preenchido?
            if 'DATAPGTOSAP' in linha_df.columns and pd.notnull(linha_df['DATAPGTOSAP'].iloc[0]):
                contadores["ja_preenchidos"] += 1
                logs["preenchidos"].append(ticket_num)
                email.Move(pasta_processados)
                tqdm.write(f"{Cores.AZUL}ℹ️ Ticket {ticket_num} já estava preenchido (movido para Processados){Cores.RESET}")
                continue

            # 4) Extrai MIRO/MIGO e Data (novo padrão + retrocompatível)
            miros, migos, payment_dates = parse_miro_migo_and_payment(corpo_texto)

            miro_num = miros[0] if miros else None

            if payment_dates:
                data_formatada = payment_dates[0]
            else:
                match_data = PADRAO_DATA.search(corpo_texto) if 'PADRAO_DATA' in globals() else None
                if not match_data:
                    contadores["sem_data"] += 1
                    logs["sem_data"].append(ticket_num)
                    motivo = "Padrão de pagamento não detectado"
                    try:
                        data_receb = email.ReceivedTime.strftime("%d/%m/%Y %H:%M")
                    except Exception:
                        data_receb = str(email.ReceivedTime)
                    tqdm.write(
                        f"[IGNORADO] Assunto='{assunto_email[:80]}' Remetente='{email.SenderName}' "
                        f"Data='{data_receb}' Motivo='{motivo}'"
                    )
                    continue
                data_str = match_data.group(1).strip().replace(".", "/").replace("-", "/")
                partes = (data_str + "/").split("/")[:3]
                dia, mes = partes[0], partes[1]
                ano = partes[2] if len(partes) > 2 and partes[2] else str(datetime.now().year)
                if len(ano) == 2:
                    ano = "20" + ano
                data_formatada = datetime.strptime(f"{dia}/{mes}/{ano}", "%d/%m/%Y").date()

            excel_row_index = linha_df.index[0] + 2
            if 'DATAPGTOSAP' in df_pagamento.columns:
                ws_pag.range(excel_row_index, df_pagamento.columns.get_loc('DATAPGTOSAP') + 1).value = data_formatada
            if miro_num and 'MIRO' in df_pagamento.columns:
                ws_pag.range(excel_row_index, df_pagamento.columns.get_loc('MIRO') + 1).value = miro_num

            status_envio_forn = enviar_email_fornecedor(outlook_app, linha_df, df_email, ticket_num, data_formatada, miro_num)
            if status_envio_forn == "enviado":
                contadores["emails_enviados_fornecedor"] += 1
                status_email_log = f"{Cores.VERDE}📧 E-mail ao fornecedor enviado{Cores.RESET}"
            elif status_envio_forn == "sem_email":
                status_email_log = f"{Cores.AMARELO}⚠️ Fornecedor sem e-mail na aba E-MAIL{Cores.RESET}"
            else:
                status_email_log = f"{Cores.VERMELHO}❌ Erro ao enviar e-mail ao fornecedor{Cores.RESET}"

            status_coord = enviar_email_coordenador(outlook_app, linha_df, ticket_num, data_formatada, miro_num)
            if status_coord == "enviado":
                contadores["emails_enviados_coordenador"] += 1
            elif status_coord == "coord_sem_email":
                contadores["coord_sem_email"] += 1
                coord_name = str(linha_df['COORDENADOR'].iloc[0]) if 'COORDENADOR' in linha_df.columns else "(sem coordenador)"
                logs["coord_sem_email"].append(coord_name)

            tqdm.write(
                f"{Cores.VERDE}✅ Ticket {ticket_num}: Data {data_formatada.strftime('%d/%m/%Y')} "
                f"MIRO {miro_num or 'N/A'}{Cores.RESET} | {status_email_log}"
            )

            email.Move(pasta_processados)
            contadores["atualizados"] += 1

        except Exception as e:
            contadores["erros"] += 1
            error_msg = f"Ticket {ticket_num if ticket_num else 'N/A'}: {e}"
            logs["erros"].append(error_msg)
            tqdm.write(f"{Cores.VERMELHO}💥 Erro ao processar e-mail: {error_msg}{Cores.RESET}")

    print(f"\n{Cores.AZUL}💾 Salvando e fechando o Excel...{Cores.RESET}")
    try:
        wb.save()
        wb.close()
        app.quit()
        print(f"{Cores.VERDE}✅ Excel salvo e fechado com sucesso.{Cores.RESET}")
    except Exception as e:
        print(f"{Cores.VERMELHO}❌ Erro ao salvar/fechar o Excel: {e}{Cores.RESET}")

    print("\n" + "="*50)
    print("📊 RESUMO FINAL DA EXECUÇÃO".center(50))
    print("="*50)
    print(f"🔢 Total de e-mails com 'ticket' analisados: {contadores['total']}")
    print(f"{Cores.VERDE}✅ Tickets atualizados na planilha: {contadores['atualizados']}{Cores.RESET}")
    print(f"{Cores.VERDE}📧 E-mails enviados ao fornecedor: {contadores['emails_enviados_fornecedor']}{Cores.RESET}")
    print(f"{Cores.VERDE}👤 E-mails enviados ao coordenador: {contadores['emails_enviados_coordenador']}{Cores.RESET}")
    print(f"{Cores.AZUL}ℹ️ Tickets já preenchidos (movidos): {contadores['ja_preenchidos']}{Cores.RESET}")
    if logs['preenchidos']:
        print(f" ↳ Tickets: {', '.join(map(str, logs['preenchidos'][:10]))}{'...' if len(logs['preenchidos']) > 10 else ''}")

    print(f"\n{Cores.AMARELO}⏭️ E-mails NÃO processados (permanecem na Caixa de Entrada):{Cores.RESET}")
    print(f" • Sem data de pagamento: {contadores['sem_data']}")
    if logs['sem_data']:
        print(f" ↳ Tickets: {', '.join(map(str, logs['sem_data'][:10]))}{'...' if len(logs['sem_data']) > 10 else ''}")
    print(f" • Ticket não encontrado na base: {contadores['nao_encontrados']}")
    if logs['nao_encontrados']:
        print(f" ↳ Tickets: {', '.join(map(str, logs['nao_encontrados'][:10]))}{'...' if len(logs['nao_encontrados']) > 10 else ''}")

    if contadores['coord_sem_email'] > 0:
        print(f"\n{Cores.AMARELO}⚠️ Coordenadores sem e-mail no dicionário: {contadores['coord_sem_email']}{Cores.RESET}")
        if logs['coord_sem_email']:
            print(f" ↳ {', '.join(map(str, logs['coord_sem_email'][:10]))}{'...' if len(logs['coord_sem_email']) > 10 else ''}")

    if contadores['erros'] > 0:
        print(f"\n{Cores.VERMELHO}💥 E-mails com erro: {contadores['erros']}{Cores.RESET}")
        for erro in logs['erros'][:5]:
            print(f" - {erro}")
        excedentes = max(len(logs['erros']) - 5, 0)
        if excedentes > 0:
            print(f" ... e mais {excedentes} erros")
    print("="*50)


if __name__ == "__main__":
    main()
