#!/usr/bin/env python
# coding: utf-8

# In[1]:


# ================================================================
# Monitor de E-mails – Pedidos de Compra (PDF)
# Filtro: últimos 7 dias | remetente Paula
# ================================================================

import os
import re
import pandas as pd
from datetime import datetime, timedelta
from PyPDF2 import PdfReader
import win32com.client as win32

# ================================================================
# CONFIGURAÇÕES
# ================================================================
EMAIL_CONTA = "tiagooliveira@pmenos.com.br"
REMETENTE_AUTORIZADO = "paula"     # filtro simples por nome
DIAS_FILTRO_EMAIL = 7

PASTA_PDF = r"C:\Users\126815\OneDrive - paguemenos.com.br\Área de Trabalho\LEITOR_PDF_COMPRA"
EXCEL_SAIDA = os.path.join(PASTA_PDF, "resumo_pedidos_compra.xlsx")

# ================================================================
# CORES (terminal)
# ================================================================
class Cores:
    VERDE = '\033[92m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    AZUL = '\033[94m'
    RESET = '\033[0m'

# ================================================================
# OUTLOOK
# ================================================================
def conectar_outlook():
    print(f"{Cores.AZUL}📧 Conectando ao Outlook...{Cores.RESET}")
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.Session
    conta = session.Folders[EMAIL_CONTA]
    inbox = conta.Folders["Caixa de Entrada"]

    # Garante pasta Processados\Pedidos de Compra
    try:
        processados = conta.Folders["Processados"].Folders["Pedidos de Compra"]
    except Exception:
        try:
            processados_pai = conta.Folders["Processados"]
        except Exception:
            processados_pai = conta.Folders.Add("Processados")
        processados = processados_pai.Folders.Add("Pedidos de Compra")

    print(f"{Cores.VERDE}✅ Outlook conectado{Cores.RESET}")
    return inbox, processados

# ================================================================
# LEITURA DO PDF
# ================================================================
def extrair_dados_pdf(caminho_pdf):
    reader = PdfReader(caminho_pdf)
    texto = ""

    for pagina in reader.pages:
        texto += pagina.extract_text() + "\n"

    pedido = None
    fornecedor = None
    cnpj_fatura = None
    valor_total = None

    # Nº do Pedido
    m_pedido = re.search(r"N[°ºo]?\s*do\s*Pedido:\s*(\d{8,})", texto, re.IGNORECASE)
    if m_pedido:
        pedido = m_pedido.group(1)

    # Fornecedor
    m_forn = re.search(r"Fornecedor Nome:\s*(.+?)\s*Endereço:", texto, re.DOTALL)
    if m_forn:
        fornecedor = m_forn.group(1).strip()

    # CNPJ da Fatura
    m_cnpj = re.search(r"Endereço de Fatura.*?CNPJ:\s*([\d\./-]+)", texto, re.DOTALL)
    if m_cnpj:
        cnpj_fatura = m_cnpj.group(1)

    # Último Total
    totais = re.findall(r"Total:\s*([\d\.,]+)", texto)
    if totais:
        valor_total = totais[-1]

    return pedido, fornecedor, cnpj_fatura, valor_total

# ================================================================
# MAIN
# ================================================================
def main():
    if not os.path.exists(PASTA_PDF):
        os.makedirs(PASTA_PDF)

    inbox, processados = conectar_outlook()
    data_limite = datetime.now() - timedelta(days=DIAS_FILTRO_EMAIL)

    registros = []

    print(f"{Cores.AZUL}📥 Varredura de e-mails (últimos {DIAS_FILTRO_EMAIL} dias)...{Cores.RESET}")

    for email in list(inbox.Items):
        if email.Class != 43:
            continue

        # -------- filtro de data --------
        try:
            recebido_em = email.ReceivedTime.replace(tzinfo=None)
        except Exception:
            continue

        if recebido_em < data_limite:
            continue

        # -------- filtro remetente --------
        remetente = (email.SenderName or "").lower()
        if REMETENTE_AUTORIZADO not in remetente:
            continue

        possui_pdf = False

        for anexo in email.Attachments:
            if not anexo.FileName.lower().endswith(".pdf"):
                continue

            possui_pdf = True
            caminho_pdf = os.path.join(PASTA_PDF, anexo.FileName)
            anexo.SaveAsFile(caminho_pdf)

            pedido, fornecedor, cnpj, valor = extrair_dados_pdf(caminho_pdf)

            registros.append({
                "DATA_PROCESSAMENTO": datetime.now(),
                "PEDIDO": pedido,
                "FORNECEDOR": fornecedor,
                "CNPJ_FATURA": cnpj,
                "VALOR_TOTAL": valor,
                "ARQUIVO_PDF": anexo.FileName,
                "EMAIL_REMETENTE": email.SenderName
            })

            print(
                f"{Cores.VERDE}✅ Pedido {pedido} | {fornecedor} | Total {valor}{Cores.RESET}"
            )

        if possui_pdf:
            email.Move(processados)

    if not registros:
        print(f"{Cores.AMARELO}ℹ️ Nenhum pedido novo encontrado.{Cores.RESET}")
        return

    df_novo = pd.DataFrame(registros)

    if os.path.exists(EXCEL_SAIDA):
        df_exist = pd.read_excel(EXCEL_SAIDA)
        df_final = pd.concat([df_exist, df_novo], ignore_index=True)
    else:
        df_final = df_novo

    df_final.to_excel(EXCEL_SAIDA, index=False)

    print(f"{Cores.VERDE}📊 Excel atualizado com sucesso{Cores.RESET}")
    print(f"📄 {EXCEL_SAIDA}")

# ================================================================
# EXECUÇÃO
# ================================================================
if __name__ == "__main__":
    main()


# In[ ]:




