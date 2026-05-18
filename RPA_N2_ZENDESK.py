#!/usr/bin/env python
# coding: utf-8

# In[6]:


# ===========================================
# Automação Zendesk (versão Jupyter - full)
# v3 — Ajuste: bloco de envio e captura de ID aprimorado
# ===========================================

# ---- Dependências: instala se faltar (sem warnings de depreciação) ----
import sys, subprocess, importlib.util
def ensure_installed(*pkgs):
    for p in pkgs:
        if importlib.util.find_spec(p) is None:
            print(f"Instalando {p}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", p])
    print("✓ Dependências checadas.")

ensure_installed("pandas", "xlwings", "selenium", "python-dateutil", "openpyxl")

# ---- Imports principais ----
import os
import time
import re
import getpass
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, ElementNotInteractableException
from datetime import datetime, timedelta
from dateutil import parser

# ================================================================
# CONFIGURAÇÕES GERAIS
# ================================================================
LOGIN_FIXO = "126815@pmenos.com.br"
EXCEL_PATH = r"C:\Users\126815\OneDrive - paguemenos.com.br\ENGENHARIA - OBRAS - Documentos\BANCO DE DADOS - POWER BI\BI - OUTROS\BASE CONTROLE DE PAGAMENTOS.xlsx"
FORM_URL   = "https://portaldeservicos.pmenos.com.br/hc/pt-br/requests/new?ticket_form_id=41684359104269"
BASE_PDF_DIR = r"C:\Users\126815\OneDrive - paguemenos.com.br\Área de Trabalho\NFS\2025"

# Caminho para salvar o perfil do Chrome (evita pedir login/MFA toda vez)
CHROME_PROFILE_PATH = os.path.join(os.getcwd(), "ChromeProfile")
os.makedirs(CHROME_PROFILE_PATH, exist_ok=True)

# ---- Fail-safe de caminho do Excel para evitar erro silencioso ----
if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(
        "❌ Planilha não encontrada no caminho:\n"
        f"   {EXCEL_PATH}\n"
        "💡 Verifique se o OneDrive sincronizou e está 'Disponível neste dispositivo'."
    )

# ================================================================
# FUNÇÕES UTILITÁRIAS — localizar e preencher por label
# ================================================================
def _normalize_text(txt: str) -> str:
    return re.sub(r'\s+', ' ', (txt or '').strip()).lower()

def find_input_by_label(driver, label_candidates, timeout=10):
    wait = WebDriverWait(driver, timeout)
    norm_targets = [_normalize_text(l) for l in label_candidates]

    # 1) label[for] -> id
    try:
        labels = driver.find_elements(By.TAG_NAME, "label")
        for lbl in labels:
            try:
                lbl_text = _normalize_text(lbl.text)
                if any(lbl_text.startswith(t) for t in norm_targets):
                    for_attr = lbl.get_attribute("for")
                    if for_attr:
                        try:
                            elem = driver.find_element(By.ID, for_attr)
                            return elem
                        except NoSuchElementException:
                            pass
            except Exception:
                continue
    except Exception:
        pass

    # 2) aria-label / name / placeholder
    for cand in label_candidates:
        norm = _normalize_text(cand)
        try:
            elems = driver.find_elements(By.CSS_SELECTOR, '[aria-label]')
            for e in elems:
                if _normalize_text(e.get_attribute('aria-label')).startswith(norm):
                    return e
        except Exception:
            pass
        try:
            elems = driver.find_elements(By.CSS_SELECTOR, '[name]')
            for e in elems:
                nm = _normalize_text(e.get_attribute('name'))
                if nm in [norm, norm.replace(' (opcional)', '')]:
                    return e
        except Exception:
            pass
        try:
            elems = driver.find_elements(By.CSS_SELECTOR, '[placeholder]')
            for e in elems:
                if _normalize_text(e.get_attribute('placeholder')).startswith(norm):
                    return e
        except Exception:
            pass

    # 3) Fallback: campos custom do Zendesk
    try:
        candidates = driver.find_elements(By.CSS_SELECTOR, 'input[id^="request_custom_fields_"], textarea[id^="request_custom_fields_"], select[id^="request_custom_fields_"]')
        for elem in candidates:
            try:
                if elem.is_displayed() and elem.is_enabled():
                    tag = elem.tag_name.lower()
                    if tag in ('input', 'textarea', 'select'):
                        try:
                            parent = elem.find_element(By.XPATH, './ancestor::div[1]')
                            maybe_label = parent.find_element(By.TAG_NAME, 'label')
                            text = _normalize_text(maybe_label.text)
                            if any(text.startswith(t) for t in norm_targets):
                                return elem
                        except Exception:
                            continue
            except Exception:
                continue
    except Exception:
        pass

    return None

def fill_text_or_select(driver, element, value: str):
    if not value:
        return
    tag = element.tag_name.lower()
    try:
        if tag == 'select':
            try:
                Select(element).select_by_visible_text(value)
            except Exception:
                try:
                    Select(element).select_by_value(value)
                except Exception:
                    element.click()
                    time.sleep(0.3)
                    opts = driver.find_elements(By.XPATH, f'//li[contains(@class,"select2-results__option") and normalize-space()="{value}"] | //div[contains(@class,"option") and normalize-space()="{value}"]')
                    if opts:
                        opts[0].click()
        elif tag in ('input', 'textarea'):
            try:
                element.clear()
            except Exception:
                pass
            element.send_keys(value)
        else:
            try:
                element.click()
                element.send_keys(value)
            except Exception:
                pass
    except (ElementNotInteractableException, NoSuchElementException):
        pass

def preencher_fornecedor(driver, fornecedor_valor: str, timeout=12):
    fornecedor_valor = (fornecedor_valor or "").strip()
    if not fornecedor_valor:
        return

    label_candidates = ["Fornecedor", "Fornecedor (opcional)"]
    elem = find_input_by_label(driver, label_candidates, timeout=timeout)

    if elem is None:
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input, textarea, select')))
            elem = find_input_by_label(driver, label_candidates, timeout=5)
        except TimeoutException:
            elem = None

    if elem:
        fill_text_or_select(driver, elem, fornecedor_valor)
    else:
        print("⚠️ Campo 'Fornecedor' não encontrado. Seguindo sem preencher.")

# ================================================================
# FUNÇÕES DE LOGIN
# ================================================================
def inicializar_driver():
    """Inicializa o Chrome com parâmetros de estabilidade, persistência de perfil e modo Headless."""
    chrome_options = Options()

    # 👉 DICA: na PRIMEIRA execução, comente a linha abaixo para autenticar (MFA) com interface:
    #chrome_options.add_argument("--headless=new")  # comente esta linha na 1ª vez, depois reative

    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    # ✅ FIX: Desativa AI Mode, AI Overviews e NTP overlay do Chrome
    chrome_options.add_argument("--disable-features=AIMode,SearchAIMode,AIOverviews,SearchGenerativeExperience,NtpRealbox,NtpMiddleSlotPromo,NewTabPageFeatures")
    chrome_options.add_argument("--disable-search-engine-choice-screen")

    # ✅ FIX: Suprime popup "Restaurar páginas?" quando Chrome foi fechado abruptamente
    chrome_options.add_argument("--disable-session-crashed-bubble")
    chrome_options.add_argument("--hide-crash-restore-bubble")
    chrome_options.add_argument("--restore-last-session=false")
    chrome_options.add_argument("--no-first-run")

    chrome_options.add_argument(f"--user-data-dir={CHROME_PROFILE_PATH}")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    )

    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    prefs = {
        "download.default_directory": "C:/temp",
        "download.prompt_for_download": False,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_settings.popups": 0,
        "homepage": "about:blank",
        "homepage_is_newtabpage": False,
        "session.startup_urls": [FORM_URL],
        "browser.startup_page": 4,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(45)
        driver.implicitly_wait(10)
        driver.get(FORM_URL)

        # ✅ FIX: Chrome 147 + Modo IA abre aba Google que trava ao fechar
        time.sleep(2)
        handles = driver.window_handles
        if len(handles) > 1:
            print(f"⚠️ {len(handles) - 1} aba(s) extra(s) do Modo IA. Redirecionando...")
            for handle in handles[1:]:
                try:
                    driver.switch_to.window(handle)
                    driver.get("about:blank")
                except Exception:
                    pass
            driver.switch_to.window(handles[0])
        if FORM_URL not in driver.current_url:
            print(f"⚠️ URL incorreta ({driver.current_url}). Redirecionando para FORM_URL...")
            driver.get(FORM_URL)
            time.sleep(3)

        print("✅ Driver Chrome inicializado com sucesso", "(Modo Headless)" if "--headless=new" in chrome_options.arguments else "(Com interface)")
        return driver
    except Exception as e:
        print(f"❌ Erro ao inicializar Chrome: {e}")
        print("⚠️ Certifique-se de que não há outras janelas do Chrome abertas pela automação.")
        raise

def realizar_login(driver, login_fixo):
    """Verifica se já está logado ou realiza o login no portal."""
    print("\n" + "="*64)
    print("                 VERIFICANDO ESTADO DE LOGIN")
    print("="*64)

    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "request_subject")))
        print("✅ Sessão já ativa. Login ignorado.")
        return driver
    except:
        print("ℹ️ Sessão não encontrada ou expirada. Iniciando login...")

    senha = getpass.getpass(f"🔑 Digite a senha para {login_fixo}: ")

    # 1. Inserir Login
    try:
        print("Buscando campo de login...")
        seletores_login = [(By.ID, "identifierId"), (By.NAME, "loginfmt"), (By.ID, "i0116"), (By.CSS_SELECTOR, "input[type='email']")]
        campo_login = None
        for by, selector in seletores_login:
            try:
                campo_login = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, selector)))
                if campo_login: break
            except:
                continue
        if campo_login:
            campo_login.clear()
            campo_login.send_keys(login_fixo)
            campo_login.send_keys(Keys.ENTER)
            print("✅ Login inserido.")
            time.sleep(2)
    except Exception as e:
        print(f"❌ Erro na etapa de login: {str(e)[:100]}")

    # 2. Inserir Senha
    try:
        print("Buscando campo de senha...")
        seletores_senha = [(By.NAME, "password"), (By.NAME, "passwd"), (By.ID, "i0118"), (By.CSS_SELECTOR, "input[type='password']")]
        campo_senha = None
        for by, selector in seletores_senha:
            try:
                campo_senha = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by, selector)))
                if campo_senha: break
            except:
                continue
        if campo_senha:
            campo_senha.send_keys(senha)
            time.sleep(1)
            campo_senha.send_keys(Keys.ENTER)
            print("✅ Senha inserida.")
            time.sleep(3)
    except Exception as e:
        print(f"❌ Erro na etapa de senha: {str(e)[:100]}")

    # 3. Inserir Token Microsoft (MFA/2FA)
    try:
        print("Buscando campo de Token/MFA...")
        seletores_token = [
            (By.ID, "idTxtBx_SAOTCC_OTC"),
            (By.NAME, "otc"),
            (By.CSS_SELECTOR, "input[aria-label='Código de verificação']"),
            (By.ID, "idTxtBx_SAOTAS_OTC"),
        ]
        campo_token = None
        start_time = time.time()
        while time.time() - start_time < 20:
            for by, selector in seletores_token:
                try:
                    campo_token = driver.find_element(by, selector)
                    if campo_token.is_displayed():
                        break
                except:
                    continue
            if campo_token:
                break
            try:
                msg_app = driver.find_element(By.ID, "idDiv_SAOTCAS_Title")
                if "Aprovar" in msg_app.text or "App" in msg_app.text:
                    print("📱 Verificação por Aplicativo detectada. Aprove no seu celular...")
                    break
            except:
                pass
            time.sleep(1)

        if campo_token:
            token = input("🔒 Digite o Token Microsoft (MFA/2FA) e pressione Enter: ")
            campo_token.send_keys(token)
            time.sleep(1)
            campo_token.send_keys(Keys.ENTER)
            print("✅ Token inserido.")
        else:
            print("ℹ️ Campo de Token não encontrado. Verifique se o login já concluiu.")
    except Exception as e:
        print(f"❌ Erro na etapa de token: {str(e)[:100]}")

    # 4. Mantenha-me conectado (Continuar conectado?)
    print("Buscando tela 'Continuar conectado?'...")
    try:
        seletores_sim = [(By.ID, "idSIButton9"), (By.CSS_SELECTOR, "input[type='submit'][value='Sim']"), (By.XPATH, "//input[@value='Sim']")]
        btn_sim = None
        start_time = time.time()
        while time.time() - start_time < 15:
            for by, selector in seletores_sim:
                try:
                    btn_sim = driver.find_element(by, selector)
                    if btn_sim.is_displayed() and btn_sim.is_enabled():
                        break
                except:
                    continue
            if btn_sim:
                break
            time.sleep(1)
        if btn_sim:
            try:
                btn_sim.click()
            except:
                driver.execute_script("arguments[0].click();", btn_sim)
            print("✅ Botão 'Sim' clicado na tela 'Continuar conectado?'.")
    except:
        pass

    print("="*64)
    print("             PROCESSO DE LOGIN CONCLUÍDO")
    print("="*64)
    time.sleep(5)
    return driver

# ================================================================
# LÓGICA DE ABERTURA DE CHAMADO (SEQUÊNCIA DEFINITIVA)
# ================================================================
# ✅ FIX: `excel_row` agora é passado diretamente do main() já calculado
#         e usado sem somar +2 aqui dentro — evita desalinhamento de linha.
def abrir_chamado(excel_row, row_data, driver, wb, ws):
    assunto = str(row_data["ASSUNTO"]).strip()
    pedido_raw = row_data["PEDIDO"]
    pedido = str(int(float(pedido_raw))).strip() if pd.notna(pedido_raw) and str(pedido_raw).replace('.', '', 1).isdigit() else str(pedido_raw).strip()

    valor_raw = str(row_data["VALOR BI"]).strip()
    try:
        valor_limpo = re.sub(r"[^\d.,]", "", valor_raw)
        if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", valor_limpo):
            valor_formatado = valor_limpo
        else:
            valor_num = float(valor_limpo.replace(',', '.'))
            valor_formatado = f"{valor_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        valor_formatado = "0,00"

    # Lógica de Vencimento
    venc_match = re.search(r"VENC\s*(\d{2}/\d{2}/\d{4})", assunto)
    data_vencimento_raw = venc_match.group(1) if venc_match else None

    if data_vencimento_raw:
        try:
            data_obj = parser.parse(data_vencimento_raw, dayfirst=True)
            data_vencimento = data_obj.strftime("%d/%m/%Y")
        except:
            data_vencimento = (datetime.now() + timedelta(days=7)).strftime("%d/%m/%Y")
    else:
        data_vencimento = (datetime.now() + timedelta(days=7)).strftime("%d/%m/%Y")

    cd_match = re.search(r"CD\s*([A-Z0-9]{2,4})", assunto, re.IGNORECASE)
    if cd_match:
        loja = f"CD {cd_match.group(1).strip()}"
        projeto_match = re.match(r"(.+?)\s*CD\s*[A-Z0-9]{2,4}", assunto, re.IGNORECASE)
        projeto = projeto_match.group(1).strip() if projeto_match else ""
    else:
        projeto_loja_match = re.match(r"([A-Z\s]+)\s+(\d{3,4})", assunto)
        projeto = projeto_loja_match.group(1).strip() if projeto_loja_match else ""
        loja = projeto_loja_match.group(2).strip() if projeto_loja_match else ""

    if not projeto:
        projeto_fallback = re.match(r"(.+?)(?=\s*-\s*|\s*N[°º])", assunto, re.IGNORECASE)
        projeto = projeto_fallback.group(1).strip() if projeto_fallback else ""

    fornecedor_match = re.search(r"-\s*([A-ZÀ-Úa-z0-9 ]+?)(?=\s*-|\s*N[°º])", assunto)
    fornecedor = fornecedor_match.group(1).strip() if fornecedor_match else ""

    nf_match = re.search(r"N[°º]?\s*(\d+)", assunto)
    nf = nf_match.group(1) if nf_match else ""

    descricao_final = f"{assunto} | PEDIDO: {pedido}"

    if loja.startswith("CD "):
        pdf_path = os.path.join(BASE_PDF_DIR, projeto, fornecedor, f"NF {nf} - {loja}.pdf")
    else:
        pdf_path = os.path.join(BASE_PDF_DIR, projeto, fornecedor, f"NF {nf} - LJ {loja}.pdf")
    pdf_existe = os.path.exists(pdf_path)

    fornecedor_base = ""
    try:
        if hasattr(row_data, "index") and "FORNECEDOR" in row_data.index and pd.notna(row_data["FORNECEDOR"]):
            fornecedor_base = str(row_data["FORNECEDOR"]).strip()
    except Exception:
        fornecedor_base = ""
    fornecedor_final = fornecedor_base or fornecedor

    print(f"\n{'='*60}")
    print(f"🟦 PROCESSANDO LINHA EXCEL {excel_row}")
    print(f"{'='*60}")
    print(f"📄 NF: {nf} | 🏪 Loja: {loja} | 🏗️ Projeto: {projeto}")
    print(f"🏢 Fornecedor: {fornecedor_final} | 💰 Valor: {valor_formatado}")
    print(f"📁 PDF: {'✅ Encontrado' if pdf_existe else '❌ Não encontrado — chamado será aberto sem anexo'}")
    print(f"{'='*60}\n")

    try:
        time.sleep(3)
        actions = ActionChains(driver)

        # 1. NAVEGAÇÃO INICIAL (Tab, Espaço, 7 Baixo)
        print("⏳ Passo 1: Navegação inicial (7x Baixo)...")
        actions.send_keys(Keys.TAB).perform(); time.sleep(0.5)
        actions.send_keys(Keys.SPACE).perform(); time.sleep(0.5)
        for _ in range(7):
            actions.send_keys(Keys.ARROW_DOWN).perform(); time.sleep(0.1)
        actions.send_keys(Keys.ENTER).perform(); time.sleep(1)

        # 2. DROPDOWN 1 (Tab, Espaço, 1 Baixo)
        print("⏳ Passo 2: Dropdown 1 (1x Baixo)...")
        actions.send_keys(Keys.TAB).perform(); time.sleep(0.3)
        actions.send_keys(Keys.SPACE).perform(); time.sleep(0.5)
        actions.send_keys(Keys.ARROW_DOWN).perform(); time.sleep(0.1)
        actions.send_keys(Keys.ENTER).perform(); time.sleep(0.5)

        # 3. FORMA DE PAGAMENTO (Tab, Espaço, 2 Baixo)
        print("⏳ Passo 6: Forma de Pagamento (2x Baixo)...")
        actions.send_keys(Keys.TAB).perform(); time.sleep(0.3)
        actions.send_keys(Keys.SPACE).perform(); time.sleep(0.5)
        for _ in range(2):
            actions.send_keys(Keys.ARROW_DOWN).perform(); time.sleep(0.1)
        actions.send_keys(Keys.ENTER).perform(); time.sleep(0.5)

        # 4. PEDIDO DE COMPRA
        print("⏳ Passo 7: Pedido de Compra...")
        actions.send_keys(Keys.TAB).send_keys(pedido).perform(); time.sleep(0.5)

        # 5. VALOR
        print("⏳ Passo 5: Valor...")
        actions.send_keys(Keys.TAB).send_keys(valor_formatado).perform(); time.sleep(0.5)

        # 6. DATA DE VENCIMENTO
        print("⏳ Passo 4: Data de Vencimento...")
        actions.send_keys(Keys.TAB).send_keys(data_vencimento).send_keys(Keys.ENTER).perform(); time.sleep(0.5)

        # 6.1 FORNECEDOR
        print("⏳ Passo 4.1: Preenchendo Fornecedor (opcional)...")
        try:
            preencher_fornecedor(driver, fornecedor_final, timeout=12)
        except Exception as e:
            print(f"⚠️ Erro ao preencher 'Fornecedor': {e}")

        # 7. DROPDOWN 2 (Tab, 2 Baixo)
        print("⏳ Passo 3: Dropdown 2 (2x Baixo)...")
        actions.send_keys(Keys.TAB).perform(); time.sleep(0.3)
        for _ in range(2):
            actions.send_keys(Keys.ARROW_DOWN).perform(); time.sleep(0.1)
        actions.send_keys(Keys.ENTER).perform(); time.sleep(0.5)

        # 8. ASSUNTO
        print("⏳ Passo 9: Preenchendo Assunto...")
        campo_assunto = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "request_subject")))
        campo_assunto.clear(); campo_assunto.send_keys(assunto)

        # 9. DESCRIÇÃO
        print("⏳ Passo 10: Preenchendo Descrição...")
        campo_descricao = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "request_description")))
        campo_descricao.clear(); campo_descricao.send_keys(descricao_final)

        # 10. CHECKBOX ANEXO NF
        print("⏳ Passo 8: Flagando Checkbox...")
        actions.send_keys(Keys.TAB).perform(); time.sleep(0.3)
        actions.send_keys(Keys.SPACE).perform(); time.sleep(0.5)

        # 11. ANEXOS
        if pdf_existe:
            print("⏳ Passo 11: Anexando PDF...")
            attachment_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "request-attachments")))
            driver.execute_script("arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible'; arguments[0].removeAttribute('hidden');", attachment_input)
            attachment_input.send_keys(pdf_path)
            time.sleep(5)
        else:
            print("⚠️ PDF não encontrado — seguindo sem anexo.")

        # ================================================================
        # ✅ v3 — BLOCO DE ENVIO E CAPTURA DE ID (SUBSTITUÍDO)
        # ================================================================
        print("🚀 Enviando formulário...")
        url_antes_envio = driver.current_url
        envio_sucesso = False

        for metodo in range(1, 4):
            try:
                if metodo == 1:
                    btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    btn.click()
                elif metodo == 2:
                    btn = driver.find_element(By.XPATH, "//input[@type='submit']")
                    driver.execute_script("arguments[0].click();", btn)
                elif metodo == 3:
                    ActionChains(driver).send_keys(Keys.TAB).send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
                envio_sucesso = True
                print(f"   ✅ Enviado (método {metodo})")
                break
            except Exception as e:
                print(f"   ⚠️ Método {metodo} falhou: {str(e)[:80]}")
                continue

        if not envio_sucesso:
            raise Exception("❌ Não conseguiu acionar o botão de envio!")

        # ✅ Log imediato da URL para debug
        print(f"🔍 URL imediatamente após click submit: {driver.current_url}")
        time.sleep(2)  # Aguarda processamento inicial do submit
        print(f"🔍 URL após 2s: {driver.current_url}")

        # ✅ Aguarda redirecionamento com padrões múltiplos + detecção de erro de validação
        print("⏳ Aguardando confirmação de envio...")

        def url_mudou_ou_erro(d):
            """Retorna True se a URL mudou (sucesso) ou levanta exceção se houver erro de validação."""
            url_atual = d.current_url
            if "/requests/" in url_atual and url_atual != url_antes_envio:
                return True
            try:
                erros = d.find_elements(
                    By.CSS_SELECTOR,
                    ".form-field.is-error, .notification-error, [class*='error']:not([style*='display: none'])"
                )
                if erros:
                    msgs = [e.text.strip() for e in erros if e.text.strip()]
                    if msgs:
                        raise Exception(f"Erro de validação no formulário: {'; '.join(msgs[:3])}")
            except Exception as ex:
                if "validação" in str(ex):
                    raise  # Repropaga erro de validação
            return False

        try:
            WebDriverWait(driver, 20).until(url_mudou_ou_erro)
        except TimeoutException:
            # Última tentativa: verifica URL manualmente
            url_final_check = driver.current_url
            print(f"⚠️ Timeout na espera. URL atual: {url_final_check}")
            if "/requests/" in url_final_check and url_final_check != url_antes_envio:
                print("✅ Redirecionamento detectado na verificação final.")
            else:
                # Tenta capturar mensagens de erro da página
                try:
                    page_errors = driver.find_elements(
                        By.XPATH,
                        "//*[contains(@class,'error') or contains(@class,'Error')]"
                    )
                    erros_visiveis = [e.text for e in page_errors if e.text.strip() and e.is_displayed()]
                    if erros_visiveis:
                        raise Exception(
                            f"❌ Formulário não foi aceito. Erros encontrados: {erros_visiveis[:3]}\n"
                            f"URL: {url_final_check}"
                        )
                except Exception as inner:
                    if "❌" in str(inner):
                        raise
                raise Exception(
                    f"❌ Timeout: formulário pode não ter sido enviado.\n"
                    f"URL antes: {url_antes_envio}\n"
                    f"URL atual: {url_final_check}"
                )

        link_final = driver.current_url
        print(f"🔍 URL após redirecionamento: {link_final}")

        # ✅ Padrão de extração de ID mais amplo — cobre /requests/123 e /hc/*/requests/123
        id_match = re.search(r"/requests/(\d+)", link_final)
        if not id_match:
            # Fallback: tenta extrair da página (título ou heading do ticket)
            try:
                heading = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//h1[contains(text(),'#')] | //*[@class='request-title']")
                    )
                )
                id_match_page = re.search(r"#(\d+)", heading.text)
                id_extraido = id_match_page.group(1) if id_match_page else "ID_NAO_ENCONTRADO"
            except Exception:
                id_extraido = "ID_NAO_ENCONTRADO"
            print(f"⚠️ ID não encontrado na URL. Tentativa via página: {id_extraido}")
        else:
            id_extraido = id_match.group(1)

        print(f"✅ SUCESSO! ID: {id_extraido}")
        print(f"🔗 LINK DO TICKET: {link_final}")
        # ================================================================
        # FIM DO BLOCO v3
        # ================================================================

        # ✅ FIX: usa excel_row diretamente (já calculado no main com idx + 2)
        ws.range(f"V{excel_row}").value = id_extraido
        cell = ws.range(f"W{excel_row}")
        cell.value = "Abrir chamado"
        cell.api.Hyperlinks.Add(Anchor=cell.api, Address=link_final, TextToDisplay="Abrir chamado")
        ws.range(f"AE{excel_row}").value = "Feito pela automação"
        ws.range(f"AM{excel_row}").value = datetime.now().date()

        # ✅ FIX: wb.save() com try/except explícito para não perder dados silenciosamente
        try:
            wb.save()
            print(f"💾 Planilha salva — linha {excel_row}")
        except Exception as e:
            print(f"⚠️ FALHA AO SALVAR planilha na linha {excel_row}: {e}")
            print("💡 Verifique se o arquivo está aberto em outro processo ou bloqueado pelo OneDrive.")

        driver.get(FORM_URL)

        # ✅ FIX: mesma proteção contra AI Mode na renavegação entre tickets
        time.sleep(2)
        handles = driver.window_handles
        if len(handles) > 1:
            for handle in handles[1:]:
                try:
                    driver.switch_to.window(handle)
                    driver.get("about:blank")
                except Exception:
                    pass
            driver.switch_to.window(handles[0])
        if FORM_URL not in driver.current_url:
            print(f"⚠️ URL incorreta após renavegação ({driver.current_url}). Forçando FORM_URL...")
            driver.get(FORM_URL)
            time.sleep(3)

        return True, driver

    except Exception as e:
        print(f"❌ ERRO NA LINHA EXCEL {excel_row}: {str(e)[:200]}")
        return False, None

# ================================================================
# FLUXO PRINCIPAL
# ================================================================
def main():
    print("🚀 Iniciando automação com Chrome...")
    driver = inicializar_driver()
    driver = realizar_login(driver, LOGIN_FIXO)

    print("📊 Abrindo planilha Excel...")
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = None
    try:
        wb = app.books.open(EXCEL_PATH)
        ws = wb.sheets['SOLICITAÇÃO DE PAGAMENTO']
        ultima_linha = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
        df = ws.range(f"A1:AF{ultima_linha}").options(pd.DataFrame, header=1, index=False).value

        linhas_filtradas = df[df["Abrir ticket?"].astype(str).str.lower() == "sim"]
        total = len(linhas_filtradas)
        print(f"📊 Total de linhas a processar: {total}\n")

        sucesso_count = 0
        for i, (idx, row) in enumerate(linhas_filtradas.iterrows(), 1):
            # ✅ FIX: calcula a linha real do Excel aqui e passa para abrir_chamado
            # idx é 0-based (índice do DataFrame), +1 pelo header, +1 pelo Excel ser 1-based
            excel_row = idx + 2

            print(f"{'#'*60}\n📍 PROCESSANDO {i}/{total} — Linha Excel: {excel_row}\n{'#'*60}")
            sucesso, driver_resultado = abrir_chamado(excel_row, row, driver, wb, ws)

            if sucesso:
                driver = driver_resultado
                sucesso_count += 1
            else:
                print("🔄 Reiniciando driver após falha...")
                try:
                    driver.quit()
                except:
                    pass
                time.sleep(3)
                driver = inicializar_driver()
                driver = realizar_login(driver, LOGIN_FIXO)

        print(f"\n{'='*60}\n📊 RESUMO FINAL\n{'='*60}")
        print(f"✅ Sucessos: {sucesso_count} | ❌ Falhas: {total - sucesso_count}")

    finally:
        try:
            if wb: wb.close()
            app.quit()
            driver.quit()
        except:
            pass
    print("\n✅ AUTOMAÇÃO FINALIZADA!")

# ---- Execução imediata (em Jupyter) ----
main()


# In[ ]:




