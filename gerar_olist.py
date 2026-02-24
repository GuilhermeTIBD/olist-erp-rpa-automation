import os
import shutil
import time
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    UnexpectedAlertPresentException,
    NoAlertPresentException,
    TimeoutException
)

# =====================================================
# CONFIG
# =====================================================

ARQUIVO_ENTRADA = "_GESTAO_FINANCEIRA_SHOPEE_NOV.2025.xlsx"
ARQUIVO_SAIDA = "resultado_olist.xlsx"

COLUNA_CODIGO = "ID do pedido"
COLUNA_DATA = "Data"
COLUNA_TOTAL_TAXAS = "TOTAL TAXAS"
COLUNA_FRETE_COBRADO = "Frete cobrado do comprador"
COLUNA_VALOR_LIQUIDO = "VALOR LIQUIDO"
COLUNA_VALIDACAO = "VALIDAÇÃO"   # coluna S
COLUNA_STATUS = "BAIXADO"        # gravar "SIM"

URL = "https://erp.olist.com/contas_receber"

WAIT_TIMEOUT = 20
SALVAR_A_CADA = 2000              # ✅checkpoint
TENTATIVAS_POR_PEDIDO = 1

PASTA_DEBUG = "debug"

# Busca
SEARCH_INPUT_ID = "pesquisa-mini"
SEARCH_BUTTON_CSS = "span.input-group-btn > button.btn.btn-default"

# Tabela
RESULT_ROW_SELECTOR = "table tbody tr"
NO_RESULTS_XPATH = "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nenhum')]"

# Perfil Chrome
CHROME_USER_DATA_ORIGINAL = r"C:\Users\DELL\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE_ORIGINAL = "Default"
CHROME_USER_DATA_CLONE = r"C:\olist_profile_selenium"
CHROME_PROFILE_CLONE = "Default"

# Selects do Olist
CONTA_SHOPEE_VALUE = "737401193"
CONTA_SHOPEE_TEXTO = "SHOPEE"
CATEGORIA_SHOPEE_TEXTO = "RECEITA SHOPEE"

# Inputs
INPUT_TAXA_ID = "taxa0"
INPUT_DESCONTO_ID = "desconto0"
INPUT_VALOR_ID = "valor0"

# Botão final
BTN_SALVAR_BORDERO_ID = "salvarBordero"

# =====================================================
# HELPERS
# =====================================================

def garantir_pasta(path: str):
    if not os.path.exists(path):
        os.makedirs(path)

def screenshot(driver, nome: str):
    garantir_pasta(PASTA_DEBUG)
    try:
        driver.save_screenshot(os.path.join(PASTA_DEBUG, nome))
    except Exception:
        pass

def salvar_html_debug(driver, nome: str):
    garantir_pasta(PASTA_DEBUG)
    try:
        html = driver.page_source
        with open(os.path.join(PASTA_DEBUG, nome), "w", encoding="utf-8") as f:
            f.write(html)
    except Exception:
        pass

def click_js(driver, element):
    driver.execute_script("arguments[0].click();", element)

def fechar_alerta_se_existir(driver) -> str | None:
    try:
        alert = driver.switch_to.alert
        texto = alert.text
        alert.accept()
        return texto
    except NoAlertPresentException:
        return None
    except Exception:
        return None

def sessao_expirada(texto_alerta: str) -> bool:
    if not texto_alerta:
        return False
    t = texto_alerta.lower()
    return ("sessão expirou" in t) or ("sessao expirou" in t) or ("login em outra máquina" in t)

def clonar_perfil():
    print("\n🔁 Clonando perfil do Chrome para uso do robô...")
    garantir_pasta(CHROME_USER_DATA_CLONE)

    src_profile = os.path.join(CHROME_USER_DATA_ORIGINAL, CHROME_PROFILE_ORIGINAL)
    dst_profile = os.path.join(CHROME_USER_DATA_CLONE, CHROME_PROFILE_CLONE)

    if os.path.exists(dst_profile):
        shutil.rmtree(dst_profile, ignore_errors=True)

    shutil.copytree(src_profile, dst_profile)

    src_local_state = os.path.join(CHROME_USER_DATA_ORIGINAL, "Local State")
    dst_local_state = os.path.join(CHROME_USER_DATA_CLONE, "Local State")
    if os.path.exists(src_local_state):
        shutil.copy2(src_local_state, dst_local_state)

    print("✅ Perfil clonado em:", CHROME_USER_DATA_CLONE)

def formatar_data_br(valor) -> str | None:
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%d/%m/%Y")
    s = str(valor).strip()
    if not s:
        return None
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.isna(dt):
        return None
    return dt.strftime("%d/%m/%Y")

def br_money(valor) -> str | None:
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return f"{float(valor):.2f}".replace(".", ",")
    s = str(valor).strip()
    if not s:
        return None
    s = s.replace("R$", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "")
    elif "." in s and "," not in s:
        s = s.replace(".", ",")
    try:
        v = float(s.replace(".", "").replace(",", "."))
        return f"{v:.2f}".replace(".", ",")
    except Exception:
        return s

def valor_num(valor) -> float:
    if valor is None or pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    if not s:
        return 0.0
    s = s.replace("R$", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

# =====================================================
# ACHAR CAMPO DE BUSCA (ROBUSTO)
# =====================================================

BUSCA_SELECTORS = [
    (By.ID, "pesquisa-mini"),
    (By.CSS_SELECTOR, "input#pesquisa-mini"),
    (By.NAME, "pesquisa-mini"),
    (By.CSS_SELECTOR, "input[name='pesquisa-mini']"),
    (By.CSS_SELECTOR, "input[type='search']"),
    (By.CSS_SELECTOR, "input[placeholder*='Pesquisar']"),
    (By.CSS_SELECTOR, "input[placeholder*='pesquis']"),
    (By.CSS_SELECTOR, "input.form-control"),
    (By.XPATH, "//input[contains(@id,'pesquisa') or contains(@name,'pesquisa')]"),
]

def achar_input_busca(driver, timeout=25):
    w = WebDriverWait(driver, timeout)

    try:
        w.until(lambda d: d.execute_script("return document.readyState") in ("interactive", "complete"))
    except Exception:
        pass

    for by, sel in BUSCA_SELECTORS:
        try:
            els = driver.find_elements(by, sel)
            if els:
                for e in els:
                    try:
                        if e.is_displayed() and e.is_enabled():
                            return e
                    except Exception:
                        continue
        except Exception:
            continue

    try:
        return w.until(EC.element_to_be_clickable((By.ID, SEARCH_INPUT_ID)))
    except Exception:
        pass

    screenshot(driver, "nao_achei_busca.png")
    salvar_html_debug(driver, "nao_achei_busca.html")
    raise RuntimeError("Ainda não achei o campo de busca. Salvei debug/nao_achei_busca.png e .html")

def achar_botao_lupa(driver, timeout=20):
    w = WebDriverWait(driver, timeout)
    try:
        return w.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEARCH_BUTTON_CSS)))
    except Exception:
        pass

    for css in [
        "button.btn.btn-default",
        "span.input-group-btn button",
        "button[type='submit']",
    ]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, css)
            for e in els:
                try:
                    if e.is_displayed() and e.is_enabled():
                        return e
                except Exception:
                    continue
        except Exception:
            continue

    screenshot(driver, "nao_achei_lupa.png")
    salvar_html_debug(driver, "nao_achei_lupa.html")
    raise RuntimeError("Não achei o botão da lupa. Salvei debug/nao_achei_lupa.png e .html")

def garantir_na_tela_contas_receber(driver):
    driver.get(URL)
    time.sleep(2)
    fechar_alerta_se_existir(driver)

    try:
        _ = achar_input_busca(driver, timeout=15)
        return
    except Exception:
        pass

    screenshot(driver, "antes_pedir_navegar.png")
    input("\n👉 Não achei o campo de busca. No Chrome do robô, clique em Contas a Receber (nessa tela) e aperte ENTER... ")
    time.sleep(1)

    try:
        _ = achar_input_busca(driver, timeout=25)
        return
    except Exception:
        driver.get(URL)
        time.sleep(2)
        fechar_alerta_se_existir(driver)
        _ = achar_input_busca(driver, timeout=25)
        return

# =====================================================
# CONFIRMAR SE BAIXOU (EVITA ERRO FALSO)
# =====================================================

def confirmar_se_baixou(driver, codigo: str, timeout=20) -> bool:
    """
    Volta pra lista, busca o pedido e tenta detectar se já está baixado.
    Ajuste as palavras-chave se no seu Olist aparecer diferente.
    """
    garantir_na_tela_contas_receber(driver)

    busca = achar_input_busca(driver, timeout=timeout)
    busca.click()
    busca.send_keys(Keys.CONTROL, "a")
    busca.send_keys(Keys.DELETE)
    time.sleep(0.05)
    busca.send_keys(codigo)

    lupa = achar_botao_lupa(driver, timeout=timeout)
    click_js(driver, lupa)

    resp = esperar_resultado_da_busca(driver, codigo, timeout=timeout)
    if resp != "OK":
        return False

    rows = driver.find_elements(By.CSS_SELECTOR, RESULT_ROW_SELECTOR)
    for r in rows:
        try:
            if codigo in r.text:
                txt = r.text.lower()
                # ✅ sinais comuns de "já baixado"
                if ("baixad" in txt) or ("recebid" in txt) or ("pago" in txt) or ("liquid" in txt):
                    return True
                return False
        except Exception:
            continue

    return False

# =====================================================
# OLIST ACTIONS
# =====================================================

def esperar_resultado_da_busca(driver, codigo: str, timeout=35) -> str:
    t0 = time.time()
    while time.time() - t0 < timeout:
        txt = fechar_alerta_se_existir(driver)
        if txt and sessao_expirada(txt):
            return "RELOGAR"

        if driver.find_elements(By.XPATH, NO_RESULTS_XPATH):
            return "NAO_ENCONTRADO"

        rows = driver.find_elements(By.CSS_SELECTOR, RESULT_ROW_SELECTOR)
        for r in rows:
            try:
                if codigo in r.text:
                    return "OK"
            except Exception:
                pass

        time.sleep(0.3)

    return "TIMEOUT"

def clicar_navigate_da_linha(driver, codigo: str):
    rows = driver.find_elements(By.CSS_SELECTOR, RESULT_ROW_SELECTOR)
    for r in rows:
        try:
            if codigo in r.text:
                btn = r.find_element(By.CSS_SELECTOR, "button.button-navigate")
                click_js(driver, btn)
                return
        except Exception:
            continue
    raise RuntimeError("Não encontrei a linha do pedido (ou button-navigate) para este código.")

def clicar_receber_baixar(driver, timeout=35):
    w = WebDriverWait(driver, timeout)
    el = w.until(EC.element_to_be_clickable((
        By.XPATH,
        "//a[.//i[contains(@class,'fa-check')] and contains(normalize-space(.),'Receber') and contains(normalize-space(.),'baixar')]"
    )))
    click_js(driver, el)

def selecionar_shopee_conta_contabil(driver, timeout=20):
    w = WebDriverWait(driver, timeout)

    sel = w.until(EC.element_to_be_clickable((By.ID, "idContaContabil")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
    time.sleep(0.2)

    WebDriverWait(driver, 10).until(
        lambda d: len(Select(d.find_element(By.ID, "idContaContabil")).options) > 1
    )

    try:
        ActionChains(driver).move_to_element(sel).click(sel).perform()
        time.sleep(0.2)
        opt = w.until(EC.element_to_be_clickable((
            By.XPATH, f"//select[@id='idContaContabil']/option[normalize-space()='{CONTA_SHOPEE_TEXTO}']"
        )))
        opt.click()
        atual = Select(driver.find_element(By.ID, "idContaContabil")).first_selected_option.text.strip().upper()
        if CONTA_SHOPEE_TEXTO in atual:
            return
    except Exception:
        pass

    try:
        Select(driver.find_element(By.ID, "idContaContabil")).select_by_value(CONTA_SHOPEE_VALUE)
        atual = Select(driver.find_element(By.ID, "idContaContabil")).first_selected_option.text.strip().upper()
        if CONTA_SHOPEE_TEXTO in atual:
            return
    except Exception:
        pass

    sel2 = driver.find_element(By.ID, "idContaContabil")
    driver.execute_script("""
        const select = arguments[0];
        const val = arguments[1];
        select.value = val;
        select.dispatchEvent(new Event('input', { bubbles: true }));
        select.dispatchEvent(new Event('change', { bubbles: true }));
    """, sel2, CONTA_SHOPEE_VALUE)

    atual = Select(driver.find_element(By.ID, "idContaContabil")).first_selected_option.text.strip().upper()
    if CONTA_SHOPEE_TEXTO not in atual:
        raise RuntimeError("Não consegui selecionar SHOPEE em idContaContabil.")

def selecionar_receita_shopee_categoria(driver, timeout=20):
    w = WebDriverWait(driver, timeout)
    sel_cat = w.until(EC.presence_of_element_located((By.ID, "idCategoria")))
    Select(sel_cat).select_by_visible_text(CATEGORIA_SHOPEE_TEXTO)

def preencher_data(driver, data_br: str | None, timeout=20):
    if not data_br:
        return
    w = WebDriverWait(driver, timeout)
    inp = w.until(EC.element_to_be_clickable((By.ID, "data")))
    inp.click()
    inp.send_keys(Keys.CONTROL, "a")
    inp.send_keys(Keys.DELETE)
    inp.send_keys(data_br)
    inp.send_keys(Keys.ENTER)

def preencher_taxas_e_frete(driver, taxa_br: str | None, frete_br: str | None, frete_num: float, timeout=25):
    w = WebDriverWait(driver, timeout)

    try:
        lbl = w.until(EC.presence_of_element_located((By.XPATH, "//label[normalize-space()='Taxas']")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", lbl)
        time.sleep(0.2)
    except Exception:
        pass

    if taxa_br is not None:
        inp_taxa = w.until(EC.element_to_be_clickable((By.ID, INPUT_TAXA_ID)))
        inp_taxa.click()
        inp_taxa.send_keys(Keys.CONTROL, "a")
        inp_taxa.send_keys(Keys.DELETE)
        inp_taxa.send_keys(taxa_br)
        inp_taxa.send_keys(Keys.ENTER)
        time.sleep(0.15)

    if frete_br is not None and frete_num > 0:
        inp_desc = w.until(EC.element_to_be_clickable((By.ID, INPUT_DESCONTO_ID)))
        inp_desc.click()
        inp_desc.send_keys(Keys.CONTROL, "a")
        inp_desc.send_keys(Keys.DELETE)
        inp_desc.send_keys(frete_br)
        inp_desc.send_keys(Keys.ENTER)
        time.sleep(0.15)

def preencher_valor_liquido(driver, valor_br: str | None, timeout=25):
    if valor_br is None:
        return
    w = WebDriverWait(driver, timeout)
    inp = w.until(EC.element_to_be_clickable((By.ID, INPUT_VALOR_ID)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    time.sleep(0.15)
    inp.click()
    inp.send_keys(Keys.CONTROL, "a")
    inp.send_keys(Keys.DELETE)
    inp.send_keys(valor_br)
    inp.send_keys(Keys.ENTER)

def aplicar_mais_opcoes_shopee(driver, data_br: str | None, timeout=35):
    w = WebDriverWait(driver, timeout)
    btn_mais = w.until(EC.element_to_be_clickable((By.ID, "linkUmaConta")))
    click_js(driver, btn_mais)

    selecionar_shopee_conta_contabil(driver, timeout=timeout)
    selecionar_receita_shopee_categoria(driver, timeout=timeout)
    preencher_data(driver, data_br=data_br, timeout=timeout)

def clicar_receber_contas_final(driver, timeout=35):
    w = WebDriverWait(driver, timeout)
    btn = w.until(EC.element_to_be_clickable((By.ID, BTN_SALVAR_BORDERO_ID)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.2)
    click_js(driver, btn)
    time.sleep(0.8)
    fechar_alerta_se_existir(driver)

# =====================================================
# START
# =====================================================

print("⚠️ Feche TODAS as janelas do Chrome antes de continuar.")
input("Quando tiver fechado, aperte ENTER... ")

clonar_perfil()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={CHROME_USER_DATA_CLONE}")
options.add_argument(f"--profile-directory={CHROME_PROFILE_CLONE}")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

prefs = {"profile.default_content_setting_values.notifications": 2}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, WAIT_TIMEOUT)

driver.get(URL)
time.sleep(2)

txt = fechar_alerta_se_existir(driver)
if txt:
    screenshot(driver, "alerta_inicial.png")
    print("\n⚠️ Alerta fechado:", txt)
    if sessao_expirada(txt):
        input("\n👉 Faça login novamente no Olist e volte para Contas a Receber. ENTER... ")
        garantir_na_tela_contas_receber(driver)

garantir_na_tela_contas_receber(driver)

# =====================================================
# EXCEL
# =====================================================

print("\nAbrindo Excel...")
df = pd.read_excel(ARQUIVO_ENTRADA)

df[COLUNA_CODIGO] = df[COLUNA_CODIGO].astype(str).str.strip()

if COLUNA_STATUS not in df.columns:
    df[COLUNA_STATUS] = ""
df[COLUNA_STATUS] = df[COLUNA_STATUS].astype("string")

if COLUNA_VALIDACAO in df.columns:
    df[COLUNA_VALIDACAO] = df[COLUNA_VALIDACAO].astype("string")

df["_DATA_BR"] = df[COLUNA_DATA].apply(formatar_data_br)

mapa_data = {}
mapa_taxas = {}
mapa_frete = {}
mapa_valor = {}
mapa_validacao = {}
mapa_frete_num = {}

for _, row in df.iterrows():
    pid = str(row.get(COLUNA_CODIGO, "")).strip()
    if not pid or pid.lower() == "nan":
        continue

    mapa_data[pid] = row.get("_DATA_BR")
    mapa_taxas[pid] = br_money(row.get(COLUNA_TOTAL_TAXAS, None))
    mapa_frete[pid] = br_money(row.get(COLUNA_FRETE_COBRADO, None))
    mapa_frete_num[pid] = valor_num(row.get(COLUNA_FRETE_COBRADO, None))
    mapa_valor[pid] = br_money(row.get(COLUNA_VALOR_LIQUIDO, None))

    v = str(row.get(COLUNA_VALIDACAO, "")).strip().lower()
    mapa_validacao[pid] = v

pedidos = df[COLUNA_CODIGO].dropna().astype(str).str.strip().tolist()
print(f"Total de pedidos: {len(pedidos)}")

# =====================================================
# PROCESSO
# =====================================================

def processar_pedido(codigo: str) -> str:
    alerta = fechar_alerta_se_existir(driver)
    if alerta and sessao_expirada(alerta):
        return "RELOGAR"

    codigo = str(codigo).strip()

    if mapa_validacao.get(codigo, "") != "ok":
        return "PULADO_VALIDACAO"

    try:
        atual = df.loc[df[COLUNA_CODIGO] == codigo, COLUNA_STATUS].iloc[0]
        if str(atual).strip().upper() == "SIM":
            return "JA_BAIXADO"
    except Exception:
        pass

    data_br = mapa_data.get(codigo)
    taxa_br = mapa_taxas.get(codigo)
    frete_br = mapa_frete.get(codigo)
    frete_num = mapa_frete_num.get(codigo, 0.0)
    valor_br = mapa_valor.get(codigo)

    busca = achar_input_busca(driver, timeout=25)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", busca)
    time.sleep(0.1)

    busca.click()
    busca.send_keys(Keys.CONTROL, "a")
    busca.send_keys(Keys.DELETE)
    time.sleep(0.05)
    busca.send_keys(codigo)

    lupa = achar_botao_lupa(driver, timeout=20)
    click_js(driver, lupa)

    resp = esperar_resultado_da_busca(driver, codigo, timeout=35)
    if resp != "OK":
        return resp

    clicar_navigate_da_linha(driver, codigo)
    time.sleep(1.2)

    clicar_receber_baixar(driver, timeout=35)
    time.sleep(0.8)

    aplicar_mais_opcoes_shopee(driver, data_br=data_br, timeout=35)

    preencher_taxas_e_frete(driver, taxa_br=taxa_br, frete_br=frete_br, frete_num=frete_num, timeout=35)
    preencher_valor_liquido(driver, valor_br=valor_br, timeout=35)

    clicar_receber_contas_final(driver, timeout=40)

    return "BAIXADO_OK"

def processar_com_tentativas(codigo: str) -> str:
    codigo = str(codigo).strip()

    for t in range(1, TENTATIVAS_POR_PEDIDO + 1):
        try:
            status = processar_pedido(codigo)

            if status in ("BAIXADO_OK", "NAO_ENCONTRADO", "PULADO_VALIDACAO", "JA_BAIXADO"):
                return status

            if status == "RELOGAR":
                screenshot(driver, f"relogar_{codigo}_t{t}.png")
                input("\n👉 Sessão expirou. Faça login novamente e volte para Contas a Receber. ENTER... ")
                garantir_na_tela_contas_receber(driver)
                continue

            if status == "TIMEOUT":
                screenshot(driver, f"timeout_{codigo}_t{t}.png")

                # ✅ antes de dar TIMEOUT, confirma se baixou mesmo
                try:
                    if confirmar_se_baixou(driver, codigo, timeout=20):
                        return "BAIXADO_OK"
                except Exception:
                    pass

                time.sleep(0.5)
                continue

            screenshot(driver, f"outro_{codigo}_{status}_t{t}.png")
            time.sleep(0.5)

        except UnexpectedAlertPresentException:
            txt = fechar_alerta_se_existir(driver)
            screenshot(driver, f"unexpected_alert_{codigo}_t{t}.png")
            if txt and sessao_expirada(txt):
                input("\n👉 Sessão expirou. Faça login novamente e volte para Contas a Receber. ENTER... ")
                garantir_na_tela_contas_receber(driver)
                continue

        except Exception:
            screenshot(driver, f"erro_{codigo}_t{t}.png")

            # ✅ antes de dizer que deu ERRO, tenta confirmar se baixou mesmo
            try:
                if confirmar_se_baixou(driver, codigo, timeout=20):
                    return "BAIXADO_OK"
            except Exception:
                pass

            time.sleep(0.5)

    return "ERRO"

# =====================================================
# TESTE + LOOP
# =====================================================

if pedidos:
    teste = str(pedidos[0]).strip()
    print(f"\n🔎 Teste com: {teste}")
    print("   VALIDAÇÃO:", mapa_validacao.get(teste))
    st = processar_com_tentativas(teste)
    print("Resultado teste:", st)
    screenshot(driver, f"teste_{teste}_{st}.png")
    input("\nSe deu tudo certo, ENTER para iniciar tudo... ")

print("\nIniciando...\n")

for i, codigo in enumerate(pedidos, start=1):
    codigo = str(codigo).strip()
    print(f"[{i}/{len(pedidos)}] {codigo}")

    status = processar_com_tentativas(codigo)

    if status == "BAIXADO_OK":
        df.loc[df[COLUNA_CODIGO] == codigo, COLUNA_STATUS] = "SIM"
    else:
        atual = ""
        try:
            atual = df.loc[df[COLUNA_CODIGO] == codigo, COLUNA_STATUS].iloc[0]
        except Exception:
            pass
        if str(atual).strip().upper() != "SIM":
            df.loc[df[COLUNA_CODIGO] == codigo, COLUNA_STATUS] = status

    if i % SALVAR_A_CADA == 0:
        df.to_excel(ARQUIVO_SAIDA, index=False)
        print(f"💾 Checkpoint salvo: {ARQUIVO_SAIDA}\n")

df.to_excel(ARQUIVO_SAIDA, index=False)
print("\n✅ Finalizado!")
print("Arquivo salvo:", ARQUIVO_SAIDA)

driver.quit()
