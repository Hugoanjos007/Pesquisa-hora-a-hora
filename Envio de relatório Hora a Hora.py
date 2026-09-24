import os
import time
import shutil
from selenium.webdriver.common import options
from selenium.common.exceptions import StaleElementReferenceException
import schedule
from datetime import datetime
from PIL import Image
import win32clipboard
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv

# ================= CONFIGURAÇÕES =================
load_dotenv()
USUARIO = os.getenv("USUARIO")
SENHA = os.getenv("SENHA")
TIMEOUT = int(os.getenv("TIMEOUT", "20"))
URL_LOGIN = os.getenv("URL_LOGIN")
URL_RELATORIO_1 = os.getenv("URL_RELATORIO_1")
URL_RELATORIO_2 = os.getenv("URL_RELATORIO_2")
URL_RELATORIO_3 = os.getenv("URL_RELATORIO_3")
PASTA_RELATORIO_1 = os.getenv("PASTA_RELATORIO_1")
PASTA_RELATORIO_2 = os.getenv("PASTA_RELATORIO_2")
PASTA_RELATORIO_3 = os.getenv("PASTA_RELATORIO_3")
PASTA_TEMP_DOWNLOAD = os.getenv("PASTA_TEMP_DOWNLOAD")
PASTA_PRINTS_LOCAL = os.getenv("PASTA_PRINTS_LOCAL")
PASTA_REDE_INDICADORES = os.getenv("PASTA_REDE_INDICADORES")
CAMINHO_EXCEL = os.getenv("CAMINHO_EXCEL")
NOME_MACRO_PRINCIPAL = os.getenv("NOME_MACRO_PRINCIPAL")
NOME_MACRO_DASHBOARD = os.getenv("NOME_MACRO_DASHBOARD")
TEAMS_URL = os.getenv("TEAMS_URL")
CHROME_PROFILE = os.getenv("CHROME_PROFILE")
# ================= UTILITÁRIOS =================
def log(msg):
    print(f"[{datetime.now().strftime('%d/%m %H:%M:%S')}] {msg}")
def limpar_pasta(pasta):
    os.makedirs(pasta, exist_ok=True)
    for f in os.listdir(pasta):
        caminho = os.path.join(pasta, f)
        if os.path.isfile(caminho):
            try:
                os.remove(caminho)
            except PermissionError as erro:
                log(f"Não foi possível apagar {caminho}: {erro}")
                pass
def aguardar_download(pasta, extensoes=(".csv", ".xlsx", ".xls"), timeout=180):
    fim = time.time() + timeout
    while time.time() < fim:
        arquivos = [
            os.path.join(pasta, a)
            for a in os.listdir(pasta)
            if a.lower().endswith(extensoes)
        ]
        parciais = [a for a in os.listdir(pasta) if a.endswith(".crdownload")]
        if arquivos and not parciais:
            return max(arquivos, key=os.path.getmtime)
        time.sleep(1)
    raise TimeoutError("Download não concluído no tempo esperado.")
def mover_arquivo(origem, destino):
    os.makedirs(destino, exist_ok=True)
    shutil.move(origem, os.path.join(destino, os.path.basename(origem)))
def salvar_print(ws, intervalo, nome_arquivo):
    pasta_destino = os.path.dirname(nome_arquivo)
    os.makedirs(pasta_destino, exist_ok=True)
    os.makedirs(PASTA_TEMP_DOWNLOAD, exist_ok=True)
    nome_temp = os.path.join(
        PASTA_TEMP_DOWNLOAD,
        f"print_temp_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.png"
    )
    ws.Activate()
    rng = ws.Range(intervalo)
    chart = None
    try:
        rng.CopyPicture(Appearance=1, Format=2)
        time.sleep(3)
        chart = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        chart.Chart.Paste()
        # Salva localmente; o Excel costuma falhar ao exportar direto em unidade de rede
        chart.Chart.Export(nome_temp, "PNG")
        # Python copia para o destino final, local ou rede
        shutil.copy2(nome_temp, nome_arquivo)
        log(f"📸 Print salvo: {os.path.basename(nome_arquivo)}")
    finally:
        if chart is not None:
            chart.Delete()
        if os.path.exists(nome_temp):
            os.remove(nome_temp)
# ================= CLIPBOARD (TEAMS) =================
def copiar_imagem_clipboard(caminho):
    img = Image.open(caminho).convert("RGB")
    temp_bmp = os.path.join(os.environ["TEMP"], "clipboard_image.bmp")
    img.save(temp_bmp, "BMP")
    with open(temp_bmp, "rb") as f:
        data = f.read()[14:]
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()
def ultimas_duas_imagens(pasta):
    arquivos = [
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.lower().endswith(".png")
    ]
    arquivos.sort(key=os.path.getmtime, reverse=True)
    return arquivos[:2]
def remover_linhas_ecommerce(wb):
    ws = wb.Worksheets("pesquisa hr a hr relatorio cham")

    if ws.ListObjects.Count == 0:
        raise RuntimeError(
            "Não foi encontrada uma Tabela do Excel na aba "
            "'pesquisa hr a hr relatorio cham'."
        )
    tabela = ws.ListObjects.Item(1)
    # Localiza a coluna pelo nome, sem depender da posição BH
    coluna_campanha = None
    for i in range(1, tabela.ListColumns.Count + 1):
        nome_coluna = str(tabela.ListColumns.Item(i).Name).strip()

        if nome_coluna.upper() == "NOME CAMPANHA":
            coluna_campanha = i
            break
    if coluna_campanha is None:
        raise RuntimeError(
            "A coluna 'Nome Campanha' não foi encontrada na tabela."
        )
    try:
        if ws.FilterMode:
            log("🔎 Removendo filtro ativo da tabela")
            ws.ShowAllData()
    except Exception as erro:
        log(f"Ocorreu um erro: {erro}")
        pass
    removidas = 0
    for linha in range(tabela.ListRows.Count, 0, -1):
        valor = tabela.DataBodyRange.Cells(linha, coluna_campanha).Value
        campanha = str(valor or "").strip().upper()
        if campanha == "E_COMMERCE":
            tabela.ListRows.Item(linha).Delete()
            removidas += 1
    log(f"🧹 Linhas E_COMMERCE removidas: {removidas}")
# ================= ENVIO TEAMS =================
def enviar_teams():
    log("➡️ Abrindo Teams")
    driver = None
    try:
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        os.makedirs(CHROME_PROFILE, exist_ok=True)
        options.add_argument(f"--user-data-dir={CHROME_PROFILE}")
        options.add_argument("--profile-directory=Default")
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(
            driver,
            120,
            ignored_exceptions=(StaleElementReferenceException,)
        )
        driver.get(TEAMS_URL)
        def localizar_campo_mensagem(navegador):
            campos = navegador.find_elements(
                By.XPATH,
                "//div[@contenteditable='true' and @role='textbox']"
            )
            for campo in reversed(campos):
                try:
                    if campo.is_displayed() and campo.is_enabled():
                        return campo
                except StaleElementReferenceException:
                    continue
            return False
        # Tenta novamente caso o Teams atualize a tela durante o carregamento
        for tentativa in range(3):
            try:
                campo_msg = wait.until(localizar_campo_mensagem)
                driver.execute_script("""
                    arguments[0].scrollIntoView({block: 'center'});
                    arguments[0].focus();
                """, campo_msg)
                break
            except StaleElementReferenceException:
                if tentativa == 2:
                    raise
                time.sleep(2)
        time.sleep(1)
        prints = ultimas_duas_imagens(PASTA_PRINTS_LOCAL)
        if len(prints) < 2:
            log("❌ Não foram encontradas 2 imagens para envio")
            return
        hora = datetime.now().strftime("%Hh%M")
        log("📤 Enviando imagens no chat")
        for img in reversed(prints):
            copiar_imagem_clipboard(img)
            time.sleep(2)
            campo_msg.send_keys(Keys.CONTROL, "v")
            time.sleep(3)
        campo_msg.send_keys(f" {hora}")
        time.sleep(1)
        campo_msg.send_keys(Keys.ENTER)
        log("✅ Mensagem enviada com sucesso")
        time.sleep(5)
    finally:
        if driver is not None:
            driver.quit()
# ================= AUTOMAÇÃO PRINCIPAL =================
def executar_automacao():

    hora_atual = datetime.now().hour
    if not (8 <= hora_atual <= 21):
        log("⏳ Fora do horário permitido")
        return
    log(f"🚀 Iniciando ciclo das {hora_atual}:00")
    limpar_pasta(PASTA_TEMP_DOWNLOAD)
    limpar_pasta(PASTA_RELATORIO_1)
    limpar_pasta(PASTA_RELATORIO_2)
    limpar_pasta(PASTA_RELATORIO_3)
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": PASTA_TEMP_DOWNLOAD,
        "download.prompt_for_download": False
    })
    # Código atualizado para o Python 3.13 / Selenium 4+
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, TIMEOUT)

    try:
        log("🔐 Login no CRM")
        driver.get(URL_LOGIN)

        wait.until(EC.presence_of_element_located((By.ID, "l_login"))).send_keys(USUARIO)
        driver.find_element(By.ID, "l_senha").send_keys(SENHA)
        driver.execute_script("enviarDados();")

        wait.until(EC.url_contains("index.php"))
        log("✅ Login realizado")

        log("📊 Relatório 1")
        driver.get(URL_RELATORIO_1)
        wait.until(EC.element_to_be_clickable((By.ID, "btn_pesquisar"))).click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//img[contains(@src,'csv') or contains(@src,'excel')]/ancestor::span"))
        ).click()
        mover_arquivo(aguardar_download(PASTA_TEMP_DOWNLOAD), PASTA_RELATORIO_1)

        log("📊 Relatório 2")
        driver.get(URL_RELATORIO_2)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//img[contains(@src,'csv') or contains(@src,'excel')]/ancestor::span"))
        ).click()
        mover_arquivo(aguardar_download(PASTA_TEMP_DOWNLOAD), PASTA_RELATORIO_2)

        log("📊 Relatório 3")
        driver.get(URL_RELATORIO_3)
        wait.until(EC.presence_of_element_located((By.ID, "agrupador")))

        driver.execute_script("""
            var select = document.getElementById('agrupador');
            select.value = 'usuario';
            select.dispatchEvent(new Event('change'));
        """)
        time.sleep(1)

        wait.until(EC.element_to_be_clickable((By.ID, "btn_pesquisar"))).click()
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//img[contains(@src,'csv') or contains(@src,'excel')]/ancestor::span"))
        ).click()
        mover_arquivo(aguardar_download(PASTA_TEMP_DOWNLOAD), PASTA_RELATORIO_3)

    finally:
        driver.quit()

    # ================= EXCEL =================
    if not os.path.isfile(CAMINHO_EXCEL):
        raise FileNotFoundError(
            f"Planilha não encontrada: {CAMINHO_EXCEL}"
        )

    log("📗 Abrindo planilha do Excel")

    # Cria uma instância exclusiva, evitando reutilizar um Excel travado
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(
        Filename=os.path.abspath(CAMINHO_EXCEL),
        UpdateLinks=0,
        ReadOnly=False
    )

    # Em alguns casos, o Excel abre o arquivo, mas não devolve sua referência
    if wb is None:
        log("⚠️ Excel não devolveu a referência. Procurando a planilha aberta...")

        time.sleep(3)
        caminho_esperado = os.path.normcase(os.path.abspath(CAMINHO_EXCEL))

        for indice in range(1, excel.Workbooks.Count + 1):
            livro = excel.Workbooks.Item(indice)

            if os.path.normcase(os.path.abspath(livro.FullName)) == caminho_esperado:
                wb = livro
                break

    if wb is None:
        excel.Quit()
        raise RuntimeError(
            "O Excel foi iniciado, mas a planilha não pôde ser acessada. "
            "Verifique se ela já está aberta, bloqueada ou exibindo algum aviso."
        )

    log(f"✅ Planilha aberta: {wb.Name}")
    for i in range(3):
        log(f"🔄 RefreshAll ({i+1}/3)")
        wb.RefreshAll()
        time.sleep(10)
    remover_linhas_ecommerce(wb)

    log("▶ Macro principal")
    excel.Run(f"'{wb.Name}'!{NOME_MACRO_PRINCIPAL}")
    time.sleep(5)

    wb.RefreshAll()
    time.sleep(10)
    remover_linhas_ecommerce(wb)
    log("▶ Macro Dashboard")
    excel.Run(f"'{wb.Name}'!{NOME_MACRO_DASHBOARD}")
    time.sleep(5)

    hora_str = datetime.now().strftime("%Hh%Mm")

    salvar_print(
        wb.Worksheets("Dashboard"),
        "A1:Q36",
        os.path.join(PASTA_PRINTS_LOCAL, f"{hora_str} Dashboard.png")
    )

    ws = wb.Worksheets("Analitico")
    valores = ws.Range("B1:B200").Value

    u_lin = 4
    for i, r in enumerate(valores):
        if r[0]:
            u_lin = i + 1

    intervalo = f"A1:U{u_lin}"

    salvar_print(
        ws,
        intervalo,
        os.path.join(PASTA_PRINTS_LOCAL, f"{hora_str} Analitico.png")
    )

    salvar_print(
        ws,
        intervalo,
        os.path.join(PASTA_REDE_INDICADORES, f"{hora_str} Analitico.png")
    )

    # >>>>>>> NOVO COMPORTAMENTO (SEU PEDIDO) <<<<<<<<
    log("📨 Disparando envio ao Teams (após salvar prints)")
    enviar_teams()

    time.sleep(3)

    log("📕 Fechando somente o relatório (sem salvar)")
    wb.Close(SaveChanges=False)

    log("🏁 Ciclo finalizado com sucesso")

# ================= PERGUNTA INICIAL (SEU PEDIDO) =================

resposta = input("Rodar agora? (S = rodar agora / N = aguardar virar a hora): ").strip().upper()

if resposta == "S":
    log("▶ Rodando imediatamente por comando do usuário")
    executar_automacao()
else:
    log("⏳ Aguardando virar a hora cheia...")

# ================= AGENDAMENTO =================
schedule.every(30).minutes.do(executar_automacao)

log("🤖 Robô ativo (09h às 21h)")

while True:
    schedule.run_pending()
    time.sleep(30)