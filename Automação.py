import os
import time
import shutil
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import threading

# Configurações iniciais
edge_driver_path = r"C:\\Users\\matheus.assis\\Downloads\\MsgDriver\\msedgedriver.exe"
url = "https://copilotstudio.microsoft.com/environments/Default-d369179f-4330-4652-b8f3-cbf871056d2f/bots/83503342-3225-f011-8c4d-00224832457c/adaptive/0e275e01-aa22-4954-a275-2b97f27bcfae/triggers/main/actions/cLz89I/edit"
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
bases_path = os.path.join(downloads_path, "Bases")

os.makedirs(bases_path, exist_ok=True)

arquivos_processados = {}

def encontrar_arquivo_novo(com_texto):
    arquivos = [f for f in os.listdir(downloads_path) if com_texto in f and f.endswith('.xlsx')]
    arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_path, x)), reverse=True)
    for arquivo in arquivos:
        caminho = os.path.join(downloads_path, arquivo)
        mtime = os.path.getmtime(caminho)
        try:
            with open(caminho, 'rb'):
                pass
        except Exception:
            continue
        if caminho not in arquivos_processados or arquivos_processados[caminho] < mtime:
            arquivos_processados[caminho] = mtime
            return caminho
    return None

def tratar_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        if len(df.columns) >= 4:
            df.drop(df.columns[3], axis=1, inplace=True)
        df.to_excel(caminho_arquivo, index=False)
        print(f"✅ Arquivo tratado: {os.path.basename(caminho_arquivo)}")
        return True
    except Exception as e:
        print(f"❌ Erro ao tratar o arquivo: {e}")
        return False

def mover_arquivo_para_bases(caminho_arquivo):
    try:
        nome_arquivo = os.path.basename(caminho_arquivo)
        destino = os.path.join(bases_path, nome_arquivo)
        shutil.move(caminho_arquivo, destino)
        print(f"✅ Arquivo movido: {nome_arquivo}")
        return destino
    except Exception as e:
        print(f"❌ Erro ao mover o arquivo: {e}")
        return None

def realizar_upload(arquivo):
    options = Options()
    options.use_chromium = True
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    service = EdgeService(executable_path=edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)

    try:
        print("🌐 Acessando Copilot Studio...")
        driver.get(url)
        wait = WebDriverWait(driver, 30)

        for _ in range(3):
            try:
                ignorar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Ignorar")]')))
                ignorar_btn.click()
                print("➡️ Botão 'Ignorar' clicado.")
                break
            except StaleElementReferenceException:
                print("🔁 Tentando clicar no botão 'Ignorar' novamente...")
                time.sleep(1)

        adicionar_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-telemetry-id="KnowledgeSourcesField-AddKnowledge"]')))
        adicionar_btn.click()

        input_file = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]')))
        input_file.send_keys(arquivo)
        print(f"📄 Upload enviado: {arquivo}")

        adicionar_final_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-telemetry-id="FilesCreate-complete"]')))
        adicionar_final_btn.click()

        nome_arquivo = os.path.basename(arquivo)
        wait.until(EC.visibility_of_element_located((By.XPATH, f'//div[contains(text(), "{nome_arquivo}")]')))
        print(f"✅ Upload confirmado: {nome_arquivo}")

        time.sleep(2)

    except Exception as e:
        print(f"❌ Erro no upload: {e}")
        traceback.print_exc()
    finally:
        driver.quit()

def abrir_site_para_interacao_manual():
    options = Options()
    options.use_chromium = True
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    service = EdgeService(executable_path=edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)

    def fechar_navegador_automaticamente():
        print("⏰ O tempo de 3 minutos expirou. O navegador será fechado agora.")
        driver.quit()

    try:
        print("🌐 Acessando Copilot Studio para interação manual...")
        driver.get(url)
        wait = WebDriverWait(driver, 30)

        try:
            ignorar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Ignorar")]')))
            ignorar_btn.click()
            print("➡️ Botão 'Ignorar' clicado.")
        except:
            print("(Ignorar não apareceu)")

        print("🔄 Executando cliques de publicação...")

        publicar_btn1 = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Publicar"]')))
        publicar_btn1.click()
        print("✅ Primeiro botão 'Publicar' clicado.")

        publicar_btn2 = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="BotPublishingStates-confirm-publish"]')))
        publicar_btn2.click()
        print("✅ Segundo botão 'Publicar' clicado.")

        print("🔄 Publicação finalizada.")

        timer = threading.Timer(180, fechar_navegador_automaticamente)
        timer.start()

        input("Pressione Enter para encerrar o navegador antes dos 3 minutos...")
        timer.cancel()
        print("🔒 Navegador fechado manualmente antes do tempo.")

    except Exception as e:
        print(f"❌ Erro ao abrir o site para interação manual: {e}")
        traceback.print_exc()
    finally:
        driver.quit()

def monitorar_pasta():
    print("📂 Monitorando Downloads... (Ctrl+C para parar)")
    while True:
        novo_arquivo = encontrar_arquivo_novo("lista_de_convites")
        if novo_arquivo:
            print(f"📄 Novo arquivo: {os.path.basename(novo_arquivo)}")
            if tratar_excel(novo_arquivo):
                realizar_upload(novo_arquivo)
                novo_caminho = mover_arquivo_para_bases(novo_arquivo)
                if novo_caminho:
                    arquivos_processados[novo_caminho] = os.path.getmtime(novo_caminho)
                abrir_site_para_interacao_manual()
        time.sleep(120)

if __name__ == "__main__":
    monitorar_pasta()
