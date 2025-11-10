import os
import shutil
import subprocess
import time
from time import sleep

import requests
import wmi
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service

# ==========================================
URL = "https://mugiwarasoficial.com/manga/chainsaw-man/capitulo-"
BASE_PATH = r"C:\Users\Lucas Alves\Downloads"
FOLDER_NAME = "Chainsaw Man"
CAPS = (207, 207)
# ==========================================


def download_images(url: str, base_path: str, folder_name: str, caps: tuple[int, int]):
    ini = caps[0]
    end = caps[1]
    for number in range(ini, end + 1):
        print(f"Baixando Capítulo {number:02d}")
        folder_complete = f"{folder_name} {number:02d}"
        pasta_destino = os.path.join(base_path, folder_complete)
        os.makedirs(pasta_destino, exist_ok=True)

        url_complete = f"{url}{number:02d}-paa-paa-paa-paa"
        options = webdriver.FirefoxOptions()
        options.add_argument("--headless")
        service = Service()
        driver = webdriver.Firefox(service=service, options=options)

        driver.get(url_complete)

        print("Rolando a página para carregar todas as imagens...")
        SCROLL_PAUSE_TIME = 1.5
        last_height = driver.execute_script("return document.body.scrollHeight")

        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(SCROLL_PAUSE_TIME)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        pages = driver.find_elements(
            By.CSS_SELECTOR,
            ".wp-manga-chapter-img.img-responsive.effect-fade.lazyloaded",
        )
        print(f"Total de Páginas: {len(pages)}")
        for index, page in enumerate(pages):
            sleep(0.5)
            image_src = page.get_attribute("src")
            response = requests.get(image_src, stream=True)
            response.raise_for_status()
            filename = f"{(index + 1):02d}.png"
            caminho_completo_arquivo = os.path.join(pasta_destino, filename)
            with open(caminho_completo_arquivo, "wb") as out_file:
                for chunk in response.iter_content(chunk_size=8192):
                    out_file.write(chunk)
            print(f"Página {filename} baixada com sucesso!")
        print(f"Download do Capítulo {number:02d} concluído")
        driver.quit()
        generate_mobi(pasta_destino)
        shutil.rmtree(pasta_destino)

        kindle = find_kindle_letter("Kindle")
        caminho_origem = os.path.join(base_path, f"{folder_complete}.mobi")
        caminho_destino_pasta = os.path.join(kindle, "documents")
        caminho_destino_final = os.path.join(
            caminho_destino_pasta, f"{folder_complete}.mobi"
        )

        if not os.path.isdir(caminho_destino_pasta):
            print("Kindle nao conectado ao computador")
            continue

        if not os.path.exists(caminho_origem):
            print("Erro: falhou ao converter para arquivo mobi")
            continue

        try:
            shutil.move(caminho_origem, caminho_destino_final)
            print(
                f"Sucesso! O arquivo '{folder_complete}.mobi' foi movido para: {caminho_destino_final}"
            )

        except Exception as e:
            print(f"Erro ao mover o arquivo: {e}")


def generate_mobi(folder_path: str):
    subprocess.call(
        [
            "kcc",
            "K11",
            folder_path,
            "-m",
            "-u",
            "-s",
            "-c",
            "2",
            "-r",
            "1",
            "-f",
            "MOBI",
        ]
    )


def find_kindle_letter(kindle_name):
    try:
        c = wmi.WMI()

        for drive in c.Win32_LogicalDisk():
            if drive.VolumeName and drive.VolumeName.lower() == kindle_name.lower():
                return drive.DeviceID

    except Exception as e:
        print(f"Erro ao tentar acessar o WMI: {e}")
        return None

    return None


if __name__ == "__main__":
    download_images(URL, BASE_PATH, FOLDER_NAME, CAPS)
