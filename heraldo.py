import os
import wget
import shutil
import locale
import zipfile
import requests
import pyautogui
import numpy as np
import pandas as pd
import tkinter as tk
from fpdf import FPDF
from datetime import datetime
from tkinter import messagebox
from subprocess import run, PIPE
from payconpy.fpython.fpython import *
from datetime import datetime, date, timedelta


def atualiza_chromedriver(self):
    bin_folder_path = "ArquivosRobo/bin"
    os.makedirs(bin_folder_path, exist_ok=True)

    def get_actual_chromedriver_path(folder):
        """Find the ChromeDriver binary in the specified folder."""
        pattern = os.path.join(
            folder, "chromedriver.exe"
        )  # Explicitly look for .exe file
        if os.path.exists(pattern):
            return pattern
        return None

    def get_latest_chromedriver_version():
        """Fetch the latest stable ChromeDriver version."""
        url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json"
        try:
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()
            return data["channels"]["Stable"]["version"]
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to fetch ChromeDriver version: {e}")
        except:
            raise RuntimeError(
                "Failed to parse JSON data or find the expected structure."
            )

    def get_local_chromedriver_version(driver_path):
        """Get the version of the local ChromeDriver."""
        if driver_path and os.path.exists(driver_path):
            result = run(
                [driver_path, "--version"], stdout=PIPE, stderr=PIPE, text=True
            )
            if result.returncode == 0:
                return result.stdout.split(" ")[1].strip()
        return None

    chrome_driver_path = get_actual_chromedriver_path(bin_folder_path)
    local_version = get_local_chromedriver_version(chrome_driver_path)
    latest_version = get_latest_chromedriver_version()

    if not local_version or local_version.split(".")[0] != latest_version.split(".")[0]:
        print("Atualizando o chromedriver")

        download_url = f"https://storage.googleapis.com/chrome-for-testing-public/{latest_version}/win64/chromedriver-win64.zip"

        # Download the zip file using the URL built above
        latest_driver_zip = wget.download(download_url, "chromedriver.zip")

        # Ensure the bin directory exists
        destination_folder = os.path.join("bin")
        os.makedirs(destination_folder, exist_ok=True)

        # Move the downloaded ZIP file into the bin directory
        destination_zip = os.path.join(
            destination_folder, os.path.basename(latest_driver_zip)
        )
        shutil.move(latest_driver_zip, destination_zip)

        # Extract the ZIP file into the bin directory
        with zipfile.ZipFile(destination_zip, "r") as zip_ref:
            zip_ref.extractall(destination_folder)

        # Path of the extracted folder
        extracted_folder = os.path.join(destination_folder, "chromedriver-win64")

        # Move "chromedriver.exe" to the "bin" folder
        chromedriver_path = os.path.join(extracted_folder, "chromedriver.exe")

        if os.path.exists(chromedriver_path):
            if chrome_driver_path:  # Only remove if it exists
                os.remove(chrome_driver_path)  # Remove the existing file
            shutil.move(chromedriver_path, destination_folder)

        # Delete the extracted folder
        shutil.rmtree(extracted_folder)

        # Delete the ZIP file after extraction
        os.remove(destination_zip)
        print(
            f"\nChromedriver atualizado com sucesso para a versão {latest_version}.\n"
        )


def wait_for_downloads(directory, timeout=600):
    """
    Waits until there are no files with .crdownload or .part extensions in the directory,
    indicating that all downloads have completed.

    Parameters:
    - directory: The directory to monitor for downloads.
    - timeout: Maximum time to wait for downloads to complete, in seconds.

    Returns:
    - True if all downloads are completed within the timeout period, False otherwise.
    """
    start_time = time.time()

    while time.time() - start_time < timeout:
        # List all files in the directory
        files = os.listdir(directory)

        # Check for any temporary download files
        if not any(file.endswith((".crdownload", ".part")) for file in files):
            return True

        # Wait for a short period before checking again
        time.sleep(1)

    return False


def apaga_e_printa(frase, color="white"):
    os.system("cls" if os.name == "nt" else "clear")
    faz_log(frase, color=color)  # Assuming `faz_log` accepts `color` as a parameter


def progress_bar(current, total, bar_length=40):
    """
    Gera uma barra de progresso baseada em texto.

    Parâmetros:
        current (int): O progresso atual.
        total (int): O valor que representa 100% do progresso.
        bar_length (int): O comprimento da barra de progresso (o padrão é 40).

    Retorna:
        str: Uma string representando a barra de progresso.
    """

    if total == 0:
        return "Progress: [ERROR - Total cannot be zero]"

    progress = current / total
    if progress > 1:
        progress = 1  # Cap the progress at 100%

    filled_length = int(bar_length * progress)
    bar = "█" * filled_length + "-" * (bar_length - filled_length)
    percentage = progress * 100
    return f"\nPROGRESSO: \n[{bar}] Processo {current} de {total} - {percentage:.1f}%"


def show_success_popup(titulo, mensagem):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    # messagebox.showinfo()
    messagebox.showinfo(titulo, mensagem)
    root.destroy()  # Close the hidden window
def find_elements_with_retry(driver, wait_time, locator, retries=3):
    """Find elements and retry on StaleElementReferenceException."""
    for attempt in range(retries):
        try:
            elements = WebDriverWait(driver, wait_time).until(
                EC.presence_of_all_elements_located(locator)
            )
            return elements
        except StaleElementReferenceException:
            if attempt == retries - 1:
                raise


def find_element_text_with_retry(driver, wait_time, locator, retries=3):
    """Find element text and retry on StaleElementReferenceException."""
    for attempt in range(retries):
        try:
            element = WebDriverWait(driver, wait_time).until(
                EC.presence_of_element_located(locator)
            )
            return element.text
        except StaleElementReferenceException:
            if attempt == retries - 1:
                raise


def apaga_e_printa(frase, color="white"):
    os.system("cls" if os.name == "nt" else "clear")
    faz_log(frase, color=color)  # Assuming `faz_log` accepts `color` as a parameter


def progress_bar(current, total, bar_length=40):
    """
    Generates a text-based progress bar.

    Args:
        current (int): The current progress.
        total (int): The value representing 100% progress.
        bar_length (int): The length of the progress bar (default is 20).

    Returns:
        str: A string representing the progress bar.
    """
    if total == 0:
        return "Progress: [ERROR - Total cannot be zero]"

    progress = current / total
    if progress > 1:
        progress = 1  # Cap the progress at 100%

    filled_length = int(bar_length * progress)
    bar = "█" * filled_length + "-" * (bar_length - filled_length)
    percentage = progress * 100
    return f"PROGRESSO: [{bar}] {percentage:.1f}%"


# Function to download a file
def download_file(file_id, access_token, download_path):
    headers = {"Authorization": f"Bearer {access_token}"}
    download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    response = requests.get(download_url, headers=headers, stream=True)

    if response.status_code == 200:
        with open(download_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        print(f"Downloaded {download_path}")
    else:
        print(f"Failed to download file: {response.json()}")


# List files in the folder and download each file
def list_and_download_files(
    access_token, folder_id, download_dir, scope, client_id, client_secret, tenant_id
):
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children",
        headers=headers,
    )

    if response.status_code == 200:
        items = response.json().get("value", [])
        for item in items:
            if "file" in item:  # Check if the item is a file
                file_id = item["id"]
                file_name = item["name"]
                file_path = os.path.join(download_dir, file_name)
                download_file(file_id, access_token, file_path)
    elif response.status_code == 401:  # Token expired, retry with a new token
        app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        result = app.acquire_token_for_client(scopes=scope)
        if "access_token" in result:
            list_and_download_files(
                result["access_token"],
                folder_id,
                download_dir,
                scope,
                client_id,
                client_secret,
                tenant_id,
            )
        else:
            print("Failed to obtain access token.")
            print(result.get("error_description", "No error description available."))
            print(result.get("claims", "No claims available."))
    else:
        print(f"Failed to list files: {response.json()}")


# Function to get the item ID from the shared link


def get_item_id_from_shared_link(
    shared_link, access_token, scope, client_id, client_secret, tenant_id
):
    encoded_link = base64.urlsafe_b64encode(shared_link.encode()).decode().strip("=")
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{encoded_link}/driveItem",
        headers=headers,
    )
    if response.status_code == 200:
        return response.json()["id"]
    elif response.status_code == 401:  # Token expired, retry with a new token
        app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        result = app.acquire_token_for_client(scopes=scope)
        if "access_token" in result:
            return get_item_id_from_shared_link(
                shared_link,
                result["access_token"],
                scope,
                client_id,
                client_secret,
                tenant_id,
            )
        else:
            print("Failed to obtain access token.")
            print(result.get("error_description", "No error description available."))
            print(result.get("claims", "No claims available."))
            return None
    else:
        print(f"Failed to get item ID: {response.json()}")
        return None


def baixar_arquivos_onedrive(
    client_id,
    client_secret,
    tenant_id,
    caminho_pasta,
    folder_link,
    comentarios=False
    # baixar_tudo=False,
    # pasta_registros="base",
    # nome_coluna="caso",
    # nome_pasta="historico",
):
    '''
    A função original criada para o SPC possui uma lógica que verifica se o arquivo já foi baixado
    Esta parte foi removida para não ser preciso passar diversas variáveis desnecessárias
    A parte removida se encontrava no else após o if baixar_tudo
    '''
    if comentarios:
        faz_log("Iniciando o download do banco de dados dos processos encontrados", color="yellow")

    local_download_path = caminho_pasta
    # dataset_path = f"{pasta_registros}/{nome_pasta}.xlsx"  # Path to your dataset

    # Carregando o nome do arquivos já baixados
    # if comentarios:
        # faz_log("Carregando o nome do arquivos já baixados")
    # df = pd.read_excel(dataset_path)
    # downloaded_files = set(
    #     df[nome_coluna]
    # )  # Assuming 'nome' column contains the file names

    # Gerando token de acesso
    if comentarios:
        faz_log("Gerando token de acesso")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    token_response = requests.post(token_url, data=token_data)
    token_response.raise_for_status()
    access_token = token_response.json().get("access_token")

    # Codificando o link da pasta no formato base64
    if comentarios:
        faz_log("Codificando o link da pasta no formato base64")
    encoded_link = base64.b64encode(folder_link.encode()).decode()
    encoded_link = encoded_link.replace("/", "_").replace("+", "-").replace("=", "")

    # Acessando a pasta do OneDrive
    if comentarios:
        faz_log("Acessando a pasta do OneDrive")
    headers = {"Authorization": f"Bearer {access_token}"}
    shared_link_url = (
        f"https://graph.microsoft.com/v1.0/shares/u!{encoded_link}/driveItem"
    )
    shared_link_response = requests.get(shared_link_url, headers=headers)
    shared_link_response.raise_for_status()
    drive_item = shared_link_response.json()

    # Listando os arquivos na pasta do OneDrive
    if comentarios:
        faz_log("Codificando o link da pasta no formato base64")
    folder_id = drive_item["id"]
    files_url = f'https://graph.microsoft.com/v1.0/drives/{drive_item["parentReference"]["driveId"]}/items/{folder_id}/children'
    files_response = requests.get(files_url, headers=headers)
    files_response.raise_for_status()
    files = files_response.json().get("value", [])

    # Baixando arquivos pendentes
    if comentarios:
        faz_log("Baixando arquivos pendentes")
    if not os.path.exists(local_download_path):
        os.makedirs(local_download_path)

    with requests.Session() as session:
        session.headers.update({"Authorization": f"Bearer {access_token}"})

        for file in files:
            file_name = file["name"]
            # if baixar_tudo:
            if comentarios:
                faz_log(f"Baixando arquivo entitulado {file_name}")

            for attempt in range(3):  # Retry up to 3 times
                try:
                    file_id = file["id"]
                    download_url = f'https://graph.microsoft.com/v1.0/drives/{drive_item["parentReference"]["driveId"]}/items/{file_id}/content'
                    file_response = session.get(
                        download_url, timeout=30
                    )  # Increased timeout
                    file_response.raise_for_status()

                    with open(
                        os.path.join(local_download_path, file_name), "wb"
                    ) as f:
                        f.write(file_response.content)

                    if comentarios:
                        faz_log(f"Arquivo baixado com sucesso!", color="green")
                    break  # Exit loop if successful
                except requests.exceptions.ConnectionError as e:
                    if attempt < 2:  # Check if this is not the last attempt
                        time.sleep(5)  # Wait for 5 seconds before retrying
                        continue  # Retry
                    else:
                        raise e  # Re-raise the exception after all retries
    if comentarios:
        faz_log(
            f"\nProcesso de atualização do banco de dados concluido!", color="blue"
        )


def generate_access_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    token_response = requests.post(token_url, data=token_data)
    token_response.raise_for_status()
    return token_response.json().get("access_token")


def get_drive_and_folder_ids(folder_link, access_token):
    encoded_link = base64.b64encode(folder_link.encode()).decode()
    encoded_link = encoded_link.replace("/", "_").replace("+", "-").replace("=", "")

    headers = {"Authorization": f"Bearer {access_token}"}
    shared_link_url = (
        f"https://graph.microsoft.com/v1.0/shares/u!{encoded_link}/driveItem"
    )
    response = requests.get(shared_link_url, headers=headers)
    response.raise_for_status()
    drive_item = response.json()

    return drive_item["parentReference"]["driveId"], drive_item["id"]


import requests


def generate_access_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    token_response = requests.post(token_url, data=token_data)
    token_response.raise_for_status()
    return token_response.json().get("access_token")


def extract_file_id(file_link):
    # Use a regular expression to extract the file ID from the OneDrive share link
    match = re.search(r"personal/.+?/(\w{10,})\?e=", file_link)
    if match:
        file_id = match.group(1)
        return file_id
    else:
        raise ValueError("Could not extract file ID from the provided OneDrive link.")


# Function to generate access token
def generate_access_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    token_response = requests.post(token_url, data=token_data)
    token_response.raise_for_status()
    return token_response.json().get("access_token")


# Function to get site_id, drive_id, and folder_id
def get_site_drive_folder_ids(folder_link, access_token):
    # Ensure folder_link is a string
    if not isinstance(folder_link, str):
        raise ValueError(
            "Expected a string for 'folder_link', got: " + str(type(folder_link))
        )

    # Encode the folder link to base64
    encoded_link = base64.b64encode(folder_link.encode()).decode()
    encoded_link = encoded_link.replace("/", "_").replace("+", "-").replace("=", "")

    headers = {"Authorization": f"Bearer {access_token}"}

    # Step 1: Get the drive item details from the shared link
    shared_link_url = (
        f"https://graph.microsoft.com/v1.0/shares/u!{encoded_link}/driveItem"
    )
    response = requests.get(shared_link_url, headers=headers)
    response.raise_for_status()
    drive_item = response.json()

    # Extract the drive ID, folder ID, and site ID from the drive item
    drive_id = drive_item["parentReference"]["driveId"]
    folder_id = drive_item["id"]
    site_id = drive_item["parentReference"]["siteId"]

    return site_id, drive_id, folder_id


# Function to upload a file to OneDrive
def upload_file_to_onedrive(
    client_id, client_secret, tenant_id, file_path, folder_link
):
    """
    Uploads a file to a specific folder in OneDrive using Microsoft Graph API.

    Parameters:
    - client_id (str): Azure AD application client ID.
    - client_secret (str): Azure AD application client secret.
    - tenant_id (str): Azure AD tenant ID.
    - file_path (str): The local path to the file to upload.
    - folder_link (str): The shared link to the folder in OneDrive.

    Returns:
    - None: No output to terminal.
    """
    # Generate access token using the provided function
    access_token = generate_access_token(client_id, client_secret, tenant_id)

    # Get site_id, drive_id, and folder_id using the existing function
    site_id, drive_id, folder_id = get_site_drive_folder_ids(folder_link, access_token)

    # Get file name from the path
    file_name = os.path.basename(file_path)

    # Construct the upload URL
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"

    headers = {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/octet-stream",
    }

    # Upload the file
    try:
        with open(file_path, "rb") as f:
            response = requests.put(upload_url, headers=headers, data=f)

        # Suppress all output by not printing or returning anything
    except Exception as e:
        pass  # Suppress all errors and output