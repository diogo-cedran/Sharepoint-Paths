import csv
import time
import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.webs.web import Web

# Configurações
url = 'https://your-sharepoint-site-url'
username = 'your-username@your-domain.com'
password = 'your-password'
output_file = 'sharepoint_paths.csv'
checkpoint_file = 'checkpoint.txt'

# Função para salvar checkpoints
def save_checkpoint(last_processed_path):
    with open(checkpoint_file, 'w', encoding='utf-8') as file:
        file.write(last_processed_path)

# Função para carregar checkpoint
def load_checkpoint():
    if os.path.exists(checkpoint_file):
        with open(checkpoint_file, 'r', encoding='utf-8') as file:
            checkpoint = file.read().strip()
            if checkpoint:
                return checkpoint
    return None

# Função para carregar todos os caminhos já processados do CSV
def load_processed_paths():
    processed_paths = set()
    if os.path.exists(output_file):
        with open(output_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader, None)  # Pular o cabeçalho
            for row in reader:
                if row:
                    processed_paths.add(row[1])
    return processed_paths

# Função para percorrer pastas e arquivos com lógica de retry e paginação
def process_folder(folder, results, ctx, processed_paths, last_processed_path=None, depth=0, max_retries=5, retry_delay=5, batch_size=100):
    folder_path = folder.properties["ServerRelativeUrl"]
    if folder_path in processed_paths or (last_processed_path and folder_path <= last_processed_path):
        return
    results.append((len(folder_path), folder_path))
    print(f"Processing folder: {folder_path} at depth {depth}")

    # Processa arquivos na pasta atual
    folder_files = folder.files
    for _ in range(max_retries):
        try:
            ctx.load(folder_files)
            ctx.execute_query()
            break
        except Exception as e:
            print(f"Error loading files in folder {folder_path}: {e}")
            time.sleep(retry_delay)
    else:
        print(f"Failed to load files in folder {folder_path} after {max_retries} retries")
        return

    for file in folder_files:
        file_path = file.properties["ServerRelativeUrl"]
        if file_path in processed_paths or (last_processed_path and file_path <= last_processed_path):
            continue
        results.append((len(file_path), file_path))
        print(f"Processing file: {file_path}")
        if len(results) >= batch_size:
            save_results(results)
            save_checkpoint(file_path)
            results.clear()

    # Processa subpastas na pasta atual
    folder_sub_folders = folder.folders
    for _ in range(max_retries):
        try:
            ctx.load(folder_sub_folders)
            ctx.execute_query()
            break
        except Exception as e:
            print(f"Error loading subfolders in folder {folder_path}: {e}")
            time.sleep(retry_delay)
    else:
        print(f"Failed to load subfolders in folder {folder_path} after {max_retries} retries")
        return

    for sub_folder in folder_sub_folders:
        process_folder(sub_folder, results, ctx, processed_paths, last_processed_path, depth + 1, max_retries, retry_delay, batch_size)

# Função para salvar resultados em um arquivo CSV
def save_results(results):
    if not results:
        return
    mode = 'a' if os.path.exists(output_file) else 'w'
    with open(output_file, mode, newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if mode == 'w':
            writer.writerow(["Número de Caracteres", "Caminho da Pasta/Arquivo"])
        writer.writerows(results)

# Conecta-se ao SharePoint
context_auth = AuthenticationContext(url)
if context_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(url, context_auth)
    print("Authentication successful")
else:
    print("Authentication failed")
    exit(1)

# Percorre todas as bibliotecas de documentos no site raiz
print("Fetching libraries...")
root_web = ctx.web
ctx.load(root_web)
ctx.execute_query()

lists = root_web.lists
ctx.load(lists)
ctx.execute_query()

results = []
last_processed_path = load_checkpoint()
processed_paths = load_processed_paths()

def process_all_folders_in_list(sp_list, processed_paths, last_processed_path, results):
    root_folder = sp_list.root_folder
    ctx.load(root_folder)
    ctx.execute_query()
    print(f"Processing root folder of list: {sp_list.properties['Title']}")
    process_folder(root_folder, results, ctx, processed_paths, last_processed_path)
    sub_folders = root_folder.folders
    ctx.load(sub_folders)
    ctx.execute_query()
    for sub_folder in sub_folders:
        process_folder(sub_folder, results, ctx, processed_paths, last_processed_path)

for sp_list in lists:
    print(f"Found list: {sp_list.properties['Title']} (BaseTemplate: {sp_list.properties['BaseTemplate']})")
    if sp_list.properties["BaseTemplate"] == 101:  # Bibliotecas de documentos
        process_all_folders_in_list(sp_list, processed_paths, last_processed_path, results)

# Salva os resultados em um arquivo CSV
save_results(results)
print(f"Process completed. Results saved to {output_file}")
