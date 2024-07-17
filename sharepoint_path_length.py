import csv
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.webs.web import Web

# Configurações
url = 'https://your-sharepoint-site-url'
username = 'your_username'
password = 'your_password'
output_file = 'sharepoint_paths.csv'

# Função para percorrer pastas e arquivos
def process_folder(folder, results, ctx):
    folder_path = folder.properties["ServerRelativeUrl"]
    results.append((len(folder_path), folder_path))
    
    # Processa arquivos na pasta atual
    folder_files = folder.files
    ctx.load(folder_files)
    ctx.execute_query()
    for file in folder_files:
        file_path = file.properties["ServerRelativeUrl"]
        results.append((len(file_path), file_path))
    
    # Processa subpastas na pasta atual
    folder_sub_folders = folder.folders
    ctx.load(folder_sub_folders)
    ctx.execute_query()
    for sub_folder in folder_sub_folders:
        process_folder(sub_folder, results, ctx)

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

for sp_list in lists:
    print(f"Found list: {sp_list.properties['Title']} (BaseTemplate: {sp_list.properties['BaseTemplate']})")
    if sp_list.properties["BaseTemplate"] == 101:  # Bibliotecas de documentos
        root_folder = sp_list.root_folder
        ctx.load(root_folder)
        ctx.execute_query()
        print(f"Processing root folder of list: {sp_list.properties['Title']}")
        process_folder(root_folder, results, ctx)

# Salva os resultados em um arquivo CSV
with open(output_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(["Número de Caracteres", "Caminho da Pasta/Arquivo"])
    writer.writerows(results)

print(f"Process completed. Results saved to {output_file}")
