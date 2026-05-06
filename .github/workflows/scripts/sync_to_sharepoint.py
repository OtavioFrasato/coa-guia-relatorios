#!/usr/bin/env python3
"""
Script para sincronizar arquivo HTML com SharePoint via Microsoft Graph API
Usa credenciais simples (email + senha)
"""

import os
import json
import requests
from pathlib import Path

# Configurações
EMAIL = os.environ.get('SHAREPOINT_EMAIL')
PASSWORD = os.environ.get('SHAREPOINT_PASSWORD')
SITE_URL = "https://tereos.sharepoint.com/sites/ProjetoCOA"
FOLDER_PATH = "/Documents partages/General/Equipe Relatórios/10. Outros/Guia de Informações - COA Relatórios"
FILE_TO_UPLOAD = "index.html"

# Tokens
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_API_URL = "https://graph.microsoft.com/v1.0"

def get_access_token():
    """Obtém token de acesso usando credenciais do usuário"""
    
    print("🔐 Autenticando com SharePoint...")
    
    payload = {
        'grant_type': 'password',
        'client_id': '04b07795-8ddb-461a-bbee-02f9e1bf7b46',  # Microsoft Office App ID
        'scope': 'https://graph.microsoft.com/.default',
        'username': EMAIL,
        'password': PASSWORD,
    }
    
    try:
        response = requests.post(TOKEN_URL, data=payload)
        response.raise_for_status()
        token_data = response.json()
        access_token = token_data.get('access_token')
        
        if not access_token:
            print("❌ Erro: Não foi possível obter token de acesso")
            print(f"Resposta: {token_data}")
            return None
            
        print("✅ Autenticação bem-sucedida!")
        return access_token
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Erro na autenticação: {e}")
        return None

def get_site_id(access_token):
    """Obtém o ID do site SharePoint"""
    
    print(f"🔍 Procurando site: {SITE_URL}")
    
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f"{GRAPH_API_URL}/sites?search=ProjetoCOA"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        sites = response.json().get('value', [])
        
        if not sites:
            print("❌ Site não encontrado")
            return None
            
        site_id = sites[0]['id']
        print(f"✅ Site encontrado: {site_id}")
        return site_id
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Erro ao buscar site: {e}")
        return None

def get_drive_id(access_token, site_id):
    """Obtém o ID do drive (biblioteca) do site"""
    
    print("📁 Procurando drive do site...")
    
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f"{GRAPH_API_URL}/sites/{site_id}/drives"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        drives = response.json().get('value', [])
        
        if not drives:
            print("❌ Drive não encontrado")
            return None
            
        drive_id = drives[0]['id']
        print(f"✅ Drive encontrado: {drive_id}")
        return drive_id
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Erro ao buscar drive: {e}")
        return None

def upload_file(access_token, drive_id):
    """Faz upload do arquivo para SharePoint"""
    
    print(f"📤 Fazendo upload de {FILE_TO_UPLOAD}...")
    
    # Ler arquivo local
    file_path = Path(FILE_TO_UPLOAD)
    if not file_path.exists():
        print(f"❌ Arquivo {FILE_TO_UPLOAD} não encontrado")
        return False
    
    with open(file_path, 'rb') as f:
        file_content = f.read()
    
    # Preparar upload
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Construir caminho remoto
    remote_path = f"{FOLDER_PATH}/{FILE_TO_UPLOAD}".lstrip('/')
    
    # URL para upload
    upload_url = f"{GRAPH_API_URL}/drives/{drive_id}/root:/{remote_path}:/content"
    
    try:
        response = requests.put(upload_url, headers=headers, data=file_content)
        
        if response.status_code in [200, 201]:
            print(f"✅ Arquivo enviado com sucesso!")
            print(f"📍 Localização: {FOLDER_PATH}/{FILE_TO_UPLOAD}")
            return True
        else:
            print(f"❌ Erro no upload: {response.status_code}")
            print(f"Resposta: {response.text}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Erro ao fazer upload: {e}")
        return False

def main():
    """Função principal"""
    
    print("=" * 60)
    print("🔄 Sincronizador: GitHub → SharePoint")
    print("=" * 60)
    
    if not EMAIL or not PASSWORD:
        print("❌ Erro: Variáveis de ambiente não configuradas")
        print("Você precisa definir:")
        print("  - SHAREPOINT_EMAIL")
        print("  - SHAREPOINT_PASSWORD")
        return False
    
    # Autenticar
    access_token = get_access_token()
    if not access_token:
        return False
    
    # Obter site ID
    site_id = get_site_id(access_token)
    if not site_id:
        return False
    
    # Obter drive ID
    drive_id = get_drive_id(access_token, site_id)
    if not drive_id:
        return False
    
    # Upload
    success = upload_file(access_token, drive_id)
    
    print("=" * 60)
    if success:
        print("✅ Sincronização concluída com sucesso!")
    else:
        print("❌ Sincronização falhou")
    print("=" * 60)
    
    return success

if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)
