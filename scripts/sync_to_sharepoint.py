#!/usr/bin/env python3
"""
Sincronizador GitHub → SharePoint via Microsoft Graph API
Usa autenticação com email/senha (Client Credentials Flow)
"""

import os
import sys
import requests
import json
from pathlib import Path

# ============================================================================
# CONFIGURAÇÃO
# ============================================================================

EMAIL = os.environ.get('SHAREPOINT_EMAIL')
PASSWORD = os.environ.get('SHAREPOINT_PASSWORD')

# Credenciais da aplicação (usando Office App ID padrão)
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
TENANT = "common"

# SharePoint
SHAREPOINT_SITE = "tereos.sharepoint.com"
SHAREPOINT_SITE_NAME = "ProjetoCOA"
FOLDER_PATH = "Documents partages/General/Equipe Relatórios/10. Outros/Guia de Informações - COA Relatórios"
FILE_TO_UPLOAD = "index.html"

# ============================================================================
# FUNÇÕES
# ============================================================================

def log(msg, level="INFO"):
    """Log com emoji"""
    icons = {
        "INFO": "ℹ️",
        "SUCCESS": "✅",
        "ERROR": "❌",
        "WARN": "⚠️",
        "SYNC": "🔄"
    }
    print(f"{icons.get(level, '•')} {msg}")

def get_access_token():
    """Obtém token via ROPC (Resource Owner Password Credentials Flow)"""
    log("Autenticando com Microsoft...", "SYNC")
    
    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    
    data = {
        'grant_type': 'password',
        'client_id': CLIENT_ID,
        'scope': 'https://graph.microsoft.com/.default',
        'username': EMAIL,
        'password': PASSWORD,
    }
    
    try:
        response = requests.post(url, data=data, timeout=10)
        response.raise_for_status()
        
        token = response.json().get('access_token')
        if not token:
            log(f"Sem token na resposta: {response.json()}", "ERROR")
            return None
        
        log("Autenticação bem-sucedida!", "SUCCESS")
        return token
        
    except requests.exceptions.RequestException as e:
        log(f"Erro na autenticação: {e}", "ERROR")
        return None

def upload_file_simple(token):
    """Upload direto via Microsoft Graph"""
    log(f"Fazendo upload de '{FILE_TO_UPLOAD}'...", "SYNC")
    
    # Verificar se arquivo existe
    if not Path(FILE_TO_UPLOAD).exists():
        log(f"Arquivo '{FILE_TO_UPLOAD}' não encontrado!", "ERROR")
        return False
    
    # Ler arquivo
    with open(FILE_TO_UPLOAD, 'rb') as f:
        file_data = f.read()
    
    log(f"Tamanho do arquivo: {len(file_data)} bytes", "INFO")
    
    # URL para upload via SharePoint Drive
    # Formato: /sites/{site}/drives/{drive}/root:/{path}/{filename}:/content
    
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}:/sites/{SHAREPOINT_SITE_NAME}/drives"
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    try:
        log("Procurando drive...", "INFO")
        drive_response = requests.get(drive_url, headers=headers, timeout=10)
        
        if drive_response.status_code != 200:
            log(f"Erro ao buscar drive: {drive_response.text}", "ERROR")
            return False
        
        drives = drive_response.json().get('value', [])
        if not drives:
            log("Nenhum drive encontrado", "ERROR")
            return False
        
        drive_id = drives[0]['id']
        log(f"Drive encontrado: {drive_id}", "SUCCESS")
        
        # Fazer upload
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}/{FILE_TO_UPLOAD}:/content"
        
        upload_headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'text/html'
        }
        
        log(f"Enviando para: {FOLDER_PATH}/{FILE_TO_UPLOAD}", "INFO")
        upload_response = requests.put(upload_url, headers=upload_headers, data=file_data, timeout=30)
        
        if upload_response.status_code in [200, 201]:
            log("Arquivo enviado com sucesso!", "SUCCESS")
            result = upload_response.json()
            log(f"ID do arquivo: {result.get('id')}", "INFO")
            return True
        else:
            log(f"Erro no upload: {upload_response.status_code}", "ERROR")
            log(f"Resposta: {upload_response.text[:500]}", "ERROR")
            return False
            
    except requests.exceptions.RequestException as e:
        log(f"Erro na requisição: {e}", "ERROR")
        return False

# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 70)
    log("Sincronizador: GitHub → SharePoint (Microsoft Graph)", "SYNC")
    print("=" * 70)
    
    # Validar variáveis de ambiente
    if not EMAIL or not PASSWORD:
        log("Variáveis de ambiente não configuradas!", "ERROR")
        log("Configure: SHAREPOINT_EMAIL e SHAREPOINT_PASSWORD", "ERROR")
        return False
    
    log(f"Email: {EMAIL}", "INFO")
    
    # Autenticar
    token = get_access_token()
    if not token:
        log("Falha na autenticação", "ERROR")
        return False
    
    # Upload
    success = upload_file_simple(token)
    
    print("=" * 70)
    if success:
        log("Sincronização concluída com sucesso!", "SUCCESS")
        print("=" * 70)
        return True
    else:
        log("Sincronização falhou", "ERROR")
        print("=" * 70)
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
