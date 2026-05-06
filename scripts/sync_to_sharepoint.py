#!/usr/bin/env python3
"""
Sincronizador GitHub → SharePoint via OneDrive API
Usa App Password (mais seguro que senha simples)
"""

import os
import sys
import requests
import base64
from pathlib import Path

# ============================================================================
# CONFIGURAÇÃO
# ============================================================================

EMAIL = os.environ.get('SHAREPOINT_EMAIL')
APP_PASSWORD = os.environ.get('SHAREPOINT_APP_PASSWORD')

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

def upload_via_webdav():
    """Upload via WebDAV (mais confiável)"""
    log("Iniciando upload via WebDAV...", "SYNC")
    
    # Verificar se arquivo existe
    if not Path(FILE_TO_UPLOAD).exists():
        log(f"Arquivo '{FILE_TO_UPLOAD}' não encontrado!", "ERROR")
        return False
    
    # Ler arquivo
    with open(FILE_TO_UPLOAD, 'rb') as f:
        file_data = f.read()
    
    log(f"Tamanho do arquivo: {len(file_data)} bytes", "INFO")
    
    # URL WebDAV do SharePoint
    webdav_url = (
        f"https://{SHAREPOINT_SITE}/sites/{SHAREPOINT_SITE_NAME}/_vti_bin/DavWWWRoot/"
        f"{FOLDER_PATH}/{FILE_TO_UPLOAD}".replace(" ", "%20")
    )
    
    # Credenciais em Base64
    credentials = base64.b64encode(f"{EMAIL}:{APP_PASSWORD}".encode()).decode()
    
    headers = {
        'Authorization': f'Basic {credentials}',
        'Content-Type': 'text/html'
    }
    
    try:
        log(f"Enviando para: {FOLDER_PATH}/{FILE_TO_UPLOAD}", "INFO")
        response = requests.put(webdav_url, headers=headers, data=file_data, timeout=30)
        
        if response.status_code in [200, 201, 204]:
            log("Arquivo enviado com sucesso!", "SUCCESS")
            return True
        else:
            log(f"Erro no upload: {response.status_code}", "ERROR")
            log(f"Resposta: {response.text[:200]}", "ERROR")
            return False
            
    except requests.exceptions.RequestException as e:
        log(f"Erro na requisição: {e}", "ERROR")
        return False

# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 70)
    log("Sincronizador: GitHub → SharePoint (WebDAV)", "SYNC")
    print("=" * 70)
    
    # Validar variáveis de ambiente
    if not EMAIL or not APP_PASSWORD:
        log("Variáveis de ambiente não configuradas!", "ERROR")
        log("Configure: SHAREPOINT_EMAIL e SHAREPOINT_APP_PASSWORD", "ERROR")
        return False
    
    log(f"Email: {EMAIL}", "INFO")
    
    # Upload
    success = upload_via_webdav()
    
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
