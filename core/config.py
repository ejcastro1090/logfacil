"""
Arquivo: core/config.py
Descrição: Central de configurações globais do LogFácil.
Este arquivo armazena as constantes, parâmetros de funcionamento e configurações
necessárias para a aplicação. Dentre as principais configurações estão:
- Parâmetros do repositório no GitHub para o sistema de atualizações (repo, branch, versão atual).
- Padrões de sistema: raiz padrão de logs (r"C:\Quality\LOG").
- Tunings e limites de performance para leitura de logs (intervalos de escaneamento, polling, tamanho de blocos para início de leitura e limite do buffer exibido).
- Definição dos nomes lógicos dos serviços e seus respectivos executáveis (SERVICE_COMPONENTS) para o recurso de reinicialização de serviços.
"""
import os

# ============================== CONFIGURAÇÃO DO GITHUB ==============================
GITHUB_REPO = "maxinnov-tech/logfacil"
GITHUB_BRANCH = "main"
CURRENT_VERSION = "2.1.1"  # Versão oficial LogFácil Pro
VERSION = CURRENT_VERSION

# ============================== CONFIGURAÇÕES GERAIS ==============================
DEFAULT_ROOT = r"C:\Quality\LOG" if os.name == "nt" else os.path.expanduser("~/Quality/LOG")
SCAN_INTERVAL_SEC = 1.5
TAIL_POLL_INTERVAL_SEC = 0.25
MAX_START_BYTES = 128 * 1024
MAX_VIEW_LINES = 8000
TRIM_BATCH = 600
PAUSED_BUFFER_MAX = 4000

ENCODINGS = ("cp1252", "latin-1", "utf-8")
LOG_EXTENSIONS = ('.log', '.txt', '.out', '.err', '.trace', '.debug')

SERVICE_COMPONENTS = {
    "Integra":                   {"services": ["srvIntegraWeb"],     "tasks": ["IntegraWebService.exe"]},
    "PulserWeb":                 {"services": [],                     "tasks": ["PulserWeb.exe"]},
    "webPostoFiscalServer":      {"services": ["ServicoFiscal"],     "tasks": ["webPostoFiscalServer.exe"]},
    "webPostoLeituraAutomaçao":  {"services": ["ServicoAutomacao"],  "tasks": ["webPostoLeituraAutomacao.exe"]},
    "webPostoPayServer":         {"services": ["webPostoPayServer"],  "tasks": ["webPostoPaySW.exe"]},
    "webPostoPremmialntegracao": {"services": [],                     "tasks": ["webPostoPremiumIntegracao.exe"]},
}
