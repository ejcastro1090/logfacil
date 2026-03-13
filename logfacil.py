import os
import sys
import threading
import time
import queue
import datetime
import subprocess
import ctypes
import urllib.request
import json
import shutil
from dataclasses import dataclass, field
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.simpledialog as simpledialog

try:
    import ttkbootstrap as tb
except Exception:
    tb = None

# ============================== VERIFICAÇÃO DE ADMIN ==============================

def is_admin():
    """Verifica se o programa está rodando como administrador."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Reinicia o programa com privilégios de administrador."""
    try:
        script = os.path.abspath(sys.argv[0])
        
        if script.endswith('.py'):
            executable = sys.executable
            params = f'"{script}"'
        else:
            executable = script
            params = ' '.join(sys.argv[1:])
        
        ctypes.windll.shell32.ShellExecuteW(
            None, 
            "runas",
            executable, 
            params, 
            None, 
            1
        )
        
        sys.exit(0)
    except Exception as e:
        print(f"Erro ao tentar elevar privilégios: {e}")
        return False

# ============================== AUTO-UPDATE E VERSÃO ==============================
__version__ = "1.0.0"
GITHUB_REPO = "ejcastro1090/logfacil"

def check_for_updates(app_instance):
    """Verifica no GitHub se há uma nova versão disponível."""
    def worker():
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(url, headers={'User-Agent': 'LogFacil-Updater'})
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
                
            latest_version = data.get("tag_name", "").lstrip("v")
            
            # Comparação simples de versão
            if latest_version and latest_version != __version__:
                # Versão nova disponível
                assets = data.get("assets", [])
                exe_url = None
                for asset in assets:
                    if asset["name"].endswith(".exe"):
                        exe_url = asset["browser_download_url"]
                        break
                
                if exe_url:
                    app_instance.root.after(0, lambda: ask_for_update(app_instance, latest_version, exe_url))
                else:
                    app_instance.root.after(0, lambda: messagebox.showinfo("Atualização", f"Nova versão {latest_version} encontrada no GitHub, mas nenhum executável (.exe) está disponível para download."))
            else:
                app_instance.root.after(0, lambda: messagebox.showinfo("Atualização", "Você já está na versão mais recente!"))
                
        except Exception as e:
            app_instance.root.after(0, lambda: messagebox.showerror("Erro de Atualização", f"Não foi possível verificar atualizações:\n{e}"))
            
    threading.Thread(target=worker, daemon=True).start()

def ask_for_update(app_instance, new_version, exe_url):
    """Pergunta ao usuário se deseja atualizar e inicia o download."""
    if messagebox.askyesno("Nova Atualização Disponível", f"A versão {new_version} está disponível!\n\nDeseja baixar e atualizar agora?"):
        if hasattr(app_instance, 'btn_update'):
            app_instance.btn_update.configure(state="disabled", text="⏳ Baixando...")
        
        def download_worker():
            try:
                # Verifica se está rodando como executável
                is_frozen = getattr(sys, 'frozen', False)
                if is_frozen:
                    exe_path = sys.executable
                else:
                    app_instance.root.after(0, lambda: messagebox.showwarning("Modo Desenvolvedor", "Rodando como script .py. O download do .exe será feito mas não substituirá o script atual."))
                    exe_path = os.path.abspath(sys.argv[0])
                
                download_dir = os.path.dirname(exe_path)
                new_exe_path = os.path.join(download_dir, "LogFacil_update.exe")
                
                # Faz o download
                req = urllib.request.Request(exe_url, headers={'User-Agent': 'LogFacil-Updater'})
                with urllib.request.urlopen(req, timeout=60) as response, open(new_exe_path, 'wb') as out_file:
                    shutil.copyfileobj(response, out_file)
                
                # Aplica a atualização
                app_instance.root.after(0, lambda: apply_update(app_instance, exe_path, new_exe_path, is_frozen))
                
            except Exception as e:
                app_instance.root.after(0, lambda: messagebox.showerror("Erro de Download", f"Falha ao baixar a atualização:\n{e}"))
                if hasattr(app_instance, 'btn_update'):
                    app_instance.root.after(0, lambda: app_instance.btn_update.configure(state="normal", text="📥 Atualizar"))
                
        threading.Thread(target=download_worker, daemon=True).start()

def apply_update(app_instance, current_exe, new_exe, is_frozen):
    """Cria um script batch para substituir o arquivo e reinicia."""
    if not is_frozen:
        if hasattr(app_instance, 'btn_update'):
            app_instance.btn_update.configure(state="normal", text="📥 Atualizar")
        messagebox.showinfo("Atualização", f"Download concluído!\n(Modo script: arquivo '{os.path.basename(new_exe)}' salvo na pasta, mas a troca automática é só para .exe)")
        return
        
    try:
        # Cria um arquivo .bat temporário para substituir o executável
        bat_path = os.path.join(os.path.dirname(current_exe), "update_logfacil.bat")
        with open(bat_path, "w") as bat_file:
            bat_file.write(f'''@echo off
timeout /t 2 /nobreak >nul
del "{current_exe}"
ren "{new_exe}" "{os.path.basename(current_exe)}"
start "" "{current_exe}"
del "%~f0"
''')
        
        messagebox.showinfo("Atualização Concluída", "O aplicativo será reiniciado para aplicar a atualização.")
        
        # Executa o batch sem console
        CREATE_NO_WINDOW = 0x08000000
        subprocess.Popen(["cmd.exe", "/c", bat_path], creationflags=CREATE_NO_WINDOW)
        
        # Fecha a aplicação atual
        app_instance._on_close()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao preparar o script de atualização:\n{e}")
        if hasattr(app_instance, 'btn_update'):
            app_instance.btn_update.configure(state="normal", text="📥 Atualizar")


# ============================== CONFIGURAÇÕES ==============================

DEFAULT_ROOT = r"C:\Quality\LOG" if os.name == "nt" else os.path.expanduser("~/Quality/LOG")
SCAN_INTERVAL_SEC = 1.5
TAIL_POLL_INTERVAL_SEC = 0.25
MAX_START_BYTES = 128 * 1024
MAX_VIEW_LINES = 8000
TRIM_BATCH = 600
PAUSED_BUFFER_MAX = 4000
ENCODINGS = ("cp1252", "latin-1")
LOG_EXTENSIONS = ('.log', '.txt', '.out', '.err', '.trace', '.debug')

# ============================== MAPEAMENTO SERVIÇO -> TASKS E SERVIÇOS ==============================

SERVICE_COMPONENTS = {
    "Integra": {
        "services": ["srvIntegraWeb"],
        "tasks": ["IntegraWebService.exe"]
    },

    "PulserWeb": {
        "services": ["PulserWebService"],
        "tasks": ["PulserWeb.exe", "PulserWeb.Host.exe"]

    },
    "webPostoFiscalServer": {
        "services": ["WebPostoFiscalServer"],
        "tasks": ["webPostoFiscalServer.exe"]

    },
    "webPostoLeituraAutomaçao": {
        "services": ["LeituraAutomacaoService"],
        "tasks": ["webPostoLeituraAutomacao.exe", "LeituraAutomacao.exe"]
    },
    "webPostoPayServer": {
        "services": ["PayServerService"],
        "tasks": ["webPostoPayServer.exe", "PayServer.exe"]
    },
    "webPostoPremmialntegracao": {
        "services": ["PremiumIntegracaoService"],
        "tasks": ["webPostoPremiumIntegracao.exe", "PremiumIntegracao.exe"]
    },
}

def get_default_components(service_name):
    """Gera componentes padrão para serviços não mapeados."""
    return {
        "services": [f"{service_name}Service", service_name],
        "tasks": [f"{service_name}.exe"]
    }


# ============================== FUNÇÕES UTILITÁRIAS ==============================

def service_from_path(path: str):
    """Extrai o nome do serviço do caminho do arquivo de log."""
    parts = os.path.normpath(path).split(os.sep)
    for i, p in enumerate(parts):
        if p.lower() == "log" and i + 1 < len(parts):
            return parts[i + 1]
    return "?"


def check_admin_and_warn():
    """Verifica se é admin e retorna mensagem."""
    if os.name == "nt" and not is_admin():
        return False, "⚠ Sem privilégios de administrador"
    return True, "✓ Modo administrador"


def stop_windows_service(service_name):
    """
    Para um serviço Windows usando NET STOP.
    """
    is_admin_flag, admin_msg = check_admin_and_warn()
    
    try:
        result = subprocess.run(
            ["net", "stop", service_name],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            return True, f"Serviço '{service_name}' parado com sucesso"
        else:
            output = (result.stdout + result.stderr).lower()
            if "não está iniciado" in output or "not started" in output:
                return True, f"Serviço '{service_name}' já estava parado"
            elif "acesso negado" in output or "access denied" in output:
                return False, f"❌ Acesso negado! {admin_msg}"
            else:
                error_msg = result.stderr or result.stdout or "Erro desconhecido"
                return False, f"Falha ao parar serviço '{service_name}': {error_msg}"
                
    except subprocess.TimeoutExpired:
        return False, f"Timeout ao tentar parar serviço '{service_name}'"
    except Exception as e:
        return False, f"Erro ao parar serviço '{service_name}': {str(e)}"


def start_windows_service(service_name):
    """
    Inicia um serviço Windows usando NET START.
    """
    is_admin_flag, admin_msg = check_admin_and_warn()
    
    try:
        result = subprocess.run(
            ["net", "start", service_name],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            return True, f"Serviço '{service_name}' iniciado com sucesso"
        else:
            output = (result.stdout + result.stderr).lower()
            if "já foi iniciado" in output or "already been started" in output:
                return True, f"Serviço '{service_name}' já estava em execução"
            elif "acesso negado" in output or "access denied" in output:
                return False, f"❌ Acesso negado! {admin_msg}"
            else:
                error_msg = result.stderr or result.stdout or "Erro desconhecido"
                return False, f"Falha ao iniciar serviço '{service_name}': {error_msg}"
                
    except subprocess.TimeoutExpired:
        return False, f"Timeout ao tentar iniciar serviço '{service_name}'"
    except Exception as e:
        return False, f"Erro ao iniciar serviço '{service_name}': {str(e)}"


def restart_windows_service(service_name):
    """
    Reinicia um serviço Windows usando NET STOP e NET START.
    """
    results = []
    
    stop_success, stop_msg = stop_windows_service(service_name)
    results.append(stop_msg)
    
    time.sleep(2)
    
    start_success, start_msg = start_windows_service(service_name)
    results.append(start_msg)
    
    overall_success = stop_success and start_success
    return overall_success, "\n".join(results)


def kill_task(task_name):
    """
    Finaliza uma task/processo pelo nome usando TASKKILL.
    """
    is_admin_flag, admin_msg = check_admin_and_warn()
    
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", task_name],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            return True, f"✅ {task_name}: finalizado com sucesso"
        elif "not found" in result.stderr.lower():
            return True, f"⚠ {task_name}: processo não encontrado (ignorado)"
        elif "acesso negado" in result.stderr.lower() or "access denied" in result.stderr.lower():
            return False, f"❌ {task_name}: acesso negado! {admin_msg}"
        else:
            return False, f"❌ {task_name}: erro {result.stderr}"
            
    except subprocess.TimeoutExpired:
        return False, f"❌ {task_name}: timeout ao tentar finalizar"
    except Exception as e:
        return False, f"❌ {task_name}: {str(e)}"


def check_service_status(service_name):
    """
    Verifica se um serviço existe e seu status.
    """
    try:
        result = subprocess.run(
            ["sc", "query", service_name],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode != 0:
            return False, False, f"Serviço '{service_name}' não encontrado"
        
        output = result.stdout
        if "RUNNING" in output:
            return True, True, f"Serviço '{service_name}' está em execução"
        elif "STOPPED" in output:
            return True, False, f"Serviço '{service_name}' está parado"
        else:
            return True, False, f"Serviço '{service_name}' status desconhecido"
            
    except Exception as e:
        return False, False, f"Erro ao verificar serviço '{service_name}': {str(e)}"


def restart_service_components(service_name: str):
    """
    Reinicia todos os componentes de um serviço.
    """
    if os.name != "nt":
        return False, ["Funcionalidade disponível apenas no Windows"]
    
    components = SERVICE_COMPONENTS.get(service_name, get_default_components(service_name))
    
    results = []
    results.append(f"🔄 Reiniciando {service_name}...")
    results.append("-" * 40)
    
    # Mostra status de admin
    is_admin_flag, admin_msg = check_admin_and_warn()
    results.append(f"🔑 Status: {admin_msg}")
    results.append("")
    
    all_success = True
    
    # 1. Reinicia serviços Windows
    services = components.get("services", [])
    if services:
        results.append(f"📌 Serviços Windows:")
        for svc in reversed(services):
            exists, running, status_msg = check_service_status(svc)
            results.append(f"  {status_msg}")
            
            if exists:
                success, msg = restart_windows_service(svc)
                results.append(f"  → {msg}")
                if not success:
                    all_success = False
            else:
                results.append(f"  → ⚠ Serviço não encontrado, ignorado")
    else:
        results.append(f"📌 Nenhum serviço Windows configurado")
    
    time.sleep(1)
    
    # 2. Finaliza tasks/processos
    tasks = components.get("tasks", [])
    if tasks:
        results.append(f"\n📌 Processos/Tasks:")
        for task in tasks:
            success, msg = kill_task(task)
            results.append(f"  {msg}")
            if not success:
                all_success = False
    else:
        results.append(f"\n📌 Nenhum processo configurado")
    
    results.append("")
    if not is_admin_flag:
        results.append("⚠ ATENÇÃO: Execute como Administrador para evitar erros de permissão!")
    results.append("✅ Processo concluído!" if all_success else "⚠ Concluído com alguns erros")
    
    return all_success, results


# ============================== FUNÇÕES DE SCAN ==============================

def scan_log_files(root: str):
    """Escaneia recursivamente arquivos de log."""
    files_by_service = {}
    
    if not os.path.isdir(root):
        return files_by_service
    
    root_name = os.path.basename(os.path.normpath(root)).lower()
    scan_root = root if root_name == "log" else os.path.join(root, "LOG")
    
    if not os.path.isdir(scan_root):
        return files_by_service
    
    try:
        service_folders = [f for f in os.listdir(scan_root) 
                          if os.path.isdir(os.path.join(scan_root, f))]
        
        for service in service_folders:
            service_dir = os.path.join(scan_root, service)
            
            latest_file = None
            latest_mtime = 0
            
            for root_dir, dirs, files in os.walk(service_dir):
                for file in files:
                    if file.lower().endswith(LOG_EXTENSIONS):
                        filepath = os.path.join(root_dir, file)
                        try:
                            mtime = os.path.getmtime(filepath)
                            if mtime > latest_mtime:
                                latest_mtime = mtime
                                latest_file = filepath
                        except Exception:
                            continue
            
            if latest_file:
                files_by_service[service] = (latest_file, latest_mtime)
                
    except Exception:
        pass
    
    return files_by_service


def find_latest_by_service(root: str):
    """Encontra o arquivo de log mais recente para cada serviço."""
    latest = {}
    files_by_service = scan_log_files(root)
    
    for svc, (path, mtime) in files_by_service.items():
        latest[svc] = (path, mtime)
    
    return {svc: tup[0] for svc, tup in latest.items()}


def open_text_auto(path):
    """Tenta abrir arquivo com diferentes encodings."""
    last_exc = None
    for enc in ENCODINGS:
        try:
            return open(path, "r", encoding=enc, errors="replace"), True
        except Exception as e:
            last_exc = e
    
    try:
        return open(path, "rb", buffering=0), False
    except Exception as e:
        raise last_exc or e


def seek_tail(fobj, max_bytes):
    """Posiciona o ponteiro do arquivo próximo ao final."""
    try:
        fobj.seek(0, os.SEEK_END)
        size = fobj.tell()
        start = max(size - max_bytes, 0)
        fobj.seek(start, os.SEEK_SET)
        if start > 0:
            _ = fobj.readline()
    except Exception:
        pass


# ============================== CLASSE LOGTAB ==============================

@dataclass
class LogTab:
    """Aba de visualização de log."""
    
    app: "App"
    filepath: str
    frame: ttk.Frame = field(init=False)
    text: tk.Text = field(init=False)
    vsb: ttk.Scrollbar = field(init=False)
    hsb: ttk.Scrollbar = field(init=False)
    btn_frame: ttk.Frame = field(init=False)
    btn_restart: ttk.Button = field(init=False)
    stop_event: threading.Event = field(default_factory=threading.Event, init=False)
    q: "queue.Queue[str]" = field(default_factory=queue.Queue, init=False)
    follow: bool = field(default=True, init=False)
    unread: int = field(default=0, init=False)
    appended_since_trim: int = field(default=0, init=False)
    paused_buffer: list = field(default_factory=list, init=False)
    service_name: str = field(init=False)

    def __post_init__(self):
        self.service_name = service_from_path(self.filepath)
        self.frame = ttk.Frame(self.app.notebook)
        self._build_ui()
        self._start_tail()
        self._schedule_drain()

    def _with_unlocked(self, fn, *args, **kwargs):
        """Executa função com o text widget desbloqueado."""
        prev = self.text.cget("state")
        try:
            if prev != "normal":
                self.text.configure(state="normal")
            return fn(*args, **kwargs)
        finally:
            if prev != "normal":
                self.text.configure(state="disabled")

    # ---------------------------- Construção da UI ----------------------------
    
    def _build_ui(self):
        """Constrói a interface da aba."""
        # Frame superior com botões
        self.btn_frame = ttk.Frame(self.frame)
        self.btn_frame.pack(side="top", fill="x", padx=2, pady=2)
        
        # Botão de reiniciar serviço
        self.btn_restart = ttk.Button(
            self.btn_frame, 
            text=f"🔄 Reiniciar {self.service_name}",
            command=self._restart_service,
            style="danger.TButton" if tb else None
        )
        self.btn_restart.pack(side="left", padx=2)
        
        # Separador
        ttk.Separator(self.btn_frame, orient="vertical").pack(side="left", padx=5, fill="y")
        
        # Label de status
        self.status_label = ttk.Label(self.btn_frame, text="", foreground="gray")
        self.status_label.pack(side="left", padx=5)
        
        # Frame para o texto e scrollbars
        text_frame = ttk.Frame(self.frame)
        text_frame.pack(side="top", fill="both", expand=True)
        
        self.text = tk.Text(text_frame, wrap="none", undo=False)
        
        try:
            self.text.configure(font=("Consolas", 10))
        except Exception:
            self.text.configure(font=("TkFixedFont", 10))
        
        # Scrollbars
        self.vsb = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.vsb.pack(side="right", fill="y")
        
        self.hsb = ttk.Scrollbar(text_frame, orient="horizontal", command=self.text.xview)
        self.hsb.pack(side="bottom", fill="x")
        
        self.text.configure(xscrollcommand=self.hsb.set)
        self.text.pack(side="left", fill="both", expand=True)
        self.text.configure(state="disabled")

        # Configuração do scroll vertical
        def yscroll(first, last):
            self.vsb.set(first, last)
            at_bottom = abs(1.0 - float(last)) < 1e-3
            self.follow = at_bottom
            if at_bottom:
                self._flush_buffer()
                self._update_tab_label()
        
        self.text.configure(yscrollcommand=yscroll)

        # Eventos de scroll do usuário
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>", 
                   "<Key-Up>", "<Key-Down>", "<Prior>", "<Next>", 
                   "<End>", "<Home>"):
            self.text.bind(ev, self._on_user_scroll, add="+")

        # Menu de contexto
        self._build_context_menu()
        
        # Atalhos de teclado
        self._bind_shortcuts()
        
        # Bloqueia edição
        self._disable_editing()

    def _restart_service(self):
        """Callback para reiniciar o serviço."""
        def do_restart():
            # Desabilita o botão durante o processo
            self.btn_restart.configure(state="disabled", text=f"⏳ Reiniciando {self.service_name}...")
            self.status_label.configure(text="Reiniciando serviços e processos...", foreground="orange")
            
            # Executa o restart em thread separada para não travar a UI
            def restart_thread():
                success, messages = restart_service_components(self.service_name)
                
                # Atualiza UI na thread principal
                self.frame.after(0, lambda: self._restart_callback(success, messages))
            
            threading.Thread(target=restart_thread, daemon=True).start()
        
        # Confirma com o usuário
        if messagebox.askyesno(
            "Confirmar reinicialização", 
            f"Deseja realmente reiniciar o serviço '{self.service_name}'?\n\n"
            "Isso irá:\n"
            "• Reiniciar serviços Windows relacionados (usando NET)\n"
            "• Finalizar processos relacionados (usando TASKKILL)\n"
            "• Aguardar reinicialização automática"
        ):
            do_restart()

    def _restart_callback(self, success: bool, messages: list):
        """Callback após tentativa de restart."""
        # Reabilita o botão
        self.btn_restart.configure(state="normal", text=f"🔄 Reiniciar {self.service_name}")
        
        # Formata mensagens para display
        message_text = "\n".join(messages)
        
        if success:
            self.status_label.configure(text="✅ Serviço reiniciado com sucesso!", foreground="green")
            messagebox.showinfo(
                "Sucesso", 
                f"Serviço '{self.service_name}' reiniciado com sucesso!\n\n{message_text}"
            )
        else:
            self.status_label.configure(text="⚠ Reiniciado com alguns erros", foreground="orange")
            messagebox.showwarning(
                "Atenção", 
                f"Serviço '{self.service_name}' reiniciado com alguns problemas:\n\n{message_text}"
            )
        
        # Limpa a mensagem de status após 5 segundos
        self.frame.after(5000, lambda: self.status_label.configure(text=""))

    def _build_context_menu(self):
        """Constrói o menu de contexto."""
        menu = tk.Menu(self.text, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: self.text.event_generate("<<Copy>>"))
        menu.add_command(label="Selecionar tudo", 
                        command=lambda: self._with_unlocked(self.text.tag_add, "sel", "1.0", "end-1c"))
        menu.add_separator()
        menu.add_command(label="Pausar/Seguir    F2", command=self.toggle_follow)
        menu.add_command(label="Ir para o fim", command=self.scroll_to_end)
        menu.add_separator()
        menu.add_command(label="Reiniciar serviço", command=self._restart_service)
        menu.add_separator()
        menu.add_command(label="Filtrar…    Ctrl+F", command=self._aplicar_filtro)
        menu.add_command(label="Limpar filtro    Ctrl+L", command=self._limpar_filtro)
        
        self.text.bind("<Button-3>", lambda e, m=menu: m.tk_popup(e.x_root, e.y_root))

    def _aplicar_filtro(self, event=None):
        """Aplica filtro de texto."""
        termo = simpledialog.askstring("Filtro", "Digite o termo a filtrar:")
        if not termo:
            return
        
        def _apply():
            self.text.tag_remove("filtro", "1.0", "end")
            idx = "1.0"
            while True:
                idx = self.text.search(termo, idx, nocase=1, stopindex="end")
                if not idx:
                    break
                fim = f"{idx}+{len(termo)}c"
                self.text.tag_add("filtro", idx, fim)
                idx = fim
            self.text.tag_config("filtro", background="yellow", foreground="black")
        
        self._with_unlocked(_apply)

    def _limpar_filtro(self, event=None):
        """Limpa o filtro aplicado."""
        self._with_unlocked(self.text.tag_remove, "filtro", "1.0", "end")

    def _bind_shortcuts(self):
        """Configura atalhos de teclado."""
        self.text.bind("<F2>", lambda e: (self.toggle_follow(), "break"))
        self.text.bind("<Control-f>", lambda e: (self._aplicar_filtro(), "break"))
        self.text.bind("<Control-l>", lambda e: (self._limpar_filtro(), "break"))
        self.text.bind("<Command-f>", lambda e: (self._aplicar_filtro(), "break"))
        self.text.bind("<Command-l>", lambda e: (self._limpar_filtro(), "break"))

    def _disable_editing(self):
        """Desabilita edição do texto."""
        for seq in ("<<Paste>>", "<Control-v>", "<Shift-Insert>", 
                    "<<Cut>>", "<Control-x>", "<Shift-Delete>"):
            self.text.bind(seq, lambda e: "break")

    # ---------------------------- Eventos de UI ----------------------------
    
    def _on_user_scroll(self, _event=None):
        """Callback quando usuário faz scroll."""
        self.text.after(30, self._update_follow_from_view)

    def _update_follow_from_view(self):
        """Atualiza estado de follow baseado na posição de visualização."""
        _, last = self.text.yview()
        at_bottom = abs(1.0 - float(last)) < 1e-3
        self.follow = at_bottom
        if at_bottom:
            self._flush_buffer()
            self._update_tab_label()

    def toggle_follow(self):
        """Alterna modo follow."""
        self.follow = not self.follow
        if self.follow:
            self.scroll_to_end()
            self._flush_buffer()
        self._update_tab_label()

    def scroll_to_end(self):
        """Rola para o final do texto."""
        self.text.see("end")
        self._update_follow_from_view()

    def _update_tab_label(self):
        """Atualiza o rótulo da aba."""
        base = f"{self.service_name} — {os.path.basename(self.filepath)}"
        if self.unread > 0 and not self.follow:
            base += f"  ({self.unread} novos)"
        self.app.notebook.tab(self.frame, text=base)

    # ---------------------------- Thread de Tail ----------------------------
    
    def _start_tail(self):
        """Inicia thread de monitoramento do arquivo."""
        threading.Thread(target=self._tail_loop, daemon=True).start()

    def _open_file(self):
        """Abre o arquivo e posiciona no final."""
        f, is_text = open_text_auto(self.filepath)
        seek_tail(f, MAX_START_BYTES)
        return f, is_text

    def _tail_loop(self):
        """Loop principal de monitoramento do arquivo."""
        try:
            f, is_text = self._open_file()
        except Exception as e:
            self.q.put(f"[ERRO] Falha ao abrir: {e}\n")
            return
        
        decoder = (lambda b: b.decode("utf-8", errors="replace")) if not is_text else None
        
        while not self.stop_event.is_set():
            try:
                chunk = f.read()
                if chunk:
                    data = chunk if is_text else decoder(chunk)
                    self.q.put(data)
                else:
                    time.sleep(TAIL_POLL_INTERVAL_SEC)
                    try:
                        size = os.path.getsize(self.filepath)
                        pos = f.tell()
                        if size < pos:  # Arquivo truncado/rotacionado
                            f.close()
                            f, is_text = self._open_file()
                            decoder = (lambda b: b.decode("utf-8", errors="replace")) if not is_text else None
                    except Exception:
                        pass
            except Exception as e:
                self.q.put(f"\n[ERRO de leitura] {e}\n")
                time.sleep(TAIL_POLL_INTERVAL_SEC)
        
        try:
            f.close()
        except Exception:
            pass

    # ---------------------------- Atualização da UI ----------------------------
    
    def _schedule_drain(self):
        """Agenda drenagem da fila."""
        if self.stop_event.is_set():
            return
        self._drain()
        self.text.after(90, self._schedule_drain)

    def _drain(self):
        """Processa itens da fila."""
        budget = 400
        agg = []
        total_newlines = 0
        
        while budget:
            try:
                item = self.q.get_nowait()
            except queue.Empty:
                break
            agg.append(item)
            total_newlines += item.count("\n") or 1
            budget -= 1

        if not agg:
            return

        data = "".join(agg)
        if self.follow:
            self._append(data)
        else:
            self.paused_buffer.append(data)
            if len(self.paused_buffer) > PAUSED_BUFFER_MAX:
                self.paused_buffer = self.paused_buffer[-PAUSED_BUFFER_MAX:]
            self.unread += total_newlines
            self._update_tab_label()

    def _append(self, data: str):
        """Adiciona texto ao widget."""
        if not data:
            return
        
        def _do():
            self.text.insert("end", data)
            self.appended_since_trim += data.count("\n")
            if self.appended_since_trim >= TRIM_BATCH:
                self._trim()
                self.appended_since_trim = 0
            if self.follow:
                self.text.see("end")
        
        self._with_unlocked(_do)

    def _trim(self):
        """Remove linhas excedentes."""
        def _do():
            end_index = self.text.index("end-1c")
            total_lines = int(end_index.split(".")[0])
            if total_lines > MAX_VIEW_LINES:
                excess = total_lines - MAX_VIEW_LINES
                self.text.delete("1.0", f"{excess+1}.0")
        
        self._with_unlocked(_do)

    def _flush_buffer(self):
        """Descarrega buffer pausado."""
        if not self.paused_buffer:
            return
        combo = "".join(self.paused_buffer)
        self.paused_buffer.clear()
        self.unread = 0
        if combo:
            self._append(combo)


# ============================== CLASSE FOLDERWATCHER ==============================

class FolderWatcher(threading.Thread):
    """Monitora pasta por novos arquivos de log."""
    
    def __init__(self, app: "App", root_dir: str):
        super().__init__(daemon=True)
        self.app = app
        self.root_dir = root_dir
        self.stop_event = threading.Event()
        self.latest_by_service = {}

    def run(self):
        self._scan_and_open(initial=True)
        while not self.stop_event.is_set():
            self._scan_and_open(initial=False)
            for _ in range(int(SCAN_INTERVAL_SEC / 0.1)):
                if self.stop_event.is_set():
                    break
                time.sleep(0.1)

    def _scan_and_open(self, initial=False):
        """Escaneia e abre novos arquivos."""
        latest_map = find_latest_by_service(self.root_dir)
        
        for svc, path in latest_map.items():
            prev = self.latest_by_service.get(svc)
            if prev is None:
                self.latest_by_service[svc] = path
                self.app.enqueue_open(path)
            elif prev != path:
                self.latest_by_service[svc] = path
                self.app.enqueue_switch_service_log(svc, path)


# ============================== CLASSE APP ==============================

class App:
    """Aplicação principal."""
    
    def __init__(self):
        self._setup_window()
        self._build_topbar()
        self._setup_notebook()
        self._setup_queues()
        self._start_watcher()
        self._setup_close_handler()

    def _setup_window(self):
        """Configura a janela principal."""
        if tb is not None:
            try:
                self.root = tb.Window(themename="darkly")
            except Exception:
                self.root = tk.Tk()
        else:
            self.root = tk.Tk()

        self.root.title(f"LogFácil v{__version__} - Monitor de Serviços")
        self.root.geometry("1100x700")

    def _build_topbar(self):
        """Constrói a barra superior."""
        bar = ttk.Frame(self.root)
        bar.pack(side="top", fill="x", padx=6, pady=(6, 0))
        
        ttk.Label(bar, text="Raiz LOG:").pack(side="left")
        
        self.entry = ttk.Entry(bar, width=60)
        self.entry.pack(side="left", padx=4)
        self.entry.insert(0, DEFAULT_ROOT)
        
        ttk.Button(bar, text="Escolher pasta…", 
                  command=self._choose_root).pack(side="left", padx=4)
        
        # Botão para atualizar versão
        self.btn_update = ttk.Button(bar, text="📥 Verificar Atualização", 
                  command=lambda: check_for_updates(self))
        self.btn_update.pack(side="right", padx=4)
        
        # Botão para reiniciar todos os serviços
        ttk.Button(bar, text="🔄 Reiniciar Todos", 
                  command=self._restart_all_services).pack(side="right", padx=4)

    def _restart_all_services(self):
        """Reinicia todos os serviços monitorados."""
        if not self.open_tabs:
            messagebox.showinfo("Info", "Nenhum serviço sendo monitorado.")
            return
        
        services = set(tab.service_name for tab in self.open_tabs.values())
        
        if messagebox.askyesno(
            "Confirmar", 
            f"Deseja reiniciar TODOS os {len(services)} serviços?\n\n" +
            "\n".join(f"• {s}" for s in sorted(services))
        ):
            # Desabilita todos os botões de restart
            for tab in self.open_tabs.values():
                tab.btn_restart.configure(state="disabled")
            
            def restart_all_thread():
                all_results = []
                for service in sorted(services):
                    success, messages = restart_service_components(service)
                    status = "✅" if success else "⚠"
                    all_results.append(f"{status} {service}:")
                    all_results.extend(f"  {msg}" for msg in messages)
                    all_results.append("")  # linha em branco
                
                # Reabilita botões e mostra resultado
                self.root.after(0, lambda: self._restart_all_callback(all_results))
            
            threading.Thread(target=restart_all_thread, daemon=True).start()

    def _restart_all_callback(self, results):
        """Callback após reiniciar todos os serviços."""
        # Reabilita todos os botões
        for tab in self.open_tabs.values():
            tab.btn_restart.configure(state="normal")
        
        # Mostra resultados
        messagebox.showinfo(
            "Reinicialização em Massa",
            "\n".join(results)
        )

    def _setup_notebook(self):
        """Configura o notebook de abas."""
        self.notebook = (tb.Notebook(self.root, bootstyle="info") 
                        if tb else ttk.Notebook(self.root))
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

    def _setup_queues(self):
        """Configura as filas de comunicação."""
        self.open_tabs = {}
        self.tab_by_service = {}
        self.open_queue = queue.Queue()
        self.switch_queue = queue.Queue()
        self.watcher = None
        self.root.after(120, self._consume_queues)

    def _setup_close_handler(self):
        """Configura handler de fechamento."""
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------- Gerenciamento do Watcher ----------------------------
    
    def _choose_root(self):
        """Abre diálogo para escolher pasta raiz."""
        path = filedialog.askdirectory(
            initialdir=self.entry.get() or DEFAULT_ROOT,
            title="Selecione a pasta LOG (ex: C:\\Quality\\LOG)"
        )
        if path:
            self.entry.delete(0, "end")
            self.entry.insert(0, path)
            self._restart_watcher()

    def _restart_watcher(self):
        """Reinicia o watcher."""
        self._stop_watcher()
        self._start_watcher()

    def _start_watcher(self):
        """Inicia o watcher."""
        root_dir = self.entry.get().strip() or DEFAULT_ROOT
        self.watcher = FolderWatcher(self, root_dir)
        self.watcher.start()

    def _stop_watcher(self):
        """Para o watcher."""
        if self.watcher:
            self.watcher.stop_event.set()
            self.watcher = None

    # ---------------------------- Gerenciamento de Filas ----------------------------
    
    def enqueue_open(self, path: str):
        """Enfileira abertura de arquivo."""
        self.open_queue.put(path)

    def enqueue_switch_service_log(self, service: str, path: str):
        """Enfileira troca de log de serviço."""
        self.switch_queue.put((service, path))

    def _consume_queues(self):
        """Consome as filas de eventos."""
        # Processa trocas de serviço
        while True:
            try:
                svc, path = self.switch_queue.get_nowait()
            except queue.Empty:
                break
            self._switch_log_for_service(svc, path)

        # Processa aberturas de arquivo
        while True:
            try:
                p = self.open_queue.get_nowait()
            except queue.Empty:
                break
            self._open_log_enforcing_one_per_service(p)

        self.root.after(200, self._consume_queues)

    # ---------------------------- Gerenciamento de Abas ----------------------------
    
    def _switch_log_for_service(self, service: str, new_path: str):
        """Troca o log de um serviço."""
        old_path = self.tab_by_service.get(service)
        if old_path and old_path != new_path:
            self._close_log(old_path)
        self._open_log_enforcing_one_per_service(new_path)

    def _open_log_enforcing_one_per_service(self, filepath: str):
        """Abre log garantindo um por serviço."""
        if not os.path.isfile(filepath):
            return
        
        svc = service_from_path(filepath)
        existing_path = self.tab_by_service.get(svc)
        
        if existing_path and existing_path != filepath:
            self._close_log(existing_path)
        
        if filepath in self.open_tabs:
            self.notebook.select(self.open_tabs[filepath].frame)
            self.tab_by_service[svc] = filepath
            return
        
        tab = LogTab(self, filepath)
        self.open_tabs[filepath] = tab
        self.tab_by_service[svc] = filepath
        
        label = f"{svc} — {os.path.basename(filepath)}"
        self.notebook.add(tab.frame, text=label)
        self.notebook.select(tab.frame)

    def _close_log(self, filepath: str):
        """Fecha um log."""
        tab = self.open_tabs.pop(filepath, None)
        if not tab:
            return
        
        tab.stop_event.set()
        try:
            self.notebook.forget(tab.frame)
        except Exception:
            pass
        
        svc = service_from_path(filepath)
        if self.tab_by_service.get(svc) == filepath:
            del self.tab_by_service[svc]

    def _on_close(self):
        """Handler de fechamento da aplicação."""
        self._stop_watcher()
        for t in list(self.open_tabs.values()):
            t.stop_event.set()
        self.root.destroy()

    def run(self):
        """Inicia a aplicação."""
        self.root.mainloop()


# ============================== PONTO DE ENTRADA ==============================

if __name__ == "__main__":
    # Verifica admin antes de iniciar
    if os.name == "nt" and not is_admin():
        resposta = messagebox.askyesno(
            "Privilégios de Administrador",
            "Este programa precisa de privilégios de administrador para:\n"
            "• Reiniciar serviços Windows (net stop/start)\n"
            "• Finalizar processos (taskkill /F)\n\n"
            "Deseja reiniciar como Administrador agora?"
        )
        
        if resposta:
            run_as_admin()
        else:
            # Cria uma instância do App e modifica o título
            app = App()
            app.root.title(f"LogFácil v{__version__} - Monitor de Serviços [MODO RESTRITO - Sem Admin]")
            app.run()
    else:
        App().run()
