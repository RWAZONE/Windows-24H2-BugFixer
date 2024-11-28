import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import subprocess
import sys
import os
import ctypes
import threading
import time
import locale
import traceback
import queue
import logging

mutex_handle = None

def resource_path(relative_path):
    """
    PyInstaller ile paketlenmiş uygulamada kaynak dosyasının yolunu bulur.
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def set_locale():
    """
    Sistem varsayılan locale ayarını kullanır.
    """
    try:
        locale.setlocale(locale.LC_ALL, '')
        logging.info(f"Locale başarıyla ayarlandı: {locale.getlocale()}")
    except locale.Error as e:
        logging.warning(f"Locale ayarı başarısız oldu: {e}. Sistem varsayılan locale kullanılıyor.")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("RWAZONE - Windows Sorun Giderici Aracı")
        self.root.geometry("900x800")
        self.root.resizable(False, False)

        self.set_icon()

        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.configure_style()

        self.load_logo()

        self.title_label = ttk.Label(
            self.main_frame,
            text="Windows Sorun Giderici Aracı",
            font=("Segoe UI", 24, "bold"),
            foreground="#FF6666",
            background="#2e2e2e"
        )
        self.title_label.pack(pady=(20, 10))

        description = (
            "Windows 24H2 (KB:5046740) Sanallaştırma güncellemesiyle birlikte oyunlara girişte problem yaşandığından dolayı bu güncellemeyi kaldırarak oyunlara giriş yapabilirsiniz. "
            "Doğrudan ilgili güncellemeyi kaldırarak hızlı bir çözüm elde edilebilir. "
            "Windows tarafından yayınlanacak yeni güncellemelerde bu sorun çözülecektir. Bu esnada hatayla karşılaşmamak için bu araç üzerinden ilgili güncellemeyi kaldırabilirsiniz. Güncellemenin siz istemediğiniz sürece yüklenmemesi için Windows güncellemelerini durdurabilirsiniz. "
        )

        self.description_label = ttk.Label(
            self.main_frame,
            text=description,
            wraplength=860,
            justify=tk.LEFT,
            background="#2e2e2e",
            foreground="#CCCCCC",
            font=("Segoe UI", 12)
        )
        self.description_label.pack(pady=(0, 30))

        self.buttons_frame = ttk.Frame(self.main_frame)
        self.buttons_frame.pack(pady=(0, 20))

        self.progress_var = tk.IntVar()
        self.progress_canvas = tk.Canvas(
            self.main_frame,
            width=860,
            height=30,
            bg="#3e3e3e",
            highlightthickness=0
        )
        self.progress_canvas.pack(pady=(0, 20))

        self.progress_canvas.create_rectangle(
            0, 0, 860, 30,
            fill="#3e3e3e",
            outline=""
        )

        self.progress_bar = self.progress_canvas.create_rectangle(
            0, 0, 0, 30,
            fill="#FF6666",
            outline=""
        )

        self.log_frame = ttk.Frame(self.main_frame)
        self.log_frame.pack(pady=(0, 10), fill=tk.BOTH, expand=True)

        self.log_scrollbar = ttk.Scrollbar(self.log_frame, orient=tk.VERTICAL)
        self.log_text = tk.Text(
            self.log_frame,
            height=15,
            state='disabled',
            wrap='word',
            bg="#1e1e1e",
            fg="#FFFFFF",
            font=("Segoe UI", 10),
            yscrollcommand=self.log_scrollbar.set
        )
        self.log_scrollbar.config(command=self.log_text.yview)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(self.main_frame, text="", foreground="#FF6666", background="#2e2e2e", font=("Segoe UI", 12))
        self.status_label.pack(pady=(0, 10))

        self.uninstall_button = ttk.Button(
            self.buttons_frame,
            text="KB:5046740 Kaldır",
            command=self.uninstall_kb,
            style='Modern.TButton'
        )
        self.uninstall_button.grid(row=0, column=0, padx=15, pady=10, sticky='ew')

        self.stop_update_button = ttk.Button(
            self.buttons_frame,
            text="Windows Update'i Kapat",
            command=self.stop_windows_update,
            style='Modern.TButton'
        )
        self.stop_update_button.grid(row=0, column=1, padx=15, pady=10, sticky='ew')

        self.start_update_button = ttk.Button(
            self.buttons_frame,
            text="Windows Update'i Aç",
            command=self.start_windows_update,
            style='Modern.TButton'
        )
        self.start_update_button.grid(row=0, column=2, padx=15, pady=10, sticky='ew')

        self.buttons_frame.columnconfigure(0, weight=1)
        self.buttons_frame.columnconfigure(1, weight=1)
        self.buttons_frame.columnconfigure(2, weight=1)

        self.footer = ttk.Label(
            self.main_frame,
            text="rwazone.com",
            font=("Segoe UI", 10),
            background="#2e2e2e",
            foreground="#666666"
        )
        self.footer.pack(side=tk.BOTTOM, pady=(20, 0))

        self.queue = queue.Queue()
        self.root.after(100, self.process_queue)

    def set_icon(self):
        try:
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                raise FileNotFoundError(f"İkon dosyası bulunamadı: {icon_path}")
        except Exception as e:
            logging.error(f"İkon yüklenemedi: {e}")

    def configure_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', background="#2e2e2e", foreground="#ffffff")
        style.configure('TLabel', background="#2e2e2e", foreground="#ffffff")
        
        style.configure('Modern.TButton',
                        background="#444444",
                        foreground="#FFFFFF",
                        borderwidth=0,
                        focusthickness=3,
                        focuscolor='none',
                        font=("Segoe UI", 12, "bold"))
        style.map('Modern.TButton',
                  background=[('active', '#555555')],
                  foreground=[('active', '#FFFFFF')])

    def load_logo(self):
        try:
            logo_path = resource_path("logo.png")
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                try:
                    resample = Image.Resampling.LANCZOS
                except AttributeError:
                    resample = Image.ANTIALIAS
                logo_image = logo_image.resize((197, 66), resample)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                self.logo_label = ttk.Label(self.main_frame, image=self.logo_photo, background="#2e2e2e")
                self.logo_label.image = self.logo_photo
                self.logo_label.pack(pady=(0, 20))
            else:
                raise FileNotFoundError(f"Logo dosyası bulunamadı: {logo_path}")
        except Exception as e:
            logging.error(f"Logo yüklenemedi: {e}")

    def process_queue(self):
        try:
            while True:
                message = self.queue.get_nowait()
                if message['type'] == 'log':
                    self.append_log(message['content'])
                elif message['type'] == 'status':
                    self.update_status(message['content'], message.get('color', '#FF6666'))
                elif message['type'] == 'progress':
                    self.update_progress(message['value'])
                elif message['type'] == 'done':
                    self.enable_buttons()
                elif message['type'] == 'error':
                    self.append_log(message['content'])
                    self.update_status("Beklenmeyen bir hata oluştu.", "#FF6666")
                    messagebox.showerror("Hata", message['content'])
                    self.enable_buttons()
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    def append_log(self, text):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def update_status(self, text, color="#FF6666"):
        self.status_label.config(text=text, foreground=color)

    def disable_buttons(self):
        self.uninstall_button.config(state='disabled')
        self.stop_update_button.config(state='disabled')
        self.start_update_button.config(state='disabled')

    def enable_buttons(self):
        self.uninstall_button.config(state='normal')
        self.stop_update_button.config(state='normal')
        self.start_update_button.config(state='normal')

    def update_progress(self, value):
        """
        Canvas tabanlı progress bar'ı günceller.
        """
        self.progress_var.set(value)

        new_width = (value / 100) * 860
        self.progress_canvas.coords(self.progress_bar, 0, 0, new_width, 30)

    def run_command(self, cmd, encoding='cp1254'):
        """
        Verilen komutu arka planda çalıştırır, çıktısını toplar ve GUI'de gösterir.
        """
        self.disable_buttons()
        self.update_status("İşlem başlatılıyor...", "#CCCCCC") 
        self.append_log("İşlem başlatılıyor...")
        
        def task():
            try:
                full_cmd = f"chcp 1254 >nul && {cmd}"
                process = subprocess.Popen(
                    full_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    shell=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    text=True, 
                    encoding=encoding,
                    errors='replace'
                )

                while True:
                    output = process.stdout.readline()
                    if output:
                        self.queue.put({'type': 'log', 'content': output.strip()})
                        current = self.progress_var.get()
                        if current < 90:
                            self.queue.put({'type': 'progress', 'value': current + 1})
                    elif process.poll() is not None:
                        break
                    time.sleep(0.1)

                stdout, stderr = process.communicate()

                if process.returncode == 0:
                    self.queue.put({'type': 'progress', 'value': 100})
                    self.queue.put({'type': 'status', 'content': "İşlem başarılı.", 'color': "#00FF00"}) 
                    self.queue.put({'type': 'log', 'content': "İşlem başarıyla tamamlandı."})
                    messagebox.showinfo("Başarılı", "İşlem başarıyla tamamlandı.")
                else:
                    stderr_lower = stderr.lower()
                    if ("service is already running" in stderr_lower or
                        "the requested service has already been started" in stderr_lower or
                        "hizmet zaten çalışıyor" in stderr_lower):
                        self.queue.put({'type': 'progress', 'value': 100})
                        self.queue.put({'type': 'status', 'content': "Hizmet zaten çalışıyor.", 'color': "#CCCCCC"})
                        self.queue.put({'type': 'log', 'content': "Hizmet zaten çalışıyor."})
                        messagebox.showinfo("Bilgi", "Hizmet zaten çalışıyor.")
                    else:
                        self.queue.put({'type': 'progress', 'value': 100})
                        self.queue.put({'type': 'status', 'content': "İşlem başarısız oldu.", 'color': "#FF6666"})
                        error_message = stderr.strip() if stderr.strip() else "Bilinmeyen bir hata oluştu."
                        self.queue.put({'type': 'log', 'content': error_message})
                        messagebox.showerror("Hata", f"İşlem başarısız oldu:\n{error_message}")
            except Exception as e:
                error_details = f"Error: {traceback.format_exc()}"
                logging.error(error_details)
                self.queue.put({'type': 'error', 'content': f"Beklenmeyen bir hata oluştu:\n{e}"})
            finally:
                self.queue.put({'type': 'done'})

        threading.Thread(target=task, daemon=True).start()

    def check_service_status(self, service_name):
        """
        Belirli bir Windows hizmetinin durumunu kontrol eder.
        Döner: 'RUNNING', 'STOPPED', 'UNKNOWN'
        """
        try:
            cmd = f'chcp 1254 >nul && sc query "{service_name}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='cp1254', errors='replace')
            if result.returncode != 0:
                return "UNKNOWN"
            for line in result.stdout.splitlines():
                if "STATE" in line.upper():
                    if "RUNNING" in line.upper():
                        return "RUNNING"
                    elif "STOPPED" in line.upper():
                        return "STOPPED"
            return "UNKNOWN"
        except Exception as e:
            logging.error(f"Service status check failed for {service_name}: {e}")
            return "UNKNOWN"

    def uninstall_kb(self):
        if not self.check_kb_installed("5046740"):
            messagebox.showinfo("Bilgi", "Belirtilen KB güncellemesi zaten kaldırılmış.")
            return
        cmd = "wusa /uninstall /kb:5046740 /quiet /norestart"
        self.run_command(cmd)

    def stop_windows_update(self):
        wuauserv_status = self.check_service_status("wuauserv")
        bits_status = self.check_service_status("bits")
        if wuauserv_status == "STOPPED" and bits_status == "STOPPED":
            messagebox.showinfo("Bilgi", "Windows Update hizmetleri zaten durdurulmuş.")
            return
        cmd = "net stop wuauserv && net stop bits"
        self.run_command(cmd)

    def start_windows_update(self):
        wuauserv_status = self.check_service_status("wuauserv")
        bits_status = self.check_service_status("bits")
        if wuauserv_status == "RUNNING" and bits_status == "RUNNING":
            messagebox.showinfo("Bilgi", "Windows Update hizmetleri zaten çalışıyor.")
            return
        cmd = "net start wuauserv && net start bits"
        self.run_command(cmd)

    def check_kb_installed(self, kb_number):
        """
        Belirli bir KB güncellemesinin yüklü olup olmadığını kontrol eder.
        """
        try:
            cmd = f'chcp 1254 >nul && wmic qfe get HotFixID'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='utf-16le', errors='replace')
            hotfixes = [line.strip().upper() for line in result.stdout.splitlines() if line.strip() and line.strip().upper().startswith('KB')]
            return kb_number.upper() in hotfixes
        except Exception as e:
            logging.error(f"KB kontrolü başarısız: {e}")
            return False

def is_admin():
    """
    Uygulamanın yönetici yetkisine sahip olup olmadığını kontrol eder.
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """
    Uygulamayı yönetici yetkisiyle yeniden başlatır.
    """
    try:
        script = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
        params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}" {params}', None, 1
        )
    except Exception as e:
        messagebox.showerror("Yönetici Yetkisi Gerekli", f"Uygulama yönetici yetkisi ile yeniden başlatılamadı:\n{e}")
        sys.exit()

def check_single_instance():
    """
    Aynı anda birden fazla uygulama örneğinin çalışmasını engeller.
    """
    global mutex_handle
    mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False, "WindowsSorunGidericiAraciMutex")
    last_error = ctypes.windll.kernel32.GetLastError()
    if last_error == 183:
        messagebox.showwarning("Uyarı", "Uygulama zaten çalışıyor.")
        sys.exit()

def setup_logging():
    """
    Uygulama için logging yapılandırması.
    """
    log_path = resource_path("app_log.txt")
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("Uygulama başlatılıyor...")


