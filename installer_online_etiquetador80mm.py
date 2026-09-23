import ctypes
import datetime as dt
import hashlib
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import tkinter as tk
from tkinter import messagebox, ttk

import gdown


APP_NAME = "Etiquetador80mm"
WINDOW_TITLE = "Instalador en linea de Etiquetador80mm"
DRIVE_FOLDER_URL = (
    "https://drive.google.com/drive/folders/"
    "1Hark9w4ef-rrdpkd3VtoXtMCl8V6cYmO?usp=sharing"
)
INSTALLER_FILENAME = "Instalador_Etiquetador80mm.exe"
SILENT_ARGS = {"--silent", "/silent", "-silent", "--quiet", "/quiet"}
PASSIVE_ARGS = {
    "--passive",
    "/passive",
    "-passive",
    "--pasive",
    "/pasive",
    "-pasive",
}
SELF_TEST_ARGS = {"--self-test", "/self-test", "-self-test"}
DOWNLOAD_ATTEMPTS = 3
MINIMUM_INSTALLER_SIZE = 1_000_000


def _log_path() -> Path:
    candidates = [
        Path(os.environ.get("ProgramData", r"C:\ProgramData"))
        / APP_NAME
        / "Logs"
        / "instalador_online.log",
        Path(tempfile.gettempdir()) / APP_NAME / "instalador_online.log",
    ]
    for candidate in candidates:
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            with candidate.open("a", encoding="utf-8"):
                pass
            return candidate
        except OSError:
            continue
    raise OSError("No se pudo crear el registro del instalador en linea.")


LOG_PATH = _log_path()
LOG_LOCK = threading.Lock()


def write_log(message: str) -> None:
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} | {message}"
    with LOG_LOCK:
        with LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    try:
        print(line, flush=True)
    except Exception:
        pass


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> None:
    parameters = subprocess.list2cmdline(sys.argv[1:])
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        sys.executable,
        parameters,
        None,
        1,
    )
    if result <= 32:
        raise RuntimeError("No se concedieron permisos de administrador.")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _drive_files():
    last_error = None
    for attempt in range(1, DOWNLOAD_ATTEMPTS + 1):
        try:
            write_log(
                f"Consultando carpeta de Google Drive "
                f"(intento {attempt}/{DOWNLOAD_ATTEMPTS})."
            )
            files = gdown.download_folder(
                url=DRIVE_FOLDER_URL,
                skip_download=True,
                quiet=True,
                use_cookies=False,
            )
            write_log(f"Carpeta consultada; archivos encontrados: {len(files)}.")
            return files
        except Exception as exc:
            last_error = exc
            write_log(f"No se pudo consultar Google Drive: {exc}")
            if attempt < DOWNLOAD_ATTEMPTS:
                time.sleep(attempt * 1.5)
    raise RuntimeError(
        "No se pudo consultar la carpeta de Google Drive despues de "
        f"{DOWNLOAD_ATTEMPTS} intentos: {last_error}"
    )


def find_remote_installer():
    matches = [
        remote
        for remote in _drive_files()
        if Path(remote.path).name.casefold() == INSTALLER_FILENAME.casefold()
    ]
    if not matches:
        raise FileNotFoundError(
            f"No se encontro {INSTALLER_FILENAME} en la carpeta publicada."
        )
    if len(matches) > 1:
        paths = ", ".join(remote.path for remote in matches)
        raise RuntimeError(
            f"Se encontraron varias copias de {INSTALLER_FILENAME}: {paths}"
        )
    remote = matches[0]
    write_log(f"Instalador localizado: {remote.path}; id={remote.id}.")
    return remote


def download_installer(destination_dir: Path, status_callback=None) -> Path:
    if status_callback:
        status_callback("Localizando la version publicada en Google Drive...")
    remote = find_remote_installer()
    destination = destination_dir / INSTALLER_FILENAME
    last_error = None

    for attempt in range(1, DOWNLOAD_ATTEMPTS + 1):
        try:
            if destination.exists():
                destination.unlink()
            if status_callback:
                status_callback(
                    f"Descargando el instalador (intento {attempt} de "
                    f"{DOWNLOAD_ATTEMPTS})..."
                )
            write_log(
                f"Iniciando descarga; id={remote.id}; destino={destination}; "
                f"intento={attempt}."
            )
            result = gdown.download(
                id=remote.id,
                output=str(destination),
                quiet=True,
                use_cookies=False,
                resume=False,
            )
            if not result or not destination.is_file():
                raise RuntimeError("Google Drive no devolvio el archivo solicitado.")

            size = destination.stat().st_size
            if size < MINIMUM_INSTALLER_SIZE:
                raise RuntimeError(
                    f"La descarga esta incompleta: {size:,} bytes recibidos."
                )
            with destination.open("rb") as handle:
                if handle.read(2) != b"MZ":
                    raise RuntimeError("El archivo descargado no es un ejecutable de Windows.")

            digest = sha256_file(destination)
            write_log(
                f"Descarga validada; bytes={size}; sha256={digest}; "
                f"archivo={destination}."
            )
            return destination
        except Exception as exc:
            last_error = exc
            write_log(f"Fallo de descarga: {exc}")
            if attempt < DOWNLOAD_ATTEMPTS:
                time.sleep(attempt * 1.5)

    raise RuntimeError(
        "No se pudo descargar y validar el instalador despues de "
        f"{DOWNLOAD_ATTEMPTS} intentos: {last_error}"
    )


def run_downloaded_installer(installer_path: Path, status_callback=None) -> None:
    if status_callback:
        status_callback("Instalando Etiquetador80mm...")
    # El instalador en linea mantiene su propia ventana de progreso. El
    # instalador descargado se ejecuta silenciosamente para evitar dos ventanas.
    command = [str(installer_path), "/silent"]
    write_log(f"Ejecutando instalador descargado: {command!r}")
    completed = subprocess.run(
        command,
        cwd=str(installer_path.parent),
        check=False,
        capture_output=True,
        text=True,
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    output = (completed.stdout or "").strip()
    error_output = (completed.stderr or "").strip()
    if output:
        write_log(f"Salida del instalador: {output[-4000:]}")
    if error_output:
        write_log(f"Salida de error del instalador: {error_output[-4000:]}")
    write_log(f"El instalador finalizo con codigo {completed.returncode}.")
    if completed.returncode != 0:
        raise RuntimeError(
            "El instalador descargado no pudo completar la instalacion; "
            f"codigo de salida {completed.returncode}."
        )


def perform_online_install(status_callback=None) -> None:
    write_log("=" * 72)
    write_log(
        f"Inicio de instalacion en linea; version_python={sys.version.split()[0]}; "
        f"ejecutable={sys.executable}."
    )
    with tempfile.TemporaryDirectory(prefix="Etiquetador80mm_online_") as temp_dir:
        installer_path = download_installer(Path(temp_dir), status_callback)
        run_downloaded_installer(installer_path, status_callback)
    write_log("Instalacion en linea completada correctamente.")


def self_test() -> None:
    write_log("=" * 72)
    write_log("Iniciando autodiagnostico del instalador en linea.")
    remote = find_remote_installer()
    write_log(
        f"AUTODIAGNOSTICO CORRECTO; archivo={remote.path}; id={remote.id}; "
        f"log={LOG_PATH}."
    )


class OnlineInstallerWindow(tk.Tk):
    def __init__(self, passive: bool = False):
        super().__init__()
        self.passive = passive
        self.exit_code = 0
        self.installing = False
        self.title(WINDOW_TITLE)
        self.geometry("620x300")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._close_requested)

        frame = ttk.Frame(self, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Etiquetador 80 mm",
            font=("Segoe UI", 21, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text="Instalacion en linea desde Google Drive",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(4, 18))
        ttk.Label(
            frame,
            text=(
                "Se descargara la version publicada de "
                f"{INSTALLER_FILENAME} y se instalara automaticamente."
            ),
            wraplength=555,
            justify="left",
        ).pack(anchor="w")
        self.status = ttk.Label(frame, text="Listo para descargar e instalar.")
        self.status.pack(anchor="w", pady=(18, 10))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x")

        self.buttons = ttk.Frame(frame)
        self.buttons.pack(fill="x", pady=(18, 0))
        self.cancel_button = ttk.Button(
            self.buttons,
            text="Cancelar",
            command=self._close_requested,
        )
        self.cancel_button.pack(side="right")
        self.install_button = ttk.Button(
            self.buttons,
            text="Descargar e instalar",
            command=self.start_install,
        )
        self.install_button.pack(side="right", padx=(0, 10))

        if self.passive:
            self.buttons.pack_forget()
            self.after(200, self.start_install)

    def _close_requested(self):
        if not self.installing:
            self.destroy()

    def _set_status(self, value: str) -> None:
        self.after(0, lambda: self.status.configure(text=value))

    def start_install(self):
        if self.installing:
            return
        self.installing = True
        self.install_button.configure(state="disabled")
        self.cancel_button.configure(state="disabled")
        self.progress.start(12)
        self._set_status("Preparando la descarga...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            perform_online_install(self._set_status)
        except Exception as exc:
            write_log(f"ERROR: {exc}")
            write_log(traceback.format_exc())
            self.exit_code = 1
            self.after(0, lambda error=exc: self._finish_error(error))
            return
        self.after(0, self._finish_success)

    def _finish_success(self):
        self.installing = False
        self.progress.stop()
        self.progress.configure(mode="determinate", value=100)
        self.status.configure(text=f"Instalacion completada. Registro: {LOG_PATH}")
        if self.passive:
            self.after(1600, self.destroy)
        else:
            messagebox.showinfo(
                WINDOW_TITLE,
                f"Instalacion completada correctamente.\n\nRegistro:\n{LOG_PATH}",
                parent=self,
            )
            self.destroy()

    def _finish_error(self, exc: Exception):
        self.installing = False
        self.progress.stop()
        self.status.configure(text=f"No se pudo completar. Registro: {LOG_PATH}")
        if self.passive:
            self.after(5000, self.destroy)
        else:
            messagebox.showerror(
                WINDOW_TITLE,
                f"No se pudo completar la instalacion:\n\n{exc}\n\nRegistro:\n{LOG_PATH}",
                parent=self,
            )
            self.install_button.configure(state="normal")
            self.cancel_button.configure(state="normal")


def main() -> int:
    args = {argument.lower() for argument in sys.argv[1:]}
    silent = bool(args & SILENT_ARGS)
    passive = bool(args & PASSIVE_ARGS)
    self_test_requested = bool(args & SELF_TEST_ARGS)

    if self_test_requested:
        self_test()
        return 0

    if not is_admin():
        write_log("Solicitando permisos de administrador.")
        relaunch_as_admin()
        return 0

    if silent:
        perform_online_install()
        return 0

    window = OnlineInstallerWindow(passive=passive)
    window.mainloop()
    return window.exit_code


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        try:
            write_log(f"ERROR FATAL: {exc}")
            write_log(traceback.format_exc())
        except Exception:
            pass
        normalized_args = {argument.lower() for argument in sys.argv[1:]}
        if not normalized_args & (SILENT_ARGS | PASSIVE_ARGS | SELF_TEST_ARGS):
            try:
                messagebox.showerror(
                    WINDOW_TITLE,
                    f"Error del instalador:\n\n{exc}\n\nRegistro:\n{LOG_PATH}",
                )
            except Exception:
                pass
        raise SystemExit(1)
