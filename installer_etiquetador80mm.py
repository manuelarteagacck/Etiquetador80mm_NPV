import ctypes
import datetime as dt
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import traceback
import winreg


APP_NAME = "Etiquetador80mm"
TARGET_DIR = Path(r"D:\Etiquetador80mm")
INSTALLER_FILENAME = "Instalador_Etiquetador80mm.exe"
PAYLOAD_DIRNAME = "payload"
TASK_NAME = "NPV Backup Precios Venta"
TASKBAR_LAYOUT_FILENAME = "TaskbarLayoutModification.xml"
TASKBAR_POLICY_KEY = r"SOFTWARE\Policies\Microsoft\Windows\Explorer"
CONFIG_FILES = {"config.json", "config.key", "paramconf.json"}
SILENT_ARGS = {"--silent", "/silent", "-silent", "--quiet", "/quiet"}
PASSIVE_ARGS = {"--passive", "/passive", "-passive", "--pasive", "/pasive", "-pasive"}
UNINSTALL_ARGS = {"--uninstall", "/uninstall", "-uninstall"}
PAYLOAD_FILES = {
    "Etiquetador80mm.exe": "Etiquetador80mm.exe",
    "BackupPreciosVenta.exe": "BackupPreciosVenta.exe",
    "config.json": "config.json",
    "config.key": "config.key",
    "paramconf.json": "paramconf.json",
    "ticket_printer.ico": "ticket_printer.ico",
    "Recorrido_Etiquetador80mm_CirculoK_v7_actualizado.mp4": "Video_Funcionamiento_Etiquetador80mm.mp4",
}


def bundle_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
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


def create_shortcut(shortcut: Path, target: Path, arguments: str = "", description: str = ""):
    shortcut.parent.mkdir(parents=True, exist_ok=True)
    def ps_literal(value) -> str:
        return "'" + str(value).replace("'", "''") + "'"

    powershell = f"""
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut({ps_literal(shortcut)})
$link.TargetPath = {ps_literal(target)}
$link.Arguments = {ps_literal(arguments)}
$link.WorkingDirectory = {ps_literal(target.parent)}
$link.IconLocation = {ps_literal(f'{target},0')}
$link.Description = {ps_literal(description)}
$link.Save()
"""
    subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            powershell,
        ],
        check=True,
        capture_output=True,
        text=True,
    )


def start_menu_shortcuts(app_exe: Path, video_path: Path, installer_exe: Path):
    program_data = Path(os.environ.get("ProgramData", r"C:\ProgramData"))
    programs_dir = program_data / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    menu_dir = programs_dir / APP_NAME
    common_shortcut = programs_dir / f"{APP_NAME}.lnk"
    create_shortcut(common_shortcut, app_exe, description="Etiquetador térmico de 80 mm")
    create_shortcut(menu_dir / f"{APP_NAME}.lnk", app_exe, description="Etiquetador térmico de 80 mm")
    create_shortcut(menu_dir / "Video de funcionamiento.lnk", video_path, description="Recorrido funcional del Etiquetador80mm")
    create_shortcut(
        menu_dir / f"Desinstalar {APP_NAME}.lnk",
        installer_exe,
        arguments="--uninstall",
        description=f"Desinstalar {APP_NAME}",
    )
    return menu_dir, common_shortcut


def configure_taskbar_policy(common_shortcut: Path, app_exe: Path):
    layout_path = TARGET_DIR / TASKBAR_LAYOUT_FILENAME
    common_link = r"%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Etiquetador80mm.lnk"
    layout_xml = f'''<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <LayoutOptions StartTileGroupCellWidth="6" />
  <DefaultLayoutOverride LayoutCustomizationRestrictionType="OnlySpecifiedGroups">
    <StartLayoutCollection>
      <defaultlayout:StartLayout GroupCellWidth="6">
        <start:Group Name="Etiquetador 80 mm">
          <start:DesktopApplicationTile
              Size="2x2"
              Column="0"
              Row="0"
              DesktopApplicationLinkPath="{common_link}" />
        </start:Group>
      </defaultlayout:StartLayout>
    </StartLayoutCollection>
  </DefaultLayoutOverride>
  <CustomTaskbarLayoutCollection>
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationLinkPath="{common_link}" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
'''
    layout_path.write_text(layout_xml, encoding="utf-8")

    with winreg.CreateKeyEx(
        winreg.HKEY_LOCAL_MACHINE,
        TASKBAR_POLICY_KEY,
        0,
        winreg.KEY_READ | winreg.KEY_WRITE,
    ) as policy_key:
        try:
            existing_path, _ = winreg.QueryValueEx(policy_key, "StartLayoutFile")
        except FileNotFoundError:
            existing_path = ""
        if existing_path and Path(os.path.expandvars(existing_path)).resolve() != layout_path.resolve():
            raise RuntimeError(
                "El equipo ya tiene una directiva de diseño de Inicio/barra de tareas. "
                f"No se reemplazó: {existing_path}"
            )
        winreg.SetValueEx(policy_key, "StartLayoutFile", 0, winreg.REG_EXPAND_SZ, str(layout_path))
        winreg.SetValueEx(policy_key, "LockedStartLayout", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(policy_key, "ReapplyStartLayoutEveryLogon", 0, winreg.REG_DWORD, 0)

    # Conserva una copia por perfil para Windows anteriores. En Windows 11 el
    # anclaje visible lo realiza la directiva XML, no la mera copia del .lnk.
    pinned, failures = taskbar_shortcuts(app_exe)
    immediate_pin = pin_current_user_windows10(app_exe)
    try:
        result = ctypes.c_ulong()
        ctypes.windll.user32.SendMessageTimeoutW(
            0xFFFF,
            0x001A,
            0,
            "Policy",
            0x0002,
            5000,
            ctypes.byref(result),
        )
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
    except Exception:
        pass
    return layout_path, pinned, failures, immediate_pin


def windows_build_number() -> int:
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
        ) as version_key:
            value, _ = winreg.QueryValueEx(version_key, "CurrentBuildNumber")
        return int(value)
    except (FileNotFoundError, OSError, TypeError, ValueError):
        return 0


def pin_current_user_windows10(app_exe: Path) -> str:
    build = windows_build_number()
    if build >= 22000:
        return f"Omitido: compilación {build} corresponde a Windows 11"
    ps_path = str(app_exe.parent).replace("'", "''")
    ps_name = app_exe.name.replace("'", "''")
    powershell = f"""
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace('{ps_path}')
$item = $folder.ParseName('{ps_name}')
$item.InvokeVerb('taskbarpin')
Start-Sleep -Milliseconds 800
"""
    try:
        subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                powershell,
            ],
            check=True,
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return "Solicitado mediante el verbo taskbarpin de Windows 10"
    except Exception as exc:
        return f"No disponible; se aplicará al iniciar sesión mediante directiva: {exc}"


def taskbar_shortcuts(shortcut_target: Path):
    users_root = Path(os.environ.get("SystemDrive", "C:")) / "Users"
    profile_dirs = []
    if users_root.exists():
        for profile in users_root.iterdir():
            if not profile.is_dir() or profile.name.lower() in {"all users", "default user", "public"}:
                continue
            profile_dirs.append(profile)
    default_profile = users_root / "Default"
    if default_profile.exists() and default_profile not in profile_dirs:
        profile_dirs.append(default_profile)

    created = []
    failures = []
    for profile in profile_dirs:
        pinned = profile / "AppData" / "Roaming" / "Microsoft" / "Internet Explorer" / "Quick Launch" / "User Pinned" / "TaskBar"
        try:
            shortcut = pinned / f"{APP_NAME}.lnk"
            create_shortcut(shortcut, shortcut_target, description="Etiquetador térmico de 80 mm")
            created.append(shortcut)
        except Exception as exc:
            failures.append((profile.name, str(exc)))
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
    except Exception:
        pass
    return created, failures


def register_backup_task(backup_exe: Path):
    task_command = f'"{backup_exe}" --apply'
    subprocess.run(
        [
            "schtasks.exe",
            "/Create",
            "/TN",
            TASK_NAME,
            "/TR",
            task_command,
            "/SC",
            "DAILY",
            "/ST",
            "21:00",
            "/RU",
            "SYSTEM",
            "/RL",
            "HIGHEST",
            "/F",
        ],
        check=True,
        capture_output=True,
        text=True,
    )


def write_install_log(lines):
    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    log = TARGET_DIR / "instalacion.log"
    with log.open("a", encoding="utf-8") as handle:
        handle.write("=" * 80 + "\n")
        handle.write(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        handle.write("\n".join(lines) + "\n")
    return log


def install():
    payload = bundle_path(PAYLOAD_DIRNAME)
    if not payload.exists():
        raise FileNotFoundError(f"No se encontró el contenido del instalador: {payload}")
    TARGET_DIR.mkdir(parents=True, exist_ok=True)

    messages = []
    for source_name, destination_name in PAYLOAD_FILES.items():
        source = payload / source_name
        destination = TARGET_DIR / destination_name
        if not source.exists():
            raise FileNotFoundError(source)
        if destination_name in CONFIG_FILES and destination.exists():
            messages.append(f"Configuración conservada: {destination}")
            continue
        shutil.copy2(source, destination)
        messages.append(f"Instalado: {destination}")

    app_exe = TARGET_DIR / "Etiquetador80mm.exe"
    backup_exe = TARGET_DIR / "BackupPreciosVenta.exe"
    video_path = TARGET_DIR / "Video_Funcionamiento_Etiquetador80mm.mp4"
    installer_source = Path(sys.executable).resolve()
    installer_exe = TARGET_DIR / INSTALLER_FILENAME
    if installer_source != installer_exe.resolve():
        shutil.copy2(installer_source, installer_exe)
        messages.append(f"Instalador conservado: {installer_exe}")
    menu_dir, common_shortcut = start_menu_shortcuts(app_exe, video_path, installer_exe)
    layout_path, pinned, pin_failures, immediate_pin = configure_taskbar_policy(common_shortcut, app_exe)
    if pin_failures:
        detail = "; ".join(f"{profile}: {error}" for profile, error in pin_failures)
        raise RuntimeError(f"No se pudieron crear todos los accesos de barra de tareas: {detail}")
    register_backup_task(backup_exe)

    messages.append(f"Menú Inicio común: {menu_dir}")
    messages.append(f"Acceso principal del menú Inicio: {common_shortcut}")
    messages.append(f"Diseño de barra de tareas para todos los usuarios: {layout_path}")
    messages.append(f"Anclaje inmediato del usuario instalador: {immediate_pin}")
    pinned_profiles = sorted({shortcut.parts[2] for shortcut in pinned if len(shortcut.parts) > 2})
    messages.append(
        f"Accesos de barra de tareas creados: {len(pinned)} "
        f"({', '.join(pinned_profiles)})"
    )
    messages.append(f"Tarea programada: {TASK_NAME}, diariamente a las 21:00 como SYSTEM")
    messages.append("Instalación completada correctamente.")
    log = write_install_log(messages)
    return messages, log


def uninstall():
    subprocess.run(
        ["schtasks.exe", "/Delete", "/TN", TASK_NAME, "/F"],
        check=False,
        capture_output=True,
        text=True,
    )
    program_data = Path(os.environ.get("ProgramData", r"C:\ProgramData"))
    programs_dir = program_data / "Microsoft" / "Windows" / "Start Menu" / "Programs"
    menu_dir = programs_dir / APP_NAME
    if menu_dir.exists():
        shutil.rmtree(menu_dir, ignore_errors=True)
    try:
        (programs_dir / f"{APP_NAME}.lnk").unlink()
    except FileNotFoundError:
        pass
    layout_path = TARGET_DIR / TASKBAR_LAYOUT_FILENAME
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            TASKBAR_POLICY_KEY,
            0,
            winreg.KEY_READ | winreg.KEY_WRITE,
        ) as policy_key:
            try:
                existing_path, _ = winreg.QueryValueEx(policy_key, "StartLayoutFile")
            except FileNotFoundError:
                existing_path = ""
            if existing_path and Path(os.path.expandvars(existing_path)).resolve() == layout_path.resolve():
                for value_name in ("StartLayoutFile", "LockedStartLayout", "ReapplyStartLayoutEveryLogon"):
                    try:
                        winreg.DeleteValue(policy_key, value_name)
                    except FileNotFoundError:
                        pass
    except FileNotFoundError:
        pass
    users_root = Path(os.environ.get("SystemDrive", "C:")) / "Users"
    if users_root.exists():
        for shortcut in users_root.glob("*/AppData/Roaming/Microsoft/Internet Explorer/Quick Launch/User Pinned/TaskBar/Etiquetador80mm.lnk"):
            try:
                shortcut.unlink()
            except Exception:
                pass
    for filename in PAYLOAD_FILES.values():
        if filename in CONFIG_FILES:
            continue
        try:
            (TARGET_DIR / filename).unlink()
        except FileNotFoundError:
            pass
    try:
        layout_path.unlink()
    except FileNotFoundError:
        pass
    write_install_log(["Desinstalación completada. Se conservaron config.json, config.key y paramconf.json."])


class InstallerWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Instalador de {APP_NAME}")
        self.geometry("620x330")
        self.resizable(False, False)
        try:
            self.iconbitmap(str(bundle_path(PAYLOAD_DIRNAME) / "ticket_printer.ico"))
        except Exception:
            pass

        frame = ttk.Frame(self, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Etiquetador 80 mm", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Instalación y actualización para tienda NPV", font=("Segoe UI", 11)).pack(anchor="w", pady=(4, 22))
        ttk.Label(frame, text=f"Destino: {TARGET_DIR}", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text=(
                "Se instalará la aplicación, el video de funcionamiento, la tarea diaria de respaldo "
                "y los accesos comunes del menú Inicio y la barra de tareas. La configuración existente se conservará."
            ),
            wraplength=555,
            justify="left",
        ).pack(anchor="w", pady=(12, 24))
        self.status = ttk.Label(frame, text="Listo para instalar.")
        self.status.pack(anchor="w", pady=(0, 16))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Instalar / Actualizar", command=self.run_install).pack(side="right", padx=(0, 10))

    def run_install(self):
        self.status.configure(text="Instalando...")
        self.update_idletasks()
        try:
            _, log = install()
            self.status.configure(text=f"Instalación completada. Registro: {log}")
            messagebox.showinfo(APP_NAME, "Instalación completada correctamente.")
        except Exception as exc:
            write_install_log(["ERROR de instalación", str(exc), traceback.format_exc()])
            self.status.configure(text="La instalación no pudo completarse.")
            messagebox.showerror(APP_NAME, f"No se pudo instalar:\n\n{exc}")


class PassiveInstallerWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.exit_code = 0
        self.title(f"Instalando {APP_NAME}")
        self.geometry("500x160")
        self.resizable(False, False)
        try:
            self.iconbitmap(str(bundle_path(PAYLOAD_DIRNAME) / "ticket_printer.ico"))
        except Exception:
            pass
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Etiquetador 80 mm", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        self.status = ttk.Label(frame, text="Preparando la instalación...")
        self.status.pack(anchor="w", pady=(8, 14))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x")
        self.progress.start(12)
        self.after(200, self.run_install)

    def run_install(self):
        self.status.configure(text="Instalando y configurando accesos para todos los usuarios...")
        self.update_idletasks()
        try:
            _, log = install()
            self.progress.stop()
            self.progress.configure(mode="determinate", value=100)
            self.status.configure(text=f"Instalación completada. Registro: {log}")
            self.after(1600, self.destroy)
        except Exception as exc:
            self.exit_code = 1
            write_install_log(["ERROR de instalación pasiva", str(exc), traceback.format_exc()])
            self.progress.stop()
            self.status.configure(text=f"No se pudo completar: {exc}")
            self.after(5000, self.destroy)


def main():
    args = {argument.lower() for argument in sys.argv[1:]}
    silent = bool(args & SILENT_ARGS)
    passive = bool(args & PASSIVE_ARGS)
    uninstall_requested = bool(args & UNINSTALL_ARGS)
    if not is_admin():
        relaunch_as_admin()
        return 0
    if uninstall_requested:
        uninstall()
        if not silent:
            messagebox.showinfo(APP_NAME, "Desinstalación completada. Se conservó la configuración.")
        return 0
    if silent:
        messages, log = install()
        print("\n".join(messages))
        print(f"LOG={log}")
        return 0
    if passive:
        window = PassiveInstallerWindow()
        window.mainloop()
        return window.exit_code
    InstallerWindow().mainloop()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        try:
            write_install_log(["ERROR", str(exc), traceback.format_exc()])
        except Exception:
            pass
        normalized_args = {argument.lower() for argument in sys.argv[1:]}
        if normalized_args & (SILENT_ARGS | PASSIVE_ARGS):
            print(f"ERROR: {exc}", file=sys.stderr)
        else:
            try:
                messagebox.showerror(APP_NAME, f"Error del instalador:\n\n{exc}")
            except Exception:
                pass
        raise
