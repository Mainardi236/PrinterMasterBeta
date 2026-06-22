import customtkinter as ctk
import win32print
import win32service
import win32serviceutil
from tkinter import filedialog
import os
import subprocess

def restart():
    service_name = "Spooler"
    try:
        win32serviceutil.RestartService(service_name)
    except Exception:
        try:
            win32serviceutil.StopService(service_name)
            win32serviceutil.WaitForServiceStatus(
                service_name,
                win32service.SERVICE_STOPPED,
                30
            )
        except Exception:
            pass
        win32serviceutil.StartService(service_name)


def _printer_info_value(info, key, index):
    if isinstance(info, dict):
        return info.get(key, "N/A")
    try:
        return info[index]
    except Exception:
        return "N/A"


def get_printer_info():
    try:
        default_printer = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(default_printer)
        info = win32print.GetPrinter(handle, 2)
        win32print.ClosePrinter(handle)

        printer_name = _printer_info_value(info, "pPrinterName", 1)
        driver_name = _printer_info_value(info, "pDriverName", 4)
        port_name = _printer_info_value(info, "pPortName", 3)

        return (
            f"Impressora predeterminada: {printer_name}\n"
            f"Controlador: {driver_name}\n"
            f"Porta: {port_name}"
        )
    except Exception as error:
        return f"Erro ao obter informação da impressora: {error}"


def list_printers():
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        printers = win32print.EnumPrinters(flags, None, 1)
        printer_names = [printer[2] for printer in printers] if printers else ["Nenhuma impressora encontrada"]
        return "Impressoras instaladas:\n" + "\n".join(printer_names)
    except Exception as error:
        return f"Erro ao listar impressoras: {error}"


def refresh_info():
    info_label.configure(text=get_printer_info())


def show_printers():
    info_label.configure(text=list_printers())


def print_folder_files():
    global selected_folder
    folder = selected_folder or filedialog.askdirectory()
    if not folder:
        return

    # Get the default printer automatically
    try:
        PRINTER_NAME = win32print.GetDefaultPrinter()
    except Exception:
        PRINTER_NAME = "NOME_DA_SUA_IMPRESSORA"

    # Build PowerShell script to print all PDFs in the folder using SumatraPDF
    ps_script = (
        "$folder = '{folder}';"
        "$printer = '{printer}';"
        "Get-ChildItem $folder -Filter *.pdf | ForEach-Object {"
        " & 'C:\\Users\\leao-\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe' -print-to $printer $_.FullName;"
        " Start-Sleep -Milliseconds 800 }"
    ).format(folder=folder.replace("'", "''"), printer=PRINTER_NAME.replace("'", "''"))

    try:
        # Execute PowerShell script
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], check=True)
        info_label.configure(text=f"Enviados para impressão todos os PDFs de:\n{folder}")
    except Exception as error:
        info_label.configure(text=f"Erro ao imprimir arquivos da pasta: {error}")

app = ctk.CTk()
app.geometry("420x280")
app.title("Printer Master")

btn_restart = ctk.CTkButton(
    app,
    text="Reiniciar Spooler",
    command=restart
)

btn_refresh = ctk.CTkButton(
    app,
    text="Atualizar informações da impressora",
    command=refresh_info
)

btn_list = ctk.CTkButton(
    app,
    text="Listar impressoras",
    command=show_printers
)

def select_folder():
    global selected_folder
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        folder_label.configure(text=f"Pasta selecionada: {folder}")

selected_folder = None

btn_select_folder = ctk.CTkButton(
    app,
    text="Selecionar pasta",
    command=select_folder
)

btn_print_folder = ctk.CTkButton(
    app,
    text="Imprimir arquivos da pasta",
    command=print_folder_files
)

info_label = ctk.CTkLabel(
    app,
    text="Pressione um botão para obter informações sobre a impressora ou listar as impressoras instaladas.",
    justify="left"
)

btn_restart.pack(padx=20, pady=(20, 10), fill="x")
btn_refresh.pack(padx=20, pady=10, fill="x")
btn_list.pack(padx=20, pady=10, fill="x")
btn_select_folder.pack(padx=20, pady=10, fill="x")
btn_print_folder.pack(padx=20, pady=10, fill="x")
info_label.pack(padx=20, pady=(10, 20), fill="x")

folder_label = ctk.CTkLabel(
    app,
    text="Pasta selecionada: Nenhuma",
    justify="left"
)
folder_label.pack(padx=20, pady=(0, 20), fill="x")

app.mainloop()