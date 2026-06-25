import customtkinter as ctk
import win32api
import win32print
import win32service
import win32serviceutil
from tkinter import filedialog, messagebox
import os
import re
import shutil
import subprocess

NFE_NUMBER_PATTERN = re.compile(r"(?<!\d)\d{20}55\d{3}(\d{9})\d{10}(?!\d)")

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


def get_print_queue():
    try:
        printer_name = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
        win32print.ClosePrinter(handle)

        if not jobs:
            return f"Fila de impressão vazia para: {printer_name}"

        lines = [f"Fila de impressão ({printer_name}):"]
        for job in jobs:
            job_id = _printer_info_value(job, "JobId", 0)
            document = _printer_info_value(job, "pDocument", 4)
            status = _printer_info_value(job, "Status", 6)
            lines.append(f"ID {job_id} - {document} - Status {status}")
        return "\n".join(lines)
    except Exception as error:
        return f"Erro ao obter fila de impressão: {error}"


def show_print_queue():
    info_label.configure(text=get_print_queue())


def open_control_printers():
    try:
        subprocess.Popen(["control", "printers"])
        info_label.configure(text="Painel de impressão aberto.")
    except Exception as error:
        info_label.configure(text=f"Erro ao abrir Control Printers: {error}")


def extract_nfe_number(filename):
    match = NFE_NUMBER_PATTERN.search(filename)
    if not match:
        return None
    return int(match.group(1))


def show_nfe_print_selection(folder, pdf_files):
    nfe_files = []
    for pdf_file in pdf_files:
        nfe_number = extract_nfe_number(os.path.basename(pdf_file))
        if nfe_number is not None:
            nfe_files.append((nfe_number, pdf_file))

    if not nfe_files:
        info_label.configure(text=f"Nenhuma NF-e encontrada nos PDFs de:\n{folder}")
        return []

    nfe_files.sort(key=lambda item: item[0])
    min_nfe = nfe_files[0][0]
    max_nfe = nfe_files[-1][0]
    selected_files = None

    dialog = ctk.CTkToplevel(app)
    dialog.title("Seleção de NF-es para impressão")
    dialog.geometry("420x360")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()

    mode_var = ctk.StringVar(value="all")
    start_var = ctk.StringVar(value=str(min_nfe))
    end_var = ctk.StringVar(value=str(max_nfe))

    container = ctk.CTkFrame(dialog)
    container.pack(padx=20, pady=20, fill="both", expand=True)

    title_label = ctk.CTkLabel(
        container,
        text="Informações da pasta",
        font=ctk.CTkFont(size=16, weight="bold")
    )
    title_label.pack(anchor="w", pady=(0, 10))

    folder_info_label = ctk.CTkLabel(
        container,
        text=(
            f"Arquivos encontrados: {len(nfe_files)}\n\n"
            f"Menor NF-e encontrada: {min_nfe}\n"
            f"Maior NF-e encontrada: {max_nfe}"
        ),
        justify="left"
    )
    folder_info_label.pack(anchor="w", fill="x", pady=(0, 14))

    all_radio = ctk.CTkRadioButton(
        container,
        text="Imprimir todas as NF-es",
        variable=mode_var,
        value="all"
    )
    all_radio.pack(anchor="w", pady=(0, 8))

    range_radio = ctk.CTkRadioButton(
        container,
        text="Imprimir intervalo específico",
        variable=mode_var,
        value="range"
    )
    range_radio.pack(anchor="w", pady=(0, 12))

    range_frame = ctk.CTkFrame(container, fg_color="transparent")
    range_frame.pack(fill="x", pady=(0, 16))
    range_frame.grid_columnconfigure(1, weight=1)

    start_label = ctk.CTkLabel(range_frame, text="NF-e inicial:")
    start_label.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    start_entry = ctk.CTkEntry(range_frame, textvariable=start_var)
    start_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))

    end_label = ctk.CTkLabel(range_frame, text="NF-e final:")
    end_label.grid(row=1, column=0, sticky="w", padx=(0, 10))
    end_entry = ctk.CTkEntry(range_frame, textvariable=end_var)
    end_entry.grid(row=1, column=1, sticky="ew")

    def update_range_state(*_):
        state = "normal" if mode_var.get() == "range" else "disabled"
        start_entry.configure(state=state)
        end_entry.configure(state=state)

    def confirm_print():
        nonlocal selected_files
        if mode_var.get() == "all":
            selected_files = [pdf_file for _, pdf_file in nfe_files]
            dialog.destroy()
            return

        try:
            start_nfe = int(start_var.get().strip())
            end_nfe = int(end_var.get().strip())
        except ValueError:
            messagebox.showerror(
                "Intervalo inválido",
                "Informe apenas números nos campos de NF-e.",
                parent=dialog
            )
            return

        if start_nfe > end_nfe:
            messagebox.showerror(
                "Intervalo inválido",
                "A NF-e inicial deve ser menor ou igual a NF-e final.",
                parent=dialog
            )
            return

        selected_files = [
            pdf_file
            for nfe_number, pdf_file in nfe_files
            if start_nfe <= nfe_number <= end_nfe
        ]
        if not selected_files:
            messagebox.showerror(
                "Nenhuma NF-e encontrada",
                "Não existem PDFs no intervalo informado.",
                parent=dialog
            )
            return

        dialog.destroy()

    def cancel_print():
        dialog.destroy()

    mode_var.trace_add("write", update_range_state)
    update_range_state()

    buttons_frame = ctk.CTkFrame(container, fg_color="transparent")
    buttons_frame.pack(fill="x")
    buttons_frame.grid_columnconfigure((0, 1), weight=1)

    print_button = ctk.CTkButton(buttons_frame, text="Imprimir", command=confirm_print)
    print_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

    cancel_button = ctk.CTkButton(buttons_frame, text="Cancelar", command=cancel_print)
    cancel_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))

    dialog.protocol("WM_DELETE_WINDOW", cancel_print)
    dialog.focus()
    app.wait_window(dialog)

    return selected_files


def print_folder_files():
    global selected_folder
    folder = selected_folder or filedialog.askdirectory()
    if not folder:
        return

    pdf_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    ]
    if not pdf_files:
        info_label.configure(text=f"Nenhum PDF encontrado em:\n{folder}")
        return

    files_to_print = show_nfe_print_selection(folder, pdf_files)
    if not files_to_print:
        return

    try:
        for pdf_file in files_to_print:
            win32api.ShellExecute(0, "print", os.path.abspath(pdf_file), "", ".", 0)
        info_label.configure(
            text=f"Enviados para impressão {len(files_to_print)} PDF(s) de:\n{folder}"
        )
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

btn_print_queue = ctk.CTkButton(
    app,
    text="Fila de impressão",
    command=show_print_queue
)

btn_control_printers = ctk.CTkButton(
    app,
    text="Abrir painel de impressoras",
    command=open_control_printers
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
btn_print_queue.pack(padx=20, pady=10, fill="x")
btn_control_printers.pack(padx=20, pady=10, fill="x")
info_label.pack(padx=20, pady=(10, 20), fill="x")

folder_label = ctk.CTkLabel(
    app,
    text="Pasta selecionada: Nenhuma",
    justify="left"
)
folder_label.pack(padx=20, pady=(0, 20), fill="x")

app.mainloop()
