import os
import sys
import subprocess
import tkinter as tk
import tkinter.scrolledtext as ScrolledText
from tkinter import filedialog, messagebox
from PIL import Image
from datetime import datetime, time
from collections import defaultdict
import openpyxl
from logic.processor import processa_cartella_excel
from logic.comparer import confronta_file_cartellini
from logic.totals import esegui_count_tot
import threading
import customtkinter as ctk
from customtkinter import CTkImage


# Mostra un messaggio di errore all'utente in GUI se l'installazione fallisce
def mostra_errore_gui(msg):
    try:
        import tkinter as tk
        import tkinter.scrolledtext as ScrolledText
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Errore di installazione", msg)
    except Exception as e:
        print("Errore grave:", msg)

# Tenta di installare un modulo pip
def installa_modulo(modulo, nome_pip=None):
    nome_pip = nome_pip or modulo
    try:
        __import__(modulo)
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", nome_pip])
        except Exception as e:
            mostra_errore_gui(
                f"Non riesco a installare il modulo '{nome_pip}'.\n"
                f"Controlla la connessione a Internet o installa manualmente con:\n\npip install {nome_pip}"
            )
            sys.exit(1)

installa_modulo("openpyxl")
installa_modulo("PIL", nome_pip="pillow")
installa_modulo("customtkinter") 

os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Elaborazione Cartellini")
        self.root.geometry("700x700")
        self.root.resizable(False, False)
        self.root.configure(fg_color="white")

        self.font_base = ("Segoe UI", 10)
        self.font_bold = ("Segoe UI", 10, "bold")

        # Tab setup
        self.tab_control = ctk.CTkTabview(self.root)
        self.tab_control.add("Elabora cartellini")
        self.tab_control.add("Conta Totali")
        self.tab_control.add("Confronta")
        self.tab_elabora = self.tab_control.tab("Elabora cartellini")
        self.tab_count = self.tab_control.tab("Conta Totali")
        self.tab_confronta = self.tab_control.tab("Confronta")
        self.tab_control.pack(expand=1, fill="both")
        
        # Logo
        try:
            logo_elabora_img = Image.open("assets/icons/logo.png").resize((240, 240), Image.Resampling.LANCZOS)
            self.logo_elabora_img = CTkImage(logo_elabora_img)
            ctk.CTkLabel(self.tab_elabora, image=self.logo_elabora_img, text="", fg_color="transparent").pack(pady=10)
            
            logo_confronta_img = Image.open("assets/icons/logo.png").resize((240, 240), Image.Resampling.LANCZOS)
            self.logo_confronta_img = CTkImage(logo_confronta_img)
            ctk.CTkLabel(self.tab_confronta, image=self.logo_confronta_img, text="", fg_color="transparent").pack(pady=10, fill="x")

            # Rimosso il blocco del logo da setup_tab_count(), quindi usa quello già caricato in __init__
            logo_count_img = Image.open("assets/icons/logo.png").resize((240, 240), Image.Resampling.LANCZOS)
            self.logo_count_img = CTkImage(logo_count_img)
            ctk.CTkLabel(self.tab_count, image=self.logo_count_img, text="", fg_color="transparent").pack(pady=10, fill="x")

        except Exception as e:
            print(f"Errore caricamento logo: {e}")

        # Caricamento icone log
        try:
            self.icons = {
                "start": CTkImage(Image.open("assets/icons/icon_start.png").resize((16, 16), Image.Resampling.LANCZOS)),
                "end": CTkImage(Image.open("assets/icons/icon_end.png").resize((16, 16), Image.Resampling.LANCZOS)),
                "info": CTkImage(Image.open("assets/icons/icon_info.png").resize((16, 16), Image.Resampling.LANCZOS)),
                "success": CTkImage(Image.open("assets/icons/icon_success.png").resize((16, 16), Image.Resampling.LANCZOS)),
                "warning": CTkImage(Image.open("assets/icons/icon_warning.png").resize((16, 16), Image.Resampling.LANCZOS)),
                "error": CTkImage(Image.open("assets/icons/icon_error.png").resize((16, 16), Image.Resampling.LANCZOS))
            }
        except Exception as e:
            print(f"Errore nel caricamento delle icone: {e}")
            self.icons = {}

        # CARTELLA ESPORTAZIONI
        frame_cartella = ctk.CTkFrame(self.tab_elabora, fg_color="transparent")
        frame_cartella.pack(pady=(0, 10), padx=20, anchor="w")
        ctk.CTkLabel(frame_cartella, text="Cartella con esportazioni:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        self.btn_cartella = ctk.CTkButton(
            frame_cartella,
            text="Sfoglia...",
            font=self.font_bold,
            command=self.seleziona_cartella,
            fg_color="#A9A9A9",         # Grigio scuro per dare più risalto
            hover_color="#808080",      # Grigio ancora più scuro al passaggio del mouse
            text_color="white",         # Testo bianco per maggiore leggibilità
            corner_radius=6
        )
        self.btn_cartella.pack(side=tk.LEFT, padx=10)

        self.entry_cartella_selezionata = ctk.CTkEntry(
            self.tab_elabora,
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white",
            state="readonly",
            width=400
        )
        self.entry_cartella_selezionata.pack(pady=(0, 10), padx=20, anchor="w")

        # FILE ORE SETTIMANALI
        frame_ore = ctk.CTkFrame(self.tab_elabora, fg_color="transparent")
        frame_ore.pack(pady=(0, 10), padx=20, anchor="w")
        ctk.CTkLabel(frame_ore, text="File presenze Factorial:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        self.btn_ore_settimanali = ctk.CTkButton(
            frame_ore,
            text="Sfoglia...",
            font=self.font_bold,
            command=self.seleziona_file_ore_settimanali,
            fg_color="#A9A9A9",         # Grigio scuro per dare più risalto
            hover_color="#808080",      # Grigio ancora più scuro al passaggio del mouse
            text_color="white",         # Testo bianco per maggiore leggibilità
            corner_radius=6
        )
        self.btn_ore_settimanali.pack(side=tk.LEFT, padx=10)

        self.entry_file_ore_settimanali = ctk.CTkEntry(
            self.tab_elabora,
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white",
            state="readonly",
            width=400
        )
        self.entry_file_ore_settimanali.pack(pady=(0, 10), padx=20, anchor="w")

        # NOME FILE OUTPUT
        frame_output = ctk.CTkFrame(self.tab_elabora, fg_color="transparent")
        frame_output.pack(padx=20, anchor="w")
        ctk.CTkLabel(frame_output, text="Nome file output:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        # MODIFICATO: da tk.Entry a ctk.CTkEntry
        self.entry_output = ctk.CTkEntry(
            frame_output, 
            width=400, 
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white"
        )
        self.entry_output.insert(0, datetime.now().strftime("Cartellini_%d%B_%H.%M.xlsx"))
        self.entry_output.pack(pady=(0, 10), padx=20, anchor="w")

        # PROGRESS BAR
        self.progress = ctk.CTkProgressBar(self.tab_elabora)
        self.progress.set(0)
        self.progress.pack(pady=(0, 10), padx=20, fill="x", expand=True)

        # BOTTONE AVVIA
        self.btn_avvia = ctk.CTkButton(
            self.tab_elabora,
            text="Avvia elaborazione",
            command=self.avvia_elaborazione,
            font=("Segoe UI", 11, "bold"),       # Uniforma il font
            fg_color="#4CAF50",                  # Colore verde per coerenza
            hover_color="#45A049",
            text_color="white",
            corner_radius=8,                     # Raggio angoli uniforme
            width=180,                           # Larghezza uniforme
            height=36                            # Altezza uniforme
        )
        self.btn_avvia.pack(pady=(0, 10), padx=20)

        # FRAME POST-ELABORAZIONE
        self.frame_post_elaborazione = ctk.CTkFrame(self.tab_elabora, fg_color="transparent")
        self.frame_post_elaborazione.pack(pady=(0, 10), padx=20)

        # LOG OUTPUT
        self.log_frame = ctk.CTkFrame(self.tab_elabora, fg_color="transparent")
        self.log_frame.pack(pady=(0, 10), padx=20, anchor="w")

        self.log = ScrolledText.ScrolledText(
            self.log_frame,
            height=12,
            state='disabled',
            bg="#f5f5f5",
            fg="#333333",
            font=("Segoe UI Emoji", 10),
            wrap='word'
        )
        self.log.pack(fill="both", expand=True)
        self._configura_log_box(self.log)

        # Variabili
        self.cartella = None
        self.file_ore_settimanali = None
        self.file_vecchio = None
        self.file_nuovo = None

        # Setup tab "Confronta"
        self.confronta_due_file_cartellini()
        self.setup_tab_count()


    def _configura_log_box(self, log_box):
        """Metodo per configurare in modo coerente i tag di stile dei log box."""
        log_box.tag_config('info', foreground="#555")
        log_box.tag_config('start', foreground="#444")
        log_box.tag_config('success', foreground="#2e7d32")
        log_box.tag_config('error', foreground="#d32f2f")
        log_box.tag_config('warning', foreground="#ed6c02")
        log_box.tag_config('blank', foreground="#333")
        log_box.tag_config("end", foreground="#444")

    def logga(self, log_box, messaggio, tag="info"):
        log_box.configure(state='normal')
        emoji = {
            "start": "🚀",
            "success": "✅",
            "error": "❌",
            "warning": "⚠️",
            "info": "ℹ️",
            "end": "✅",
            "sun": "☀️"
        }
        simbolo = emoji.get(tag, "")
        log_box.insert(tk.END, f"{simbolo} {messaggio}\n", tag)
        log_box.see(tk.END)
        log_box.configure(state='disabled')

    def seleziona_cartella(self):
        self.cartella = filedialog.askdirectory()
        if self.cartella:
            self.entry_cartella_selezionata.configure(state='normal')
            self.entry_cartella_selezionata.delete(0, tk.END)
            self.entry_cartella_selezionata.insert(0, self.cartella)
            self.entry_cartella_selezionata.configure(state='readonly')

    def seleziona_file_ore_settimanali(self):
        self.file_ore_settimanali = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if self.file_ore_settimanali:
            self.entry_file_ore_settimanali.configure(state='normal')
            self.entry_file_ore_settimanali.delete(0, tk.END)
            self.entry_file_ore_settimanali.insert(0, self.file_ore_settimanali)
            self.entry_file_ore_settimanali.configure(state='readonly')

    def avvia_elaborazione(self):
        if not self.cartella:
            messagebox.showerror("Errore", "Seleziona prima una cartella con i file di esportazione.")
            return

        if not self.file_ore_settimanali:
            messagebox.showerror("Errore", "Seleziona prima il file con le ore settimanali.")
            return

        nome_output = self.entry_output.get().strip()
        if not nome_output:
            messagebox.showerror("Errore", "Inserisci un nome per il file di output.")
            return

        if not nome_output.endswith(".xlsx"):
            nome_output += ".xlsx"

        base, ext = os.path.splitext(nome_output)
        counter = 1
        while os.path.exists(nome_output):
            nome_output = f"{base}_{counter}{ext}"
            counter += 1

        files_excel = [f for f in os.listdir(self.cartella) if f.endswith((".xlsx", ".xlsm"))]
        total_files = len(files_excel)

        if total_files == 0:
            self.logga(self.log, "Nessun file Excel trovato nella cartella.", "warning")
            return

        self.progress.set(0)
        self.root.update_idletasks()

        self.logga(self.log, "Perfetto! Dammi un attimo...", "start")

        try:
            def aggiorna_progress(i):
                self.progress.set(i / total_files)
                self.root.update_idletasks()

            processa_cartella_excel(
                self.cartella,
                nome_output,
                self.file_ore_settimanali,
                logger=lambda msg, tag="info": self.logga(self.log, msg, tag),
                progress_callback=aggiorna_progress
            )

            self.logga(self.log, f"File salvato con successo: {nome_output}", "success")
            self.logga(self.log, " ", tag="blank")
            self.logga(self.log, " ", tag="blank")
            self.logga(self.log, "Buon lavoro!", "sun")

            for widget in self.frame_post_elaborazione.winfo_children():
                widget.destroy()

            btn_apri_file = ctk.CTkButton(
                master=self.frame_post_elaborazione,
                text="Apri file dei cartellini",
                command=lambda: os.startfile(nome_output),
                fg_color="#4CAF50",
                hover_color="#45A049",
                text_color="white",
                font=self.font_bold,
                corner_radius=6,
                width=180,
                height=36
            )
            btn_apri_file.pack(pady=5, padx=10)

            risposta = messagebox.askyesno(
                "Calcolo Totali",
                "Vuoi eseguire anche il calcolo dei totali sul file appena generato?"
            )

            if risposta:
                try:
                    self.progress.set(0)
                    self.root.update_idletasks()

                    def progress_callback(frazione):
                        self.progress.set(frazione)
                        self.root.update_idletasks()

                    output_count_tot = nome_output.replace(".xlsx", "_totali.xlsx")
                    base, ext = os.path.splitext(output_count_tot)
                    counter = 1
                    while os.path.exists(output_count_tot):
                        output_count_tot = f"{base}_{counter}{ext}"
                        counter += 1

                    esegui_count_tot(nome_output, output_count_tot, progress_callback)
                    self.logga(self.log, f"Totali calcolati e salvati in: {output_count_tot}", "success")

                    btn_apri_totali = ctk.CTkButton(
                        master=self.frame_post_elaborazione,
                        text="Apri conteggi",
                        command=lambda: os.startfile(output_count_tot),
                        fg_color="#FF9800",
                        hover_color="#FB8C00",
                        text_color="white",
                        font=("Arial", 10, "bold"),
                        corner_radius=6,
                        width=160,
                        height=36
                    )
                    btn_apri_totali.pack(pady=5, padx=10)

                except Exception as e:
                    self.logga(self.log, f"Errore nel calcolo dei totali: {e}", "error")
                    messagebox.showerror("Errore Totali", str(e))
        except Exception as e:
            self.logga(self.log, f"Errore: {e}", "error")
            messagebox.showerror("Errore durante l'elaborazione", str(e))
        finally:
            self.progress.set(0)

    def seleziona_file(self, tipo):
        path = filedialog.askopenfilename(title=f"Seleziona file cartellini {tipo.upper()}", filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            if tipo == "vecchio":
                self.file_vecchio = path
                self.lbl_file_vecchio.configure(state='normal')
                self.lbl_file_vecchio.delete(0, tk.END)
                self.lbl_file_vecchio.insert(0, path)
                self.lbl_file_vecchio.configure(state='readonly')
            elif tipo == "nuovo":
                self.file_nuovo = path
                self.lbl_file_nuovo.configure(state='normal')
                self.lbl_file_nuovo.delete(0, tk.END)
                self.lbl_file_nuovo.insert(0, path)
                self.lbl_file_nuovo.configure(state='readonly')

    def esegui_confronto(self):
        self.progressbar_confronta.set(0)
        self.root.update_idletasks()
        self.logga(self.text_output_confronto, "Inizio confronto file cartellini...", "start")

        file_vecchio = self.file_vecchio
        file_nuovo = self.file_nuovo
        output_path = self.entry_output_confronto.get().strip()

        if not file_vecchio or not os.path.isfile(file_vecchio):
            messagebox.showerror("Errore", "Seleziona un file cartellini VECCHIO valido.")
            return

        if not file_nuovo or not os.path.isfile(file_nuovo):
            messagebox.showerror("Errore", "Seleziona un file cartellini NUOVO valido.")
            return

        if not output_path:
            messagebox.showerror("Errore", "Inserisci un nome per il file di confronto.")
            return

        if not output_path.endswith(".xlsx"):
            output_path += ".xlsx"

        base, ext = os.path.splitext(output_path)
        counter = 1
        while os.path.exists(output_path):
            output_path = f"{base}_{counter}{ext}"
            counter += 1

        try:
            def progress_callback(frazione):
                self.progressbar_confronta.set(frazione)
                self.root.update_idletasks()

            confronta_file_cartellini(file_vecchio, file_nuovo, output_path, progress_callback=progress_callback)
            self.logga(self.text_output_confronto, f"Confronto completato. File salvato in: {output_path}", "success")
            self.progressbar_confronta.set(1)

            for widget in self.frame_post_confronto.winfo_children():
                widget.destroy()

            btn_apri_confronto = ctk.CTkButton(
                master=self.frame_post_confronto,
                text="Apri file confronto",
                command=lambda: os.startfile(output_path),
                fg_color="#2196F3",
                hover_color="#1976D2",
                text_color="white",
                font=("Arial", 10, "bold"),
                corner_radius=6,
                width=180,
                height=36
            )
            btn_apri_confronto.pack(pady=(0, 10), padx=20)
        except Exception as e:
            self.logga(self.text_output_confronto, f"Errore durante il confronto: {e}", "error")
            messagebox.showerror("Errore", f"Errore durante il confronto: {e}")
        finally:
            self.progressbar_confronta.set(0)

    def confronta_due_file_cartellini(self):
        font_base = ("Segoe UI", 10)
        font_bold = ("Segoe UI", 10, "bold")

        # File VECCHIO
        frame_file_vecchio = ctk.CTkFrame(self.tab_confronta, fg_color="transparent")
        frame_file_vecchio.pack(pady=(0, 10), padx=20, anchor="w")
        ctk.CTkLabel(frame_file_vecchio, text="File cartellini VECCHIO:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        btn_file_vecchio = ctk.CTkButton(
            frame_file_vecchio,
            text="Sfoglia...",
            font=self.font_bold,
            command=lambda: self.seleziona_file("vecchio"),
            fg_color="#A9A9A9",         # Grigio scuro per dare più risalto
            hover_color="#808080",      # Grigio ancora più scuro al passaggio del mouse
            text_color="white",         # Testo bianco per maggiore leggibilità
            corner_radius=6
        )
        btn_file_vecchio.pack(side=tk.LEFT, padx=10)

        # MODIFICATO: Da CTkLabel a CTkEntry (readonly) per uniformità
        self.lbl_file_vecchio = ctk.CTkEntry(
            self.tab_confronta,
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white",
            state="readonly",
            width=400
        )
        self.lbl_file_vecchio.pack(pady=(0, 10), padx=20, anchor="w")

        # File NUOVO
        frame_file_nuovo = ctk.CTkFrame(self.tab_confronta, fg_color="transparent")
        frame_file_nuovo.pack(pady=(0, 10), padx=20, anchor="w")
        ctk.CTkLabel(frame_file_nuovo, text="File cartellini NUOVO:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        btn_file_nuovo = ctk.CTkButton(
            frame_file_nuovo,
            text="Sfoglia...",
            font=self.font_bold,
            command=lambda: self.seleziona_file("nuovo"),
            fg_color="#A9A9A9",         # Grigio scuro per dare più risalto
            hover_color="#808080",      # Grigio ancora più scuro al passaggio del mouse
            text_color="white",         # Testo bianco per maggiore leggibilità
            corner_radius=6
        )
        btn_file_nuovo.pack(side=tk.LEFT, padx=10)

        # MODIFICATO: Da CTkLabel a CTkEntry (readonly) per uniformità
        self.lbl_file_nuovo = ctk.CTkEntry(
            self.tab_confronta,
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white",
            state="readonly",
            width=400
        )
        self.lbl_file_nuovo.pack(pady=(0, 10), padx=20, anchor="w")

        # Output filename
        frame_output_confronto = ctk.CTkFrame(self.tab_confronta, fg_color="transparent")
        frame_output_confronto.pack(padx=20, anchor="w")
        ctk.CTkLabel(frame_output_confronto, text="Nome file di confronto:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        # MODIFICATO: da tk.Entry a ctk.CTkEntry
        self.entry_output_confronto = ctk.CTkEntry(
            frame_output_confronto, 
            width=400, 
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white"
        )
        default_name = datetime.now().strftime("Confronto_%d%B_%H.%M.xlsx")
        self.entry_output_confronto.insert(0, default_name)
        self.entry_output_confronto.pack(pady=(0, 10), padx=20, anchor="w")

        # Progress bar
        self.progressbar_confronta = ctk.CTkProgressBar(self.tab_confronta)
        self.progressbar_confronta.set(0)
        self.progressbar_confronta.pack(pady=(0, 10), padx=20, fill="x", expand=True)

        # Bottone Confronta
        self.btn_esegui_confronto = ctk.CTkButton(
            self.tab_confronta,
            text="Esegui confronto",
            command=self.esegui_confronto,
            font=("Segoe UI", 11, "bold"),
            fg_color="#2196F3",
            hover_color="#42A5F5",
            text_color="white",
            corner_radius=6,
            width=180,
            height=36
        )
        self.btn_esegui_confronto.pack(pady=(0, 10), padx=20)

        # Post confronto
        self.frame_post_confronto = ctk.CTkFrame(self.tab_confronta, fg_color="transparent")
        self.frame_post_confronto.pack(pady=(0, 10), padx=20)

        # Log box
        self.text_output_confronto = ScrolledText.ScrolledText(
            self.tab_confronta,
            height=12,
            state='disabled',
            bg="#f5f5f5",
            fg="#333333",
            font=("Segoe UI", 10),
            wrap='word'
        )
        self.text_output_confronto.pack(pady=(0, 10), padx=20, anchor="w")
        self._configura_log_box(self.text_output_confronto)


    def setup_tab_count(self):
        font_base = ("Segoe UI", 10)
        font_bold = ("Segoe UI", 10, "bold")

        # Selezione file
        frame_file = ctk.CTkFrame(self.tab_count, fg_color="transparent")
        frame_file.pack(pady=(0, 10), padx=20, anchor="w")

        ctk.CTkLabel(frame_file, text="File da elaborare:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)
        btn_sfoglia = ctk.CTkButton(
            frame_file,
            text="Sfoglia...",
            font=self.font_bold,
            fg_color="#A9A9A9",         # Grigio scuro per dare più risalto
            hover_color="#808080",      # Grigio ancora più scuro al passaggio del mouse
            text_color="white",         # Testo bianco per maggiore leggibilità
            corner_radius=6,
            command=self.seleziona_file_count
        )

        btn_sfoglia.pack(side=tk.LEFT, padx=10)

        # MODIFICATO: Da CTkLabel a CTkEntry (readonly) per uniformità
        self.lbl_file_count = ctk.CTkEntry(
            self.tab_count, 
            text_color="#005bbb", 
            fg_color="white", 
            font=self.font_base,
            state="readonly",
            width=400
        )
        self.lbl_file_count.pack(pady=(0, 10), padx=20, anchor="w")

        # Output filename
        frame_output = ctk.CTkFrame(self.tab_count, fg_color="transparent")
        frame_output.pack(padx=20, anchor="w")
        ctk.CTkLabel(frame_output, text="Nome file output:", font=self.font_base, fg_color="transparent").pack(side=tk.LEFT)

        # MODIFICATO: da tk.Entry a ctk.CTkEntry
        self.entry_output_count = ctk.CTkEntry(
            frame_output, 
            width=400, 
            font=self.font_base,
            text_color="#005bbb",
            fg_color="white"
        )
        default_name = datetime.now().strftime("Totali_%d%B_%H.%M.xlsx")
        self.entry_output_count.insert(0, default_name)
        self.entry_output_count.pack(pady=(0, 10), padx=20, anchor="w")

        # Progress bar
        self.progressbar_count = ctk.CTkProgressBar(self.tab_count)
        self.progressbar_count.set(0)
        self.progressbar_count.pack(pady=(0, 10), padx=20, fill="x", expand=True)

        # Bottone avvia
        self.btn_avvia_count_tot = ctk.CTkButton(
            self.tab_count,
            text="Calcola i totali",
            command=self.avvia_count_tot,
            font=("Segoe UI", 11, "bold"),
            fg_color="#FB8C00",
            hover_color="#FFA726",
            text_color="white",
            corner_radius=8,
            width=180,
            height=36
        )
        self.btn_avvia_count_tot.pack(pady=(0, 10), padx=20)

        # Post-elaborazione
        self.frame_post_count = ctk.CTkFrame(self.tab_count, fg_color="transparent")
        self.frame_post_count.pack(pady=(0, 10), padx=20)

        # Log
        self.text_output_count = ScrolledText.ScrolledText(
            self.tab_count,
            height=12,
            state='disabled',
            bg="#f5f5f5",
            fg="#333333",
            font=("Segoe UI", 10),
            wrap='word'
        )
        self.text_output_count.pack(pady=(0, 10), padx=20, anchor="w")
        self._configura_log_box(self.text_output_count)


    def seleziona_file_count(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.file_count_path = path
            # MODIFICATO: Aggiorna il valore della CTkEntry
            self.lbl_file_count.configure(state='normal')
            self.lbl_file_count.delete(0, tk.END)
            self.lbl_file_count.insert(0, path)
            self.lbl_file_count.configure(state='readonly')


    def avvia_count_tot(self):
        file_input = getattr(self, "file_count_path", None)
        file_output = self.entry_output_count.get().strip()

        if not file_input or not os.path.isfile(file_input):
            messagebox.showerror("Errore", "Seleziona un file Excel valido.")
            return

        if not file_output:
            messagebox.showerror("Errore", "Inserisci un nome per il file di output.")
            return

        if not file_output.endswith(".xlsx"):
            file_output += ".xlsx"

        base, ext = os.path.splitext(file_output)
        counter = 1
        while os.path.exists(file_output):
            file_output = f"{base}_{counter}{ext}"
            counter += 1

        self.progressbar_count.set(0)
        self.root.update_idletasks()

        self.logga(self.text_output_count, "Inizio elaborazione...", "start")

        def run():
            def progress_callback(frazione):
                self.progressbar_count.set(frazione)
                self.root.update_idletasks()

            try:
                esegui_count_tot(file_input, file_output, progress_callback=progress_callback)
                self.logga(self.text_output_count, f"Operazione completata.\nFile salvato in: {file_output}", "success")
                self.progressbar_count.set(1)

                for widget in self.frame_post_count.winfo_children():
                    widget.destroy()

                btn_apri_file = ctk.CTkButton(
                    master=self.frame_post_count,
                    text="Apri file totali",
                    command=lambda: os.startfile(file_output),
                    fg_color="#FF9800",
                    hover_color="#FB8C00",
                    text_color="white",
                    font=("Arial", 10, "bold"),
                    corner_radius=6,
                    width=180,
                    height=36
                )
                btn_apri_file.pack(pady=(0, 10), padx=20)

            except Exception as e:
                self.logga(self.text_output_count, f"Errore durante l'elaborazione: {e}", "error")
                messagebox.showerror("Errore", str(e))
            finally:
                self.progressbar_count.set(0)

        threading.Thread(target=run).start()

def launch_app():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    launch_app()