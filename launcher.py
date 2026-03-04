"""
LaunchC39B - Launcher per pipeline di script Python
Versione 3.2 - 01/03/2026
"""

import sys
import os
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, scrolledtext
from datetime import datetime
from pathlib import Path

import yaml
from rich.console import Console
from rich.text import Text

# ---------------------------------------------------------------------------
# Versione release
# ---------------------------------------------------------------------------
VER_NAME = "1.0 - Ariel"

# ---------------------------------------------------------------------------
# Nome dello script Sgamatore (case insensitive nei confronti)
# Usato per la logica use_transl_as_act
# ---------------------------------------------------------------------------
SGAMATORE_NAME = "sgamatore"

# ---------------------------------------------------------------------------
# Console Rich per output colorato su terminale
# ---------------------------------------------------------------------------
console = Console()

# Stili Rich per ciascun livello di log
RICH_STYLES = {
    "INFO":    "white",
    "WARNING": "bold yellow",
    "ERROR":   "bold red",
}

# ---------------------------------------------------------------------------
# Colori per il log nella finestra GUI (tkinter tag) - tema scuro
# ---------------------------------------------------------------------------
LOG_COLORS = {
    "INFO":    "#d4d4d4",   # grigio chiaro
    "WARNING": "#e5c07b",   # giallo ambra
    "ERROR":   "#e06c75",   # rosso soft
    "CMD":     "#61afef",   # blu - per le righe Comando:
    "DONE":    "#98c379",   # verde - per completamento script
}

LOG_BG = "#1e1e1e"
LOG_FG = "#d4d4d4"


# ===========================================================================
# CARICAMENTO CONFIGURAZIONE
# ===========================================================================

def load_tools(tools_path: Path) -> list[dict]:
    """
    Legge tools.yml e restituisce la lista degli script configurati.
    Esce con errore se il file non esiste, non è leggibile,
    o se uno script configurato non è trovato nel path specificato.
    """
    if not tools_path.exists():
        print(f"[ERROR] File tools.yml non trovato: {tools_path}")
        sys.exit(1)

    with open(tools_path, "r", encoding="utf-8") as f:
        tools = yaml.safe_load(f)

    if not tools:
        print("[ERROR] tools.yml è vuoto o non valido.")
        sys.exit(1)

    # Verifica che ogni script esista nel path dichiarato
    for entry in tools:
        script_path = tools_path.parent / entry["script_path"]
        if not script_path.exists():
            print(f"[ERROR] Script non trovato: {script_path}")
            sys.exit(1)

    return tools


def load_configs(config_path: Path, tools: list[dict]) -> dict | None:
    """
    Legge config.yml e restituisce il dizionario delle configurazioni.
    In caso di errore restituisce None (la listbox resterà vuota).
    Valida che i nomi script nelle configurazioni esistano in tools.yml.
    """
    if not config_path.exists():
        return None

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except yaml.YAMLError as e:
        print(f"[ERROR] Errore parsing config.yml: {e}")
        return None

    if not data or "configurations" not in data:
        print("[ERROR] config.yml non contiene il blocco 'configurations'.")
        return None

    # Nomi script validi (case insensitive)
    valid_names = {entry["script_name"].lower() for entry in tools}

    configs = data["configurations"]
    for config_name, config_body in configs.items():
        if config_body is None:
            continue
        for key in config_body:
            # Le chiavi di pipeline-level (es. use_transl_as_act) non sono script
            if key.lower() == "use_transl_as_act":
                continue
            if key.lower() not in valid_names:
                print(f"[ERROR] config.yml: script '{key}' non trovato in tools.yml.")
                return None

    return configs


# ===========================================================================
# GUI
# ===========================================================================

class LauncherApp:
    def __init__(self, root: tk.Tk, tools: list[dict], configs: dict | None):
        self.root = root
        self.tools = tools
        self.configs = configs or {}

        # Stato checkbox per ogni script: BooleanVar indicizzato per script_name
        self.script_vars: dict[str, tk.BooleanVar] = {}

        self._build_gui()
        self._log(f"LaunchC39B {VER_NAME} avviato.", "INFO")
        self._log(f"Caricati {len(self.tools)} script da tools.yml.", "INFO")
        if configs is None:
            self._log("config.yml non disponibile o con errori: listbox vuota.", "WARNING")
        else:
            self._log(f"Caricate {len(self.configs)} configurazioni da config.yml.", "INFO")

    # -----------------------------------------------------------------------
    # Costruzione GUI
    # -----------------------------------------------------------------------

    def _build_gui(self):
        self.root.title(f"LaunchC39B {VER_NAME}")
        self.root.resizable(True, True)

        # Frame principale con padding
        main = tk.Frame(self.root, padx=10, pady=10)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Riga: File BUD (facoltativo, snapshot precedente) ---
        row = tk.Frame(main)
        row.pack(fill=tk.X, pady=2)
        tk.Label(row, text="File BUD:", width=12, anchor="w").pack(side=tk.LEFT)
        self.var_file_old = tk.StringVar()
        tk.Entry(row, textvariable=self.var_file_old, width=72).pack(side=tk.LEFT, padx=4)
        tk.Button(row, text="...", command=self._browse_file_old).pack(side=tk.LEFT)

        # --- Riga: File ACT (obbligatorio, dataset corrente) ---
        row = tk.Frame(main)
        row.pack(fill=tk.X, pady=2)
        tk.Label(row, text="File ACT:", width=12, anchor="w").pack(side=tk.LEFT)
        self.var_file_act = tk.StringVar()
        self.var_file_act.trace_add("write", self._on_file_act_changed)
        tk.Entry(row, textvariable=self.var_file_act, width=72).pack(side=tk.LEFT, padx=4)
        tk.Button(row, text="...", command=self._browse_file_act).pack(side=tk.LEFT)

        # --- Riga: Anno + Cartella output ---
        row = tk.Frame(main)
        row.pack(fill=tk.X, pady=2)

        tk.Label(row, text="Anno:", width=12, anchor="w").pack(side=tk.LEFT)
        self.var_anno = tk.StringVar(value=str(datetime.now().year))
        tk.Entry(row, textvariable=self.var_anno, width=6).pack(side=tk.LEFT, padx=4)

        tk.Label(row, text="  Output:", anchor="w").pack(side=tk.LEFT)
        # Nome cartella di default: Launch_YYMMDD_HHMM al momento del lancio
        default_out = datetime.now().strftime("Launch_%y%m%d_%H%M")
        self.var_output = tk.StringVar(value=default_out)
        tk.Entry(row, textvariable=self.var_output, width=36).pack(side=tk.LEFT, padx=4)
        tk.Button(row, text="...", command=self._browse_output).pack(side=tk.LEFT)

        # --- Separatore ---
        tk.Frame(main, height=1, bg="lightgray").pack(fill=tk.X, pady=6)

        # --- Area centrale: listbox config + frame script ---
        center = tk.Frame(main)
        center.pack(fill=tk.BOTH, expand=False)

        # Listbox configurazioni
        lb_frame = tk.LabelFrame(center, text="Configurazioni", padx=5, pady=5)
        lb_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        self.listbox = tk.Listbox(lb_frame, width=36, height=8, exportselection=False)
        self.listbox.pack(side=tk.LEFT, fill=tk.Y)
        scrollbar = tk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        for name in self.configs:
            self.listbox.insert(tk.END, name)

        # Selezione configurazione aggiorna checkbox
        self.listbox.bind("<<ListboxSelect>>", self._on_config_selected)

        # Frame script (uno per script)
        scripts_frame = tk.LabelFrame(center, text="Script", padx=5, pady=5)
        scripts_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for entry in self.tools:
            name = entry["script_name"]
            var = tk.BooleanVar(value=True)
            self.script_vars[name] = var
            f = tk.Frame(scripts_frame)
            f.pack(anchor="w", pady=1)
            tk.Checkbutton(f, text=name, variable=var).pack(side=tk.LEFT)

        # --- Pulsanti Start / Quit ---
        btn_frame = tk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=6)

        self.btn_start = tk.Button(
            btn_frame, text="Start", width=12,
            command=self._on_start, state=tk.DISABLED
        )
        self.btn_start.pack(side=tk.LEFT, padx=4)

        tk.Button(
            btn_frame, text="Quit", width=12,
            command=self.root.destroy
        ).pack(side=tk.LEFT, padx=4)

        # Checkbox: apri cartella di output al termine della pipeline
        self.var_open_folder = tk.BooleanVar(value=True)
        tk.Checkbutton(
            btn_frame, text="Apri cartella al termine",
            variable=self.var_open_folder
        ).pack(side=tk.LEFT, padx=16)

        # --- Finestra di log ---
        log_frame = tk.LabelFrame(main, text="Log", padx=5, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        self.log_widget = scrolledtext.ScrolledText(
            log_frame, height=14, state=tk.DISABLED,
            font=("Courier New", 9), wrap=tk.WORD,
            background=LOG_BG, foreground=LOG_FG,
            insertbackground=LOG_FG
        )
        self.log_widget.pack(fill=tk.BOTH, expand=True)

        # Tag colori per i livelli di log (tema scuro)
        for level, color in LOG_COLORS.items():
            self.log_widget.tag_config(level, foreground=color)

    # -----------------------------------------------------------------------
    # Logging
    # -----------------------------------------------------------------------

    def _log(self, message: str, level: str = "INFO"):
        """
        Scrive una riga di log nella finestra GUI e su console (via Rich).
        Formato: [HH:MM:SS] LEVEL  messaggio

        Nella GUI usa tag distinti per livello; rileva automaticamente le
        righe "Comando:" (tag CMD) e i messaggi di completamento (tag DONE).
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {level:<8} {message}\n"

        # --- Console via Rich ---
        style = RICH_STYLES.get(level, "white")
        # Evidenzia timestamp in grigio, livello con lo stile del livello
        rich_line = (
            f"[dim][{timestamp}][/dim] "
            f"[{style}]{level:<8}[/{style}] "
            f"[{style}]{message}[/{style}]"
        )
        console.print(rich_line)

        # --- GUI: determina il tag più appropriato ---
        if level == "ERROR":
            tag = "ERROR"
        elif level == "WARNING":
            tag = "WARNING"
        elif message.startswith("  Comando:"):
            tag = "CMD"
        elif "completato" in message.lower():
            tag = "DONE"
        else:
            tag = "INFO"

        def _write():
            self.log_widget.config(state=tk.NORMAL)
            self.log_widget.insert(tk.END, line, tag)
            self.log_widget.see(tk.END)
            self.log_widget.config(state=tk.DISABLED)

        # Se chiamato da thread secondario usa after, altrimenti diretto
        if threading.current_thread() is threading.main_thread():
            _write()
        else:
            self.root.after(0, _write)

    # -----------------------------------------------------------------------
    # Browse helpers
    # -----------------------------------------------------------------------

    def _browse_file_act(self):
        path = filedialog.askopenfilename(
            title="Seleziona File ACT",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.var_file_act.set(path)

    def _browse_file_old(self):
        path = filedialog.askopenfilename(
            title="Seleziona File BUD",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.var_file_old.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Seleziona cartella di output")
        if path:
            self.var_output.set(path)

    # -----------------------------------------------------------------------
    # Logica UI
    # -----------------------------------------------------------------------

    def _on_file_act_changed(self, *_):
        """Abilita/disabilita Start in base alla presenza di File ACT."""
        if self.var_file_act.get().strip():
            self.btn_start.config(state=tk.NORMAL)
        else:
            self.btn_start.config(state=tk.DISABLED)

    def _on_config_selected(self, _event=None):
        """
        Aggiorna i checkbox in base alla configurazione selezionata.
        La selezione dalla listbox non sovrascrive modifiche manuali
        successive — semplicemente imposta lo stato al momento della
        selezione; l'utente può poi modificare liberamente i checkbox.
        """
        selection = self.listbox.curselection()
        if not selection:
            return
        config_name = self.listbox.get(selection[0])
        config_body = self.configs.get(config_name, {})

        for entry in self.tools:
            name = entry["script_name"]
            var = self.script_vars[name]
            script_cfg = config_body.get(name)

            if script_cfg is None:
                # Script non menzionato nella config: lascia invariato
                continue
            elif isinstance(script_cfg, bool):
                var.set(script_cfg)
            elif isinstance(script_cfg, dict):
                var.set(script_cfg.get("enabled", True))

    def _get_selected_config(self) -> tuple[str | None, dict]:
        """Restituisce (nome, dizionario) della configurazione selezionata."""
        selection = self.listbox.curselection()
        if not selection:
            return None, {}
        name = self.listbox.get(selection[0])
        return name, self.configs.get(name, {})

    # -----------------------------------------------------------------------
    # Apertura cartella output
    # -----------------------------------------------------------------------

    def _open_output_folder(self, path: str):
        """Apre la cartella di output nel file manager di sistema."""
        import platform
        import subprocess as sp
        try:
            folder = str(Path(path).resolve())
            system = platform.system()
            if system == "Windows":
                sp.Popen(["explorer", folder])
            elif system == "Darwin":
                sp.Popen(["open", folder])
            else:
                sp.Popen(["xdg-open", folder])
        except Exception as e:
            self._log(f"Impossibile aprire la cartella: {e}", "WARNING")

    # -----------------------------------------------------------------------
    # Avvio pipeline
    # -----------------------------------------------------------------------

    def _on_start(self):
        """Validazioni pre-esecuzione, poi lancia la pipeline in un thread."""
        file_act = self.var_file_act.get().strip()
        file_old = self.var_file_old.get().strip()
        anno = self.var_anno.get().strip()
        output_dir = self.var_output.get().strip()

        config_name, config_body = self._get_selected_config()
        use_transl = config_body.get("use_transl_as_act", False)

        # --- Vincolo: use_transl_as_act richiede Sgamatore abilitato ---
        if use_transl:
            sgamatore_enabled = False
            for name, var in self.script_vars.items():
                if name.lower() == SGAMATORE_NAME and var.get():
                    sgamatore_enabled = True
                    break
            if not sgamatore_enabled:
                self._log(
                    "ERRORE: use_transl_as_act è attivo ma Sgamatore è disabilitato. "
                    "Abilitare Sgamatore o deselezionare la configurazione.",
                    "ERROR"
                )
                return

        # --- Warning se File BUD non specificato ---
        if not file_old:
            self._log("File BUD non specificato: --file_bud verrà omesso dalla chiamata agli script.", "WARNING")

        # --- Disabilita Start per tutta la durata ---
        self.btn_start.config(state=tk.DISABLED)

        # --- Lancia pipeline in thread separato ---
        thread = threading.Thread(
            target=self._run_pipeline,
            args=(file_act, file_old, anno, output_dir, config_body, use_transl),
            daemon=True
        )
        thread.start()

    def _run_pipeline(
        self,
        file_act: str,
        file_old: str,
        anno: str,
        output_dir: str,
        config_body: dict,
        use_transl: bool,
    ):
        """
        Esegue la pipeline sequenzialmente in un thread separato.
        Ogni script viene invocato tramite subprocess con sys.executable.
        """
        try:
            # Crea cartella di output se non esiste
            out_path = Path(output_dir)
            if not out_path.exists():
                out_path.mkdir(parents=True)
                self._log(f"Cartella di output creata: {out_path}", "INFO")

            # file_act corrente (può essere sostituito dopo Sgamatore)
            current_file_act = file_act
            sgamatore_done = False

            for entry in self.tools:
                name = entry["script_name"]
                var = self.script_vars.get(name)

                # Salta script disabilitati (per checkbox)
                if var is None or not var.get():
                    self._log(f"Script '{name}' disabilitato, skip.", "INFO")
                    continue

                # Sostituisce file_act con _transl dopo che Sgamatore ha girato
                if use_transl and sgamatore_done and name.lower() != SGAMATORE_NAME:
                    stem = Path(file_act).stem
                    transl_path = out_path / f"{stem}_transl.xlsx"
                    if transl_path.exists():
                        current_file_act = str(transl_path)
                        self._log(f"use_transl_as_act: file_act sostituito con {transl_path}", "INFO")
                    else:
                        self._log(
                            f"ATTENZIONE: file _transl non trovato ({transl_path}). "
                            "Uso file_act originale.",
                            "WARNING"
                        )

                # Determina le opzioni da passare allo script
                script_cfg = config_body.get(name)
                if isinstance(script_cfg, dict) and "opt_override" in script_cfg:
                    # opt_override sostituisce completamente script_opt
                    extra_opts = script_cfg["opt_override"].split()
                else:
                    raw_opt = entry.get("script_opt") or ""
                    extra_opts = raw_opt.split() if raw_opt else []

                # Costruisce il comando
                script_path = str(Path(__file__).parent / entry["script_path"])
                cmd = [sys.executable, script_path,
                       "--file_act", current_file_act,
                       "--year", anno,
                       "--out", output_dir]

                # --file_bud è facoltativo: omesso se File OLD non selezionato
                if file_old:
                    cmd += ["--file_bud", file_old]

                # Opzioni specifiche dello script
                cmd += extra_opts

                self._log(f"Avvio script '{name}'...", "INFO")
                self._log(f"  Comando: {' '.join(cmd)}", "INFO")

                try:
                    # Esegue lo script catturando stdout in tempo reale
                    process = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,  # unifica stderr in stdout
                        text=True,
                        encoding="utf-8",
                        errors="replace"
                    )

                    # Legge stdout riga per riga e lo riversa nel log
                    for line in process.stdout:
                        self._log(line.rstrip(), "INFO")

                    process.wait()

                    if process.returncode != 0:
                        raise subprocess.CalledProcessError(process.returncode, cmd)

                    self._log(f"Script '{name}' completato (returncode=0).", "INFO")

                    # Segna Sgamatore come completato per la logica use_transl
                    if name.lower() == SGAMATORE_NAME:
                        sgamatore_done = True

                except subprocess.CalledProcessError as e:
                    self._log(
                        f"Script '{name}' terminato con errore (returncode={e.returncode}). "
                        "Pipeline interrotta.",
                        "ERROR"
                    )
                    return  # Interrompe la pipeline, ma il launcher resta aperto

        except Exception as e:
            self._log(f"Errore inatteso durante l'esecuzione: {e}", "ERROR")

        finally:
            # Riabilita Start in ogni caso (successo o errore)
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            self._log("Pipeline terminata.", "INFO")

            # Apri cartella di output se l'opzione è attiva
            if self.var_open_folder.get():
                self._open_output_folder(output_dir)


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    # Percorsi dei file di configurazione (stessa cartella dello script)
    base_dir = Path(__file__).parent
    tools_path = base_dir / "tools.yml"
    config_path = base_dir / "config.yml"

    # Carica tools.yml — esce con errore se non valido
    print(f"[INFO] Lettura tools.yml da {tools_path}")
    tools = load_tools(tools_path)
    print(f"[INFO] Caricati {len(tools)} script.")

    # Verifica esistenza script (già fatta in load_tools, ma logghiamo i path)
    for entry in tools:
        script_path = base_dir / entry["script_path"]
        print(f"[INFO] Script verificato: {script_path}")

    # Carica config.yml — None se non disponibile o con errori
    print(f"[INFO] Lettura config.yml da {config_path}")
    configs = load_configs(config_path, tools)
    if configs is None:
        print("[WARNING] config.yml non disponibile o con errori: listbox vuota.")

    # Costruisce e avvia la GUI
    root = tk.Tk()
    app = LauncherApp(root, tools, configs)
    root.mainloop()


if __name__ == "__main__":
    main()