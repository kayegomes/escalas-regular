import html as html_lib
import logging
import os
import queue
import threading
import traceback
import tkinter as tk
from logging.handlers import RotatingFileHandler
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    import win32com.client as win32_client
except ImportError:  # Permite importar e testar a aplicação fora do Windows/Outlook.
    win32_client = None

try:
    from tkcalendar import DateEntry
except ImportError:  # Fallback simples para ambientes de desenvolvimento sem tkcalendar.
    class DateEntry(tk.Entry):
        def __init__(self, master=None, date_pattern="dd/mm/yyyy", **kwargs):
            super().__init__(master, **kwargs)

from app_support import (
    OutputFileLockedError,
    append_execution_history,
    is_valid_email,
    load_app_config,
    safe_filename,
)


class GeradorEscalasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Escalas - Motor de Cruzamento V3")
        self.root.geometry("1020x760")
        self.root.minsize(900, 620)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.config = load_app_config(self.base_dir)
        self._closed = False
        self._ui_queue = queue.Queue()
        self._setup_logging()
        self.root.after(100, self._drain_ui_queue)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        self.tab_gerador = ttk.Frame(self.notebook)
        self.tab_contatos = ttk.Frame(self.notebook)
        self.tab_logs = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_gerador, text="Gerador de Escalas")
        self.notebook.add(self.tab_contatos, text="Gestão de Contatos")
        self.notebook.add(self.tab_logs, text="Log de Execução")

        self._build_fluxo_gerador()
        self.setup_tab_contatos()
        self.setup_tab_logs()
        self._autofill_default_paths()

    def _setup_logging(self):
        self.logger = logging.getLogger("gerador_escalas")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        if not self.logger.handlers:
            log_name = self.config.get("outputs", {}).get("log", "gerador_escalas.log")
            log_path = os.path.join(self.base_dir, log_name)
            handler = RotatingFileHandler(
                log_path,
                maxBytes=2 * 1024 * 1024,
                backupCount=3,
                encoding="utf-8",
            )
            handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
            self.logger.addHandler(handler)

    def _on_close(self):
        self._closed = True
        try:
            for handler in self.logger.handlers:
                handler.flush()
        finally:
            self.root.destroy()

    def _ui(self, callback, *args, **kwargs):
        if self._closed:
            return
        if threading.current_thread() is threading.main_thread():
            callback(*args, **kwargs)
        else:
            self._ui_queue.put((callback, args, kwargs))

    def _drain_ui_queue(self):
        if self._closed:
            return
        while True:
            try:
                callback, args, kwargs = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback(*args, **kwargs)
            except Exception:
                self.logger.exception("Falha ao atualizar a interface")
        self.root.after(100, self._drain_ui_queue)

    def _set_progress_ui(self, value, message=""):
        self.progress_var.set(max(0, min(100, int(value))))
        if message:
            self.status_var.set(message)

    def set_progress(self, value, message=""):
        self._ui(self._set_progress_ui, value, message)

    def _set_button_state(self, button, state):
        self._ui(button.config, state=state)

    def _set_entry_value(self, entry, value):
        def update():
            entry.delete(0, tk.END)
            entry.insert(0, value)
        self._ui(update)

    def show_error(self, title, message):
        self._ui(messagebox.showerror, title, message)

    def show_warning(self, title, message):
        self._ui(messagebox.showwarning, title, message)

    def show_info(self, title, message):
        self._ui(messagebox.showinfo, title, message)

    def _row_file(self, parent, row, label, entry_width=75, label_col=0, entry_col=1, button_col=5, label_padx=0):
        tk.Label(parent, text=label).grid(row=row, column=label_col, sticky="w", padx=label_padx)
        entry = tk.Entry(parent, width=entry_width)
        if entry_width >= 60:
            entry.grid(row=row, column=entry_col, columnspan=4, padx=5, pady=2, sticky="we")
        else:
            entry.grid(row=row, column=entry_col, padx=2, pady=2)
        tk.Button(parent, text="Procurar", command=lambda: self.browse_file(entry)).grid(row=row, column=button_col)
        return entry

    def toggle_second_grades(self):
        if self.frame_grades2.winfo_ismapped():
            self.frame_grades2.grid_remove()
            self.btn_toggle_grades2.config(text="Adicionar segunda grade")
        else:
            self.frame_grades2.grid()
            self.btn_toggle_grades2.config(text="Ocultar segunda grade")

    def _build_fluxo_gerador(self):
        frame_etapa1 = tk.LabelFrame(self.tab_gerador, text="Etapa 1: Checagem 2405 vs Grades", padx=10, pady=10)
        frame_etapa1.pack(fill="x", padx=10, pady=5)

        self.entry_2405 = self._row_file(frame_etapa1, 0, "Relatório 2405 Bruto:")
        self.entry_2468 = self._row_file(frame_etapa1, 1, "Relatório 2468 Bruto:")
        tk.Label(frame_etapa1, text="Grade SporTV 1:").grid(row=2, column=0, sticky="w")
        self.entry_sportv = tk.Entry(frame_etapa1, width=75)
        self.entry_sportv.grid(row=2, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(frame_etapa1, text="Procurar", command=lambda: self.browse_file(self.entry_sportv)).grid(row=2, column=5)
        tk.Label(frame_etapa1, text="Grade Premiere 1:").grid(row=3, column=0, sticky="w")
        self.entry_premiere = tk.Entry(frame_etapa1, width=75)
        self.entry_premiere.grid(row=3, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(frame_etapa1, text="Procurar", command=lambda: self.browse_file(self.entry_premiere)).grid(row=3, column=5)
        tk.Label(frame_etapa1, text="Grade Combate 1:").grid(row=4, column=0, sticky="w")
        self.entry_combate = tk.Entry(frame_etapa1, width=75)
        self.entry_combate.grid(row=4, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(frame_etapa1, text="Procurar", command=lambda: self.browse_file(self.entry_combate)).grid(row=4, column=5)

        self.btn_toggle_grades2 = tk.Button(
            frame_etapa1,
            text="Adicionar segunda grade",
            command=self.toggle_second_grades,
        )
        self.btn_toggle_grades2.grid(row=5, column=0, columnspan=6, pady=(6, 4))

        self.frame_grades2 = tk.LabelFrame(frame_etapa1, text="Grades 2 opcionais", padx=8, pady=8)
        self.frame_grades2.grid(row=6, column=0, columnspan=6, sticky="we", pady=(0, 6))
        self.frame_grades2.grid_remove()

        tk.Label(self.frame_grades2, text="Grade SporTV 2:").grid(row=0, column=0, sticky="w")
        self.entry_sportv_2 = tk.Entry(self.frame_grades2, width=75)
        self.entry_sportv_2.grid(row=0, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(self.frame_grades2, text="Procurar", command=lambda: self.browse_file(self.entry_sportv_2)).grid(row=0, column=5)

        tk.Label(self.frame_grades2, text="Grade Premiere 2:").grid(row=1, column=0, sticky="w")
        self.entry_premiere_2 = tk.Entry(self.frame_grades2, width=75)
        self.entry_premiere_2.grid(row=1, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(self.frame_grades2, text="Procurar", command=lambda: self.browse_file(self.entry_premiere_2)).grid(row=1, column=5)

        tk.Label(self.frame_grades2, text="Grade Combate 2:").grid(row=2, column=0, sticky="w")
        self.entry_combate_2 = tk.Entry(self.frame_grades2, width=75)
        self.entry_combate_2.grid(row=2, column=1, columnspan=4, padx=5, pady=2, sticky="we")
        tk.Button(self.frame_grades2, text="Procurar", command=lambda: self.browse_file(self.entry_combate_2)).grid(row=2, column=5)

        self.btn_etapa1 = tk.Button(
            frame_etapa1,
            text="Executar Etapa 1 (Checar 2405)",
            command=self.start_etapa0,
            bg="blue",
            fg="white",
            font=("Arial", 10, "bold"),
        )
        self.btn_etapa1.grid(row=7, column=0, columnspan=6, pady=10)

        frame_etapa2 = tk.LabelFrame(self.tab_gerador, text="Etapa 2: Motor de Cruzamento (Gerar Check Pre Envio)", padx=10, pady=10)
        frame_etapa2.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_etapa2, text="Use o Relatório 2468 carregado na Etapa 1.").grid(row=0, column=0, sticky="w")
        self.btn_etapa2 = tk.Button(
            frame_etapa2,
            text="Executar Etapa 2 (Gerar Excel)",
            command=self.start_etapa1,
            bg="green",
            fg="white",
            font=("Arial", 10, "bold"),
        )
        self.btn_etapa2.grid(row=1, column=0, sticky="w", pady=10)

        frame_etapa3 = tk.LabelFrame(self.tab_gerador, text="Etapa 3: Gerador de HTML e Disparo", padx=10, pady=10)
        frame_etapa3.pack(fill="x", padx=10, pady=5)

        self.entry_check = self._row_file(frame_etapa3, 0, "Check Pre Envio (Revisado):", entry_width=60)
        self.entry_contacts = self._row_file(frame_etapa3, 1, "Planilha de Contatos:", entry_width=60)

        tk.Label(frame_etapa3, text="Filtrar Período (DD/MM/YYYY):").grid(row=2, column=0, sticky="w", pady=5)
        frame_datas = tk.Frame(frame_etapa3)
        frame_datas.grid(row=2, column=1, sticky="w", columnspan=3)
        tk.Label(frame_datas, text="De:").pack(side=tk.LEFT)
        self.entry_data_inicio = DateEntry(frame_datas, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern="dd/mm/yyyy")
        self.entry_data_inicio.delete(0, "end")
        self.entry_data_inicio.pack(side=tk.LEFT, padx=5)
        tk.Label(frame_datas, text="Até:").pack(side=tk.LEFT)
        self.entry_data_fim = DateEntry(frame_datas, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern="dd/mm/yyyy")
        self.entry_data_fim.delete(0, "end")
        self.entry_data_fim.pack(side=tk.LEFT, padx=5)

        tk.Label(frame_etapa3, text="Grupos para Enviar:").grid(row=3, column=0, sticky="nw", pady=5)
        frame_grupos = tk.Frame(frame_etapa3)
        frame_grupos.grid(row=3, column=1, sticky="w", columnspan=3)

        self.var_narradores = tk.BooleanVar(value=True)
        self.var_coment_futebol = tk.BooleanVar(value=True)
        self.var_coment_outros = tk.BooleanVar(value=True)
        self.var_coment_arbitragem = tk.BooleanVar(value=True)
        self.var_colaboradores = tk.BooleanVar(value=True)
        self.var_outros = tk.BooleanVar(value=True)

        tk.Checkbutton(frame_grupos, text="Narradores", variable=self.var_narradores).grid(row=0, column=0, sticky="w")
        tk.Checkbutton(frame_grupos, text="Comentaristas Futebol", variable=self.var_coment_futebol).grid(row=0, column=1, sticky="w")
        tk.Checkbutton(frame_grupos, text="Coment. Multimodalidade", variable=self.var_coment_outros).grid(row=0, column=2, sticky="w")
        tk.Checkbutton(frame_grupos, text="Coment. Arbitragem", variable=self.var_coment_arbitragem).grid(row=1, column=0, sticky="w")
        tk.Checkbutton(frame_grupos, text="Colaboradores", variable=self.var_colaboradores).grid(row=1, column=1, sticky="w")
        tk.Checkbutton(frame_grupos, text="Desconhecidos", variable=self.var_outros).grid(row=1, column=2, sticky="w")

        tk.Label(frame_etapa3, text="Opção de Envio:").grid(row=4, column=0, sticky="w")
        self.envio_var = tk.StringVar(value="gerar")
        frame_envio = tk.Frame(frame_etapa3)
        frame_envio.grid(row=4, column=1, sticky="w", columnspan=3)
        tk.Radiobutton(frame_envio, text="Somente Gerar HTML", variable=self.envio_var, value="gerar").pack(side=tk.LEFT)
        tk.Radiobutton(frame_envio, text="Enviar Teste (Para mim)", variable=self.envio_var, value="teste").pack(side=tk.LEFT)
        tk.Radiobutton(frame_envio, text="Disparo Oficial (criar rascunhos)", variable=self.envio_var, value="oficial").pack(side=tk.LEFT)

        tk.Label(frame_etapa3, text="E-mail para teste:").grid(row=5, column=0, sticky="w")
        self.entry_email_teste = tk.Entry(frame_etapa3, width=45)
        self.entry_email_teste.grid(row=5, column=1, columnspan=3, sticky="w", padx=5, pady=2)
        tk.Label(frame_etapa3, text="Obrigatório somente no modo de teste.", fg="#666666").grid(row=5, column=4, sticky="w")

        self.excecoes_envio = []
        tk.Button(
            frame_etapa3,
            text="Gerenciar Exceções (Ignorar Profissionais)",
            command=self.abrir_janela_excecoes,
        ).grid(row=6, column=0, columnspan=4, pady=5)

        self.btn_etapa3 = tk.Button(
            frame_etapa3,
            text="Executar Etapa 3 (Gerar HTML)",
            command=self.start_etapa2,
            bg="red",
            fg="white",
            font=("Arial", 10, "bold"),
        )
        self.btn_etapa3.grid(row=7, column=0, columnspan=4, pady=10)

        self.status_var = tk.StringVar(value="Pronto para iniciar")
        self.progress_var = tk.IntVar(value=0)
        frame_progress = tk.Frame(self.tab_gerador)
        frame_progress.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(frame_progress, textvariable=self.status_var).pack(side=tk.LEFT, padx=(0, 10))
        self.progress_bar = ttk.Progressbar(
            frame_progress,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.progress_var,
        )
        self.progress_bar.pack(side=tk.LEFT, fill="x", expand=True)

    def setup_tab_logs(self):
        """Cria uma área dedicada para leitura e gerenciamento do log de execução."""
        header = ttk.Frame(self.tab_logs, padding=(12, 10, 12, 4))
        header.pack(fill="x")
        ttk.Label(header, text="Log de Execução", font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        ttk.Label(
            header,
            text="Acompanhe o processamento completo, inclusive em telas menores.",
            foreground="#555555",
        ).pack(side=tk.LEFT, padx=(12, 0))

        actions = ttk.Frame(self.tab_logs, padding=(12, 0, 12, 8))
        actions.pack(fill="x")
        ttk.Button(actions, text="Copiar Log", command=self.copy_log).pack(side=tk.LEFT)
        ttk.Button(actions, text="Limpar Log", command=self.clear_log).pack(side=tk.LEFT, padx=6)
        ttk.Button(actions, text="Salvar Log", command=self.save_log).pack(side=tk.LEFT)

        frame_log = ttk.Frame(self.tab_logs, padding=(12, 0, 12, 12))
        frame_log.pack(fill="both", expand=True)
        frame_log.rowconfigure(0, weight=1)
        frame_log.columnconfigure(0, weight=1)

        self.txt_log = tk.Text(
            frame_log,
            wrap="none",
            font=("Consolas", 10),
            undo=False,
            state="normal",
        )
        self.txt_log.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x = ttk.Scrollbar(frame_log, orient="horizontal", command=self.txt_log.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.txt_log.config(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    def clear_log(self):
        self.txt_log.delete("1.0", tk.END)
        self.logger.info("Log visual limpo pelo usuário.")

    def copy_log(self):
        content = self.txt_log.get("1.0", tk.END).strip()
        if not content:
            self.show_info("Log de Execução", "Não há mensagens para copiar.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.root.update_idletasks()
        self.show_info("Log de Execução", "Log copiado para a área de transferência.")

    def _autofill_default_paths(self):
        """Preenche arquivos conhecidos apenas quando estão na pasta do programa."""
        defaults = self.config.get("defaults", {})
        output_names = self.config.get("outputs", {})
        candidates = {
            self.entry_2405: [defaults.get("report_2405", "Check pre envio-Macro Colunas - 15052026.xlsm")],
            self.entry_2468: [defaults.get("report_2468", "(2468) Esporte - Atividades de Equipe – Sub-Atividades_v2_ (9).xlsx"), output_names.get("etapa_1", "Check_2405_Gerado.xlsx"), output_names.get("etapa_2", "Check_Pre_Envio_Gerado.xlsx")],
            self.entry_contacts: [defaults.get("contacts", output_names.get("contacts", "contatos_nova_versao.xlsx"))],
        }
        for entry, names in candidates.items():
            if entry.get().strip():
                continue
            for name in names:
                path = os.path.join(self.base_dir, name)
                if os.path.exists(path):
                    entry.insert(0, path)
                    break

    def setup_tab_contatos(self):
        frame_tree = tk.Frame(self.tab_contatos)
        frame_tree.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("Grupo", "Nome", "Email")
        self.tree_contatos = ttk.Treeview(frame_tree, columns=columns, show="headings")
        self.tree_contatos.heading("Grupo", text="Grupo")
        self.tree_contatos.heading("Nome", text="Nome")
        self.tree_contatos.heading("Email", text="Email")
        self.tree_contatos.column("Grupo", width=180)
        self.tree_contatos.column("Nome", width=280)
        self.tree_contatos.column("Email", width=380)

        scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=self.tree_contatos.yview)
        self.tree_contatos.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_contatos.pack(side=tk.LEFT, fill="both", expand=True)
        self.tree_contatos.bind("<<TreeviewSelect>>", self.on_contato_select)

        frame_controls = tk.LabelFrame(self.tab_contatos, text="Adicionar / Editar Contato", padx=10, pady=10)
        frame_controls.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_controls, text="Nome:").grid(row=0, column=0, sticky="e")
        self.entry_nome_contato = tk.Entry(frame_controls, width=30)
        self.entry_nome_contato.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_controls, text="Email:").grid(row=0, column=2, sticky="e")
        self.entry_email_contato = tk.Entry(frame_controls, width=30)
        self.entry_email_contato.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(frame_controls, text="Grupo:").grid(row=1, column=0, sticky="e")
        self.combo_grupo = ttk.Combobox(
            frame_controls,
            values=[
                "Narrador",
                "Comentarista Futebol",
                "Comentaristas (outros)",
                "Comentaristas Arbitragem",
                "Colaboradores",
                "Outros / Desconhecidos",
            ],
            state="readonly",
            width=27,
        )
        self.combo_grupo.grid(row=1, column=1, padx=5, pady=5)

        frame_btns = tk.Frame(frame_controls)
        frame_btns.grid(row=1, column=2, columnspan=2, pady=5)
        tk.Button(frame_btns, text="Adicionar", command=self.add_contato).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btns, text="Atualizar Selecionado", command=self.update_contato).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btns, text="Excluir", command=self.delete_contato).pack(side=tk.LEFT, padx=5)

        tk.Button(
            self.tab_contatos,
            text="Salvar Planilha de Contatos",
            command=self.save_contatos,
            bg="blue",
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(pady=10)

        self.load_contatos()

    def load_contatos(self):
        for item in self.tree_contatos.get_children():
            self.tree_contatos.delete(item)

        contacts_name = self.config.get("outputs", {}).get("contacts", "contatos_nova_versao.xlsx")
        contatos_file = os.path.join(self.base_dir, contacts_name)
        if not os.path.exists(contatos_file):
            return

        try:
            df = pd.read_excel(contatos_file, sheet_name="Lista e-mails")
            if {"Grupo", "Nome", "Email"}.issubset(df.columns):
                for _, r in df.iterrows():
                    self.tree_contatos.insert("", tk.END, values=(str(r["Grupo"]), str(r["Nome"]), str(r.get("Email", ""))))
        except Exception as e:
            self.logger.exception("Erro ao carregar contatos")
            messagebox.showerror("Erro", f"Erro ao carregar contatos: {e}")

    def on_contato_select(self, event):
        selected = self.tree_contatos.selection()
        if not selected:
            return

        item = self.tree_contatos.item(selected[0])
        values = item.get("values", [])
        if len(values) >= 3:
            self.combo_grupo.set(values[0])
            self.entry_nome_contato.delete(0, tk.END)
            self.entry_nome_contato.insert(0, values[1])
            self.entry_email_contato.delete(0, tk.END)
            self.entry_email_contato.insert(0, values[2])

    def add_contato(self):
        nome = self.entry_nome_contato.get().strip()
        email = self.entry_email_contato.get().strip()
        grupo = self.combo_grupo.get()
        if not nome or not grupo:
            messagebox.showwarning("Aviso", "Nome e Grupo são obrigatórios!")
            return
        self.tree_contatos.insert("", tk.END, values=(grupo, nome, email))
        self.entry_nome_contato.delete(0, tk.END)
        self.entry_email_contato.delete(0, tk.END)

    def update_contato(self):
        selected = self.tree_contatos.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um contato na lista para atualizar!")
            return
        nome = self.entry_nome_contato.get().strip()
        email = self.entry_email_contato.get().strip()
        grupo = self.combo_grupo.get()
        if not nome or not grupo:
            messagebox.showwarning("Aviso", "Nome e Grupo são obrigatórios!")
            return
        self.tree_contatos.item(selected[0], values=(grupo, nome, email))

    def delete_contato(self):
        selected = self.tree_contatos.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um contato para excluir!")
            return
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir o contato selecionado?"):
            self.tree_contatos.delete(selected[0])

    def save_contatos(self):
        contacts_name = self.config.get("outputs", {}).get("contacts", "contatos_nova_versao.xlsx")
        contatos_file = os.path.join(self.base_dir, contacts_name)
        data = []
        for item in self.tree_contatos.get_children():
            v = self.tree_contatos.item(item)["values"]
            if len(v) >= 3:
                data.append({"Grupo": v[0], "Nome": v[1], "Email": v[2]})

        try:
            pd.DataFrame(data).to_excel(contatos_file, index=False, sheet_name="Lista e-mails")
            messagebox.showinfo("Sucesso", "Planilha de contatos atualizada com sucesso!")
        except PermissionError:
            messagebox.showerror(
                "Arquivo bloqueado",
                "Feche a planilha de contatos no Excel e tente salvar novamente.",
            )
        except Exception as e:
            self.logger.exception("Erro ao salvar contatos")
            messagebox.showerror("Erro", f"Erro ao salvar planilha: {e}")

    def log(self, message, level=logging.INFO):
        text = str(message)
        self.logger.log(level, text)

        def append_to_text():
            self.txt_log.insert(tk.END, text + "\n")
            self.txt_log.see(tk.END)
            self.root.update_idletasks()

        self._ui(append_to_text)

    def save_log(self):
        log_content = self.txt_log.get("1.0", tk.END)
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(log_content)
            messagebox.showinfo("Sucesso", f"Log salvo em: {file_path}")
        except OSError as e:
            self.logger.exception("Erro ao exportar log")
            messagebox.showerror("Erro", f"Não foi possível salvar o log: {e}")

    def abrir_janela_excecoes(self):
        check_path = self.entry_check.get()
        if not check_path or not os.path.exists(check_path):
            messagebox.showerror("Erro", "Selecione o arquivo Check Pre Envio primeiro para carregar os nomes.")
            return

        try:
            df = pd.read_excel(check_path)
            nome_col = "Nome " if "Nome " in df.columns else "Nome"
            if nome_col not in df.columns:
                messagebox.showerror("Erro", f"Coluna de Nome ({nome_col}) não encontrada na planilha.")
                return
            profissionais = sorted(
                str(p).strip()
                for p in df[nome_col].dropna().unique()
                if str(p).strip() not in ["nan", "-"]
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler planilha: {e}")
            return

        top = tk.Toplevel(self.root)
        top.title("Selecionar Exceções")
        top.geometry("400x500")
        tk.Label(top, text="Selecione os profissionais que NÃO devem receber escala:").pack(pady=10)

        frame_list = tk.Frame(top)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)
        scrollbar = tk.Scrollbar(frame_list)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame_list, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, width=50)
        for i, p in enumerate(profissionais):
            listbox.insert(tk.END, p)
            if p in self.excecoes_envio:
                listbox.selection_set(i)
        listbox.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        def salvar_excecoes():
            self.excecoes_envio = [listbox.get(i) for i in listbox.curselection()]
            self.log(f"{len(self.excecoes_envio)} exceções salvas. Estes profissionais serão ignorados na geração.")
            top.destroy()

        tk.Button(top, text="Salvar Exceções", command=salvar_excecoes, bg="orange", font=("Arial", 10, "bold")).pack(pady=10)

    def browse_file(self, entry_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

    def start_etapa0(self):
        path_2405 = self.entry_2405.get().strip()
        if not path_2405 or not os.path.exists(path_2405):
            messagebox.showerror("Erro", "Selecione o relatório 2405.")
            return

        self.btn_etapa1.config(state=tk.DISABLED)
        self.set_progress(0, "Preparando Etapa 1...")
        self.log("Iniciando checagem 2405 vs grades (Etapa 1)...")
        args = (
            path_2405,
            self.entry_sportv.get().strip(),
            self.entry_sportv_2.get().strip(),
            self.entry_premiere.get().strip(),
            self.entry_premiere_2.get().strip(),
            self.entry_combate.get().strip(),
            self.entry_combate_2.get().strip(),
        )
        threading.Thread(target=self.process_etapa0, args=args, daemon=True).start()

    def process_etapa0(self, path_2405, path_sp1, path_sp2, path_pr1, path_pr2, path_co1, path_co2):
        try:
            from engine_2405 import run_etapa1_2405

            self.set_progress(10, "Lendo relatório 2405...")
            self.log("Lendo relatório 2405 e grades...")
            self.set_progress(45, "Cruzando eventos com as grades...")
            out_path = run_etapa1_2405(path_2405, path_sp1, path_sp2, path_pr1, path_pr2, path_co1, path_co2)
            self.set_progress(100, "Etapa 1 concluída")
            self.log(f"Etapa 1 concluída! Arquivo gerado: {out_path}")
        except OutputFileLockedError as e:
            self.log(str(e), logging.ERROR)
            self.show_error("Arquivo Excel bloqueado", str(e))
            append_execution_history(os.path.dirname(path_2405), "Etapa 1 - Checagem 2405 vs grades", "ERRO", "Arquivo Excel bloqueado", str(e))
        except Exception as e:
            self.logger.exception("Erro na Etapa 1")
            self.log(f"Erro na Etapa 1: {e}", logging.ERROR)
            self.show_error("Erro na Etapa 1", str(e))
            append_execution_history(os.path.dirname(path_2405), "Etapa 1 - Checagem 2405 vs grades", "ERRO", str(e), traceback.format_exc())
        finally:
            self._set_button_state(self.btn_etapa1, tk.NORMAL)

    def start_etapa1(self):
        f_2468 = self.entry_2468.get().strip()
        if not f_2468 or not os.path.exists(f_2468):
            messagebox.showerror("Erro", "Selecione o Relatório 2468.")
            return

        self.btn_etapa2.config(state=tk.DISABLED)
        self.set_progress(0, "Preparando Etapa 2...")
        self.log("Iniciando Motor de Cruzamento (Etapa 2)...")
        args = (
            f_2468,
            self.entry_sportv.get().strip(),
            self.entry_sportv_2.get().strip(),
            self.entry_premiere.get().strip(),
            self.entry_premiere_2.get().strip(),
            self.entry_combate.get().strip(),
            self.entry_combate_2.get().strip(),
        )
        threading.Thread(target=self.process_etapa1, args=args, daemon=True).start()

    def process_etapa1(self, path_2468, path_sp1, path_sp2, path_pr1, path_pr2, path_co1, path_co2):
        try:
            from engine_cross import run_etapa1

            self.set_progress(10, "Lendo relatório 2468...")
            self.log("Lendo relatórios e grades... (Isso pode demorar alguns segundos)")
            self.set_progress(45, "Executando cruzamento e validação...")
            out_path = run_etapa1(path_2468, path_sp1, path_sp2, path_pr1, path_pr2, path_co1, path_co2)
            self.set_progress(100, "Etapa 2 concluída")
            self.log(f"Etapa 2 concluída! Arquivo gerado: {out_path}")
            self._set_entry_value(self.entry_check, out_path)
        except OutputFileLockedError as e:
            self.log(str(e), logging.ERROR)
            self.show_error("Arquivo Excel bloqueado", str(e))
            append_execution_history(os.path.dirname(path_2468), "Etapa 2 - Cruzamento 2468 vs grades", "ERRO", "Arquivo Excel bloqueado", str(e))
        except Exception as e:
            self.logger.exception("Erro na Etapa 2")
            self.log(f"Erro na Etapa 2: {e}", logging.ERROR)
            self.show_error("Erro na Etapa 2", str(e))
            append_execution_history(os.path.dirname(path_2468), "Etapa 2 - Cruzamento 2468 vs grades", "ERRO", str(e), traceback.format_exc())
        finally:
            self._set_button_state(self.btn_etapa2, tk.NORMAL)

    def start_etapa2(self):
        check_path = self.entry_check.get().strip()
        if not check_path or not os.path.exists(check_path):
            messagebox.showerror("Erro", "Selecione o arquivo Check Pre Envio.")
            return

        selected_groups = []
        if self.var_narradores.get():
            selected_groups.append("Narrador")
        if self.var_coment_futebol.get():
            selected_groups.append("Comentarista Futebol")
        if self.var_coment_outros.get():
            selected_groups.append("Comentaristas (outros)")
        if self.var_coment_arbitragem.get():
            selected_groups.append("Comentaristas Arbitragem")
        if self.var_colaboradores.get():
            selected_groups.append("Colaboradores")
        if self.var_outros.get():
            selected_groups.append("Outros / Desconhecidos")

        data_inicio = self.entry_data_inicio.get().strip()
        data_fim = self.entry_data_fim.get().strip()
        if data_inicio or data_fim:
            if not (data_inicio and data_fim):
                messagebox.showerror("Filtro de período", "Preencha as duas datas ou deixe o período vazio.")
                return
            try:
                dt_inicio = pd.to_datetime(data_inicio, format="%d/%m/%Y")
                dt_fim = pd.to_datetime(data_fim, format="%d/%m/%Y")
                if dt_inicio > dt_fim:
                    messagebox.showerror("Filtro de período", "A data inicial não pode ser posterior à data final.")
                    return
            except ValueError:
                messagebox.showerror("Filtro de período", "Use o formato DD/MM/YYYY.")
                return

        modo_envio = self.envio_var.get()
        email_teste = self.entry_email_teste.get().strip()
        if modo_envio == "teste" and not is_valid_email(email_teste):
            messagebox.showerror("E-mail de teste", "Informe um e-mail válido para receber o rascunho de teste.")
            return

        self.btn_etapa3.config(state=tk.DISABLED)
        self.set_progress(0, "Preparando Etapa 3...")
        self.log("Iniciando geração de HTMLs (Etapa 3)...")
        args = (
            check_path,
            self.entry_contacts.get().strip(),
            selected_groups,
            list(self.excecoes_envio),
            data_inicio,
            data_fim,
            modo_envio,
            email_teste,
        )
        threading.Thread(target=self.process_etapa2, args=args, daemon=True).start()

    def process_etapa2(
        self,
        check_path,
        contacts_path,
        selected_groups,
        excecoes_envio,
        data_inicio,
        data_fim,
        modo_envio,
        email_teste,
    ):
        try:
            self.set_progress(10, "Lendo arquivo consolidado...")
            self.log("Lendo arquivo consolidado...")
            df = pd.read_excel(check_path)

            nome_col = "Nome " if "Nome " in df.columns else "Nome"
            if nome_col not in df.columns:
                self.log(f"ERRO: Coluna '{nome_col}' não encontrada na planilha.")
                return

            df = df.dropna(subset=[nome_col])
            df = df[df[nome_col] != ""]

            if "Data" in df.columns:
                df["Data_obj"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
                if data_inicio and data_fim:
                    try:
                        dt_inicio = pd.to_datetime(data_inicio, format="%d/%m/%Y")
                        dt_fim = pd.to_datetime(data_fim, format="%d/%m/%Y")
                        mask = (df["Data_obj"] >= dt_inicio) & (df["Data_obj"] <= dt_fim)
                        df = df[mask]
                        self.log(f"Filtro aplicado: mostrando eventos de {data_inicio} até {data_fim}.")
                    except Exception as e:
                        self.log(f"Erro ao processar datas de filtro. Verifique o formato DD/MM/YYYY. Detalhe: {e}")

                dias_pt = {
                    "Monday": "Segunda-feira",
                    "Tuesday": "Terça-feira",
                    "Wednesday": "Quarta-feira",
                    "Thursday": "Quinta-feira",
                    "Friday": "Sexta-feira",
                    "Saturday": "Sábado",
                    "Sunday": "Domingo",
                }
                df["Dia"] = df["Data_obj"].dt.strftime("%A").map(dias_pt)
                df["Data"] = df["Data_obj"].dt.strftime("%d/%m/%Y")

            if df.empty:
                self.set_progress(100, "Nenhum evento no período selecionado")
                self.log("Nenhum evento encontrado no período especificado.")
                return

            df = df.fillna("-")
            self.set_progress(25, "Aplicando filtros e preparando contatos...")

            contacts_path = contacts_path.strip() if contacts_path else ""
            if not contacts_path or not os.path.exists(contacts_path):
                self.log("ERRO: Planilha de contatos é obrigatória para filtrar por grupos.", logging.ERROR)
                self.show_error("Planilha de contatos", "Selecione uma planilha de contatos válida.")
                return

            self.set_progress(35, "Lendo grupos da planilha de contatos...")
            self.log("Lendo grupos da planilha de contatos...")
            contacts_dict = {}
            name_to_group = {}
            try:
                df_contacts = pd.read_excel(contacts_path, sheet_name="Lista e-mails")
                if {"Grupo", "Nome"}.issubset(df_contacts.columns):
                    for _, r in df_contacts.iterrows():
                        n = str(r["Nome"]).strip().lower()
                        g = str(r["Grupo"]).strip()
                        e = str(r.get("Email", "")).strip()
                        if n and n != "nan":
                            name_to_group[n] = g
                            if e and e != "nan":
                                contacts_dict[n] = e
                else:
                    message = "Formato da aba 'Lista e-mails' inválido. Selecione a nova versão de contatos gerada."
                    self.log(f"ERRO: {message}", logging.ERROR)
                    self.show_error("Planilha de contatos", message)
                    return
            except Exception as e:
                self.logger.exception("Erro ao ler contatos na Etapa 3")
                self.log(f"Erro ao ler contatos: {e}", logging.ERROR)
                self.show_error("Planilha de contatos", str(e))
                return

            import difflib

            profissionais_a_gerar = []
            for prof in df[nome_col].unique():
                prof_lower = str(prof).lower()
                grupo_do_prof = name_to_group.get(prof_lower)
                if not grupo_do_prof:
                    matches = difflib.get_close_matches(prof_lower, name_to_group.keys(), n=1, cutoff=0.8)
                    if matches:
                        grupo_do_prof = name_to_group[matches[0]]
                if not grupo_do_prof:
                    grupo_do_prof = "Outros / Desconhecidos"
                if grupo_do_prof in selected_groups and prof not in excecoes_envio:
                    profissionais_a_gerar.append(prof)

            if not profissionais_a_gerar:
                self.log("Nenhum profissional restou após os filtros (verifique grupos e exceções).")
                return

            df = df[df[nome_col].isin(profissionais_a_gerar)]
            profissionais = df[nome_col].unique()
            self.set_progress(55, f"Gerando HTML para {len(profissionais)} profissionais...")
            self.log(f"Gerando HTML para {len(profissionais)} profissionais...")

            html_dir_name = self.config.get("outputs", {}).get("html_dir", "escalas_geradas_html")
            output_dir = os.path.join(os.path.dirname(check_path), html_dir_name)
            os.makedirs(output_dir, exist_ok=True)

            total_profissionais = len(profissionais)
            for index, prof in enumerate(profissionais, start=1):
                df_prof = df[df[nome_col] == prof]
                if not df_prof.empty:
                    self.gerar_html(prof, df_prof, output_dir)
                progress = 55 + int((index / total_profissionais) * 30)
                self.set_progress(progress, f"HTML {index}/{total_profissionais}: {prof}")

            self.log(f"Geração concluída! HTMLs salvos em: {output_dir}")
            self.set_progress(88, "Preparando rascunhos no Outlook...")

            if modo_envio == "gerar":
                self.log("Modo: 'Somente Gerar HTML'. Nenhum e-mail foi preparado.")
            elif modo_envio == "teste":
                self.log(f"Modo de teste: preparando rascunho para {email_teste}...")
                self.enviar_emails(output_dir, contacts_dict, teste=True, teste_destinatario=email_teste)
            elif modo_envio == "oficial":
                self.log("Modo oficial: preparando rascunhos para a base. Nenhum e-mail será enviado automaticamente.")
                self.enviar_emails(output_dir, contacts_dict, teste=False)
            self.set_progress(100, "Etapa 3 concluída")
        except Exception as e:
            self.logger.exception("Erro na Etapa 3")
            self.log(f"Erro na Etapa 3: {e}", logging.ERROR)
            self.show_error("Erro na Etapa 3", str(e))
            append_execution_history(os.path.dirname(check_path), "Etapa 3 - Geração e rascunhos", "ERRO", str(e), traceback.format_exc())
        finally:
            self._set_button_state(self.btn_etapa3, tk.NORMAL)


    def gerar_html(self, nome, df, output_dir):
        html_template = """
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; background: #f4f6f8; margin: 0; padding: 12px; color: #222; }}
                .container {{ max-width: 1400px; margin: 0 auto; background: #fff; padding: 22px; border-radius: 10px; box-shadow: 0 1px 8px rgba(0,0,0,.10); }}
                .greeting {{ text-align: center; border-bottom: 1px solid #d9d9d9; padding-bottom: 18px; margin-bottom: 28px; }}
                .greeting h1 {{ margin: 0 0 4px; font-size: 24px; color: #222; }}
                .greeting p {{ margin: 3px 0; color: #666; font-size: 16px; }}
                .contact-box {{ background: #f7f9fc; border-left: 4px solid #1683e8; padding: 18px 14px; margin-bottom: 24px; font-size: 16px; }}
                .contact-box strong {{ display: block; margin-bottom: 12px; }}
                .contact-box p {{ margin: 0; }}
                h3 {{ margin: 0 0 12px; color: #222; font-size: 17px; }}
                table {{ border-collapse: collapse; width: 100%; font-size: 12px; }}
                th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; vertical-align: top; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="greeting">
                    <h1>Oi {nome}</h1>
                    <p>Tudo bem?</p>
                    <p>Envio abaixo a escala</p>
                </div>
                <div class="contact-box">
                    <strong>Dúvidas ou problemas? É só nos procurar:</strong>
                    <p>Leticia Alvares: (21) 97951-2324 | Carlla Amara: (21) 99242-1837</p>
                </div>
                <h3>Escala consolidada</h3>
            <table>
                <tr>
                    <th>Nome</th>
                    <th>Plataforma</th>
                    <th>Data</th>
                    <th>Dia</th>
                    <th>Pré</th>
                    <th>Início</th>
                    <th>Fim</th>
                    <th>Evento/Descrição</th>
                    <th>Produto</th>
                    <th>Local</th>
                    <th>Elenco</th>
                    <th>Coordenador</th>
                    <th>Produtor</th>
                </tr>
                {rows}
            </table>
            </div>
        </body>
        </html>
        """

        def format_time(time_str):
            if not time_str or str(time_str).strip() in ["-", "nan", "NaT", ""]:
                return "-"
            t = str(time_str).strip()
            if " " in t:
                t = t.split(" ")[-1]
            parts = t.split(":")
            return f"{parts[0]}:{parts[1]}" if len(parts) >= 2 else t

        def html_value(value):
            if value is None or pd.isna(value):
                return "-"
            text = str(value).strip()
            if text in {"", "nan", "NaT", "None"}:
                return "-"
            return html_lib.escape(text)

        def elenco_value(row):
            """Monta o elenco exibido no HTML a partir das funções da escala."""
            people = []
            seen = set()
            for column in ("Elenco", "Narrador", "Comentarista", "Repórter"):
                value = row.get(column, "-")
                if value is None or pd.isna(value):
                    continue
                text = str(value).strip()
                if text in {"", "-", "nan", "NaT", "None"}:
                    continue
                for person in text.split(";"):
                    person = person.strip()
                    key = person.casefold()
                    if person and person != "-" and key not in seen:
                        people.append(person)
                        seen.add(key)
            return html_value(" ; ".join(people) if people else "-")

        rows_html = ""
        for _, row in df.iterrows():
            evento = html_value(
                row.get(
                    "Atividade/Descrição",
                    row.get("Evento/Descrição", row.get("Evento/Programa", row.get("Event Group", "-"))),
                )
            )
            inicio = html_value(format_time(row.get("Início", "-")))
            fim = html_value(format_time(row.get("Fim", "-")))
            pre = html_value(format_time(row.get("Pré", "-")))

            data_val = row.get("Data")
            data_formatada = "-"
            if pd.notnull(data_val) and str(data_val).strip() not in ["-", ""]:
                try:
                    dt = pd.to_datetime(data_val, dayfirst=True)
                    base_date = dt.strftime("%d/%m/%Y")
                    first_time = pre if pre != "-" else inicio
                    if first_time != "-" and first_time[:2] in ["00", "01", "02", "03", "04", "05"]:
                        next_day = dt + pd.Timedelta(days=1)
                        data_formatada = f"{base_date} para {next_day.strftime('%d/%m/%Y')}"
                    else:
                        data_formatada = base_date
                except Exception:
                    data_formatada = str(data_val).split(" ")[0]

            rows_html += "<tr>"
            rows_html += f"<td>{html_value(row.get('Nome ', row.get('Nome', '-')))}</td>"
            rows_html += f"<td>{html_value(row.get('Plataforma', '-'))}</td>"
            rows_html += f"<td>{html_value(data_formatada)}</td>"
            rows_html += f"<td>{html_value(row.get('Dia', '-'))}</td>"
            rows_html += f"<td>{pre}</td>"
            rows_html += f"<td>{inicio}</td>"
            rows_html += f"<td>{fim}</td>"
            rows_html += f"<td>{evento}</td>"
            rows_html += f"<td>{html_value(row.get('Produto (WO/Quick Hold)', row.get('Produto (WO/Shift)', row.get('Produto', '-'))))}</td>"
            rows_html += f"<td>{html_value(row.get('Local de Gravação', row.get('Local Narração', row.get('Local', '-'))))}</td>"
            rows_html += f"<td>{elenco_value(row)}</td>"
            rows_html += f"<td>{html_value(row.get('Coordenador', '-'))}</td>"
            rows_html += f"<td>{html_value(row.get('Produtor', '-'))}</td>"
            rows_html += "</tr>"

        file_name = f"escala_{safe_filename(nome)}.html"
        file_path = os.path.join(output_dir, file_name)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(html_template.format(nome=html_lib.escape(str(nome)), rows=rows_html))

    def enviar_emails(self, html_dir, contacts, teste=False, teste_destinatario=""):
        self.log("Conectando ao Outlook e preparando rascunhos...")
        if win32_client is None:
            message = "A integração com Outlook só está disponível no Windows com pywin32 instalado."
            self.log(message, logging.ERROR)
            self.show_error("Outlook indisponível", message)
            return {"encontrados": 0, "rascunhos": 0, "ignorados": 0, "erros": 1}

        try:
            outlook = win32_client.Dispatch("outlook.application")
        except Exception as e:
            message = f"Erro ao abrir o Outlook: {e}. Verifique se ele está instalado e configurado."
            self.log(message, logging.ERROR)
            self.show_error("Outlook indisponível", message)
            return {"encontrados": 0, "rascunhos": 0, "ignorados": 0, "erros": 1}

        import difflib

        arquivos = sorted(f for f in os.listdir(html_dir) if f.lower().endswith(".html"))
        self.log(f"Foram encontrados {len(arquivos)} arquivos HTML.")
        rascunhos = 0
        ignorados = 0
        erros = 0

        for index, arquivo in enumerate(arquivos, start=1):
            nome = arquivo.replace("escala_", "").replace(".html", "").replace("_", " ")
            email_dest = teste_destinatario if teste else contacts.get(nome.lower(), "")
            if not teste and not email_dest:
                matches = difflib.get_close_matches(nome.lower(), contacts.keys(), n=1, cutoff=0.7)
                if matches:
                    email_dest = contacts[matches[0]]

            if not is_valid_email(email_dest):
                ignorados += 1
                self.log(f"Rascunho ignorado para {nome}: contato sem e-mail válido.", logging.WARNING)
                continue

            caminho_completo = os.path.join(html_dir, arquivo)
            try:
                with open(caminho_completo, "r", encoding="utf-8") as f:
                    html_body = f.read()

                mail = outlook.CreateItem(0)
                mail.Subject = f"[TESTE] Escala - {nome}" if teste else f"Escala - {nome}"
                mail.To = email_dest
                mail.HTMLBody = html_body
                # Display cria o rascunho e deixa o envio sob conferência humana.
                mail.Display()
                rascunhos += 1
                self.log(f"Rascunho preparado para {nome} ({email_dest}).")
            except Exception as e:
                erros += 1
                self.logger.exception("Erro ao criar rascunho para %s", nome)
                self.log(f"Erro ao gerar rascunho para {nome}: {e}", logging.ERROR)

            self.set_progress(88 + int((index / max(len(arquivos), 1)) * 10), f"Rascunho {index}/{len(arquivos)}")

        resumo = f"{rascunhos} rascunhos preparados; {ignorados} ignorados; {erros} erros."
        self.log(f"Pronto! {resumo} O envio continua manual no Outlook.")
        append_execution_history(
            os.path.dirname(html_dir),
            "Etapa 3 - Geração e rascunhos",
            "SUCESSO" if erros == 0 else "CONCLUIDO_COM_ALERTAS",
            resumo,
        )
        return {"encontrados": len(arquivos), "rascunhos": rascunhos, "ignorados": ignorados, "erros": erros}


if __name__ == "__main__":
    root = tk.Tk()
    app = GeradorEscalasApp(root)
    root.mainloop()
