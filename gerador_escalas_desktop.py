import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import json
import win32com.client
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import locale
import openpyxl
from openpyxl.styles import PatternFill
from tkcalendar import DateEntry

class GeradorEscalasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Escalas - Motor de Cruzamento V2")
        self.root.geometry("800x650")
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        
        self.tab_gerador = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_gerador, text='Gerador de Escalas')
        
        self.tab_contatos = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_contatos, text='Gestão de Contatos')
        
        # Etapa 1 Frame
        frame_etapa1 = tk.LabelFrame(self.tab_gerador, text="Etapa 1: Motor de Cruzamento (Gerar Check Pre Envio)", padx=10, pady=10)
        frame_etapa1.pack(fill="x", padx=10, pady=5)
        
        # 2468 Input
        tk.Label(frame_etapa1, text="Relatório 2468 Bruto:").grid(row=0, column=0, sticky="w")
        self.entry_2468 = tk.Entry(frame_etapa1, width=60)
        self.entry_2468.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(frame_etapa1, text="Procurar...", command=lambda: self.browse_file(self.entry_2468)).grid(row=0, column=2)
        
        # SporTV Input
        tk.Label(frame_etapa1, text="Grade SporTV:").grid(row=1, column=0, sticky="w")
        self.entry_sportv = tk.Entry(frame_etapa1, width=60)
        self.entry_sportv.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(frame_etapa1, text="Procurar...", command=lambda: self.browse_file(self.entry_sportv)).grid(row=1, column=2)
        
        # Premiere Input
        tk.Label(frame_etapa1, text="Grade Premiere (PPV):").grid(row=2, column=0, sticky="w")
        self.entry_premiere = tk.Entry(frame_etapa1, width=60)
        self.entry_premiere.grid(row=2, column=1, padx=5, pady=2)
        tk.Button(frame_etapa1, text="Procurar...", command=lambda: self.browse_file(self.entry_premiere)).grid(row=2, column=2)
        
        # Combate Input
        tk.Label(frame_etapa1, text="Grade Combate:").grid(row=3, column=0, sticky="w")
        self.entry_combate = tk.Entry(frame_etapa1, width=60)
        self.entry_combate.grid(row=3, column=1, padx=5, pady=2)
        tk.Button(frame_etapa1, text="Procurar...", command=lambda: self.browse_file(self.entry_combate)).grid(row=3, column=2)
        
        self.btn_etapa1 = tk.Button(frame_etapa1, text="Executar Etapa 1 (Gerar Excel)", command=self.start_etapa1, bg="blue", fg="white", font=("Arial", 10, "bold"))
        self.btn_etapa1.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Etapa 2 Frame
        frame_etapa2 = tk.LabelFrame(self.tab_gerador, text="Etapa 2: Gerador de HTML e Disparo", padx=10, pady=10)
        frame_etapa2.pack(fill="x", padx=10, pady=5)
        
        # Check Pre Envio Input
        tk.Label(frame_etapa2, text="Check Pre Envio (Revisado):").grid(row=0, column=0, sticky="w")
        self.entry_check = tk.Entry(frame_etapa2, width=60)
        self.entry_check.grid(row=0, column=1, padx=5, pady=2, columnspan=2)
        tk.Button(frame_etapa2, text="Procurar...", command=lambda: self.browse_file(self.entry_check)).grid(row=0, column=3)
        
        # Contatos Input
        tk.Label(frame_etapa2, text="Planilha de Contatos:").grid(row=1, column=0, sticky="w")
        self.entry_contacts = tk.Entry(frame_etapa2, width=60)
        self.entry_contacts.grid(row=1, column=1, padx=5, pady=2, columnspan=2)
        tk.Button(frame_etapa2, text="Procurar...", command=lambda: self.browse_file(self.entry_contacts)).grid(row=1, column=3)
        
        # Período
        tk.Label(frame_etapa2, text="Filtrar Período (DD/MM/YYYY):").grid(row=2, column=0, sticky="w", pady=5)
        
        frame_datas = tk.Frame(frame_etapa2)
        frame_datas.grid(row=2, column=1, sticky="w", columnspan=3)
        tk.Label(frame_datas, text="De:").pack(side=tk.LEFT)
        self.entry_data_inicio = DateEntry(frame_datas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_data_inicio.delete(0, "end") # Começa vazio
        self.entry_data_inicio.pack(side=tk.LEFT, padx=5)
        
        tk.Label(frame_datas, text="Até:").pack(side=tk.LEFT)
        self.entry_data_fim = DateEntry(frame_datas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_data_fim.delete(0, "end") # Começa vazio
        self.entry_data_fim.pack(side=tk.LEFT, padx=5)
        
        # Grupos de Envio
        tk.Label(frame_etapa2, text="Grupos para Enviar:").grid(row=3, column=0, sticky="nw", pady=5)
        frame_grupos = tk.Frame(frame_etapa2)
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
        
        # Opções de Envio
        tk.Label(frame_etapa2, text="Opção de Envio:").grid(row=4, column=0, sticky="w")
        self.envio_var = tk.StringVar(value="gerar")
        frame_envio = tk.Frame(frame_etapa2)
        frame_envio.grid(row=4, column=1, sticky="w", columnspan=3)
        
        tk.Radiobutton(frame_envio, text="Somente Gerar HTML", variable=self.envio_var, value="gerar").pack(side=tk.LEFT)
        tk.Radiobutton(frame_envio, text="Enviar Teste (Para mim)", variable=self.envio_var, value="teste").pack(side=tk.LEFT)
        tk.Radiobutton(frame_envio, text="Disparo Oficial", variable=self.envio_var, value="oficial").pack(side=tk.LEFT)
        
        # Exceções de Envio
        self.excecoes_envio = []
        tk.Button(frame_etapa2, text="Gerenciar Exceções (Ignorar Profissionais)", command=self.abrir_janela_excecoes).grid(row=5, column=0, columnspan=4, pady=5)
        
        self.btn_etapa2 = tk.Button(frame_etapa2, text="Executar Etapa 2 (Gerar HTML)", command=self.start_etapa2, bg="green", fg="white", font=("Arial", 10, "bold"))
        self.btn_etapa2.grid(row=6, column=0, columnspan=4, pady=10)
        
        # Logs
        frame_logs = tk.Frame(self.tab_gerador)
        frame_logs.pack(fill="x", padx=10)
        
        tk.Label(frame_logs, text="Logs do Sistema:").pack(side=tk.LEFT)
        tk.Button(frame_logs, text="Salvar Log", command=self.save_log).pack(side=tk.RIGHT)
        
        self.txt_log = tk.Text(self.tab_gerador, height=10, width=95)
        self.txt_log.pack(pady=5, padx=10)
        
        self.setup_tab_contatos()
        
    def setup_tab_contatos(self):
        # Frame for Treeview
        frame_tree = tk.Frame(self.tab_contatos)
        frame_tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        columns = ("Grupo", "Nome", "Email")
        self.tree_contatos = ttk.Treeview(frame_tree, columns=columns, show="headings")
        self.tree_contatos.heading("Grupo", text="Grupo")
        self.tree_contatos.heading("Nome", text="Nome")
        self.tree_contatos.heading("Email", text="Email")
        
        self.tree_contatos.column("Grupo", width=150)
        self.tree_contatos.column("Nome", width=250)
        self.tree_contatos.column("Email", width=350)
        
        scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=self.tree_contatos.yview)
        self.tree_contatos.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_contatos.pack(side=tk.LEFT, fill="both", expand=True)
        
        # Bind select
        self.tree_contatos.bind("<<TreeviewSelect>>", self.on_contato_select)
        
        # Frame for controls
        frame_controls = tk.LabelFrame(self.tab_contatos, text="Adicionar / Editar Contato", padx=10, pady=10)
        frame_controls.pack(fill="x", padx=10, pady=5)
        
        tk.Label(frame_controls, text="Nome:").grid(row=0, column=0, sticky="e")
        self.entry_nome_contato = tk.Entry(frame_controls, width=30)
        self.entry_nome_contato.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(frame_controls, text="Email:").grid(row=0, column=2, sticky="e")
        self.entry_email_contato = tk.Entry(frame_controls, width=30)
        self.entry_email_contato.grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(frame_controls, text="Grupo:").grid(row=1, column=0, sticky="e")
        self.combo_grupo = ttk.Combobox(frame_controls, values=[
            "Narrador", "Comentarista Futebol", "Comentaristas (outros)", 
            "Comentaristas Arbitragem", "Colaboradores", "Outros / Desconhecidos"
        ], state="readonly", width=27)
        self.combo_grupo.grid(row=1, column=1, padx=5, pady=5)
        
        # Buttons
        frame_btns = tk.Frame(frame_controls)
        frame_btns.grid(row=1, column=2, columnspan=2, pady=5)
        
        tk.Button(frame_btns, text="Adicionar", command=self.add_contato).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btns, text="Atualizar Selecionado", command=self.update_contato).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btns, text="Excluir", command=self.delete_contato).pack(side=tk.LEFT, padx=5)
        
        # Save Button
        tk.Button(self.tab_contatos, text="Salvar Planilha de Contatos", command=self.save_contatos, bg="blue", fg="white", font=("Arial", 10, "bold")).pack(pady=10)
        
        # Load Contacts
        self.load_contatos()
        
    def load_contatos(self):
        for item in self.tree_contatos.get_children():
            self.tree_contatos.delete(item)
            
        contatos_file = 'contatos_nova_versao.xlsx'
        if os.path.exists(contatos_file):
            try:
                df = pd.read_excel(contatos_file, sheet_name='Lista e-mails')
                if 'Grupo' in df.columns and 'Nome' in df.columns and 'Email' in df.columns:
                    for _, r in df.iterrows():
                        self.tree_contatos.insert("", tk.END, values=(str(r['Grupo']), str(r['Nome']), str(r.get('Email', ''))))
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar contatos: {e}")
                
    def on_contato_select(self, event):
        selected = self.tree_contatos.selection()
        if selected:
            item = self.tree_contatos.item(selected[0])
            values = item['values']
            if values:
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
        contatos_file = 'contatos_nova_versao.xlsx'
        data = []
        for item in self.tree_contatos.get_children():
            v = self.tree_contatos.item(item)['values']
            data.append({"Grupo": v[0], "Nome": v[1], "Email": v[2]})
            
        df = pd.DataFrame(data)
        try:
            df.to_excel(contatos_file, index=False, sheet_name='Lista e-mails')
            messagebox.showinfo("Sucesso", "Planilha de contatos atualizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar planilha: {e}")

    def log(self, message):
        self.txt_log.insert(tk.END, message + "\\n")
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
        
    def save_log(self):
        log_content = self.txt_log.get("1.0", tk.END)
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(log_content)
            messagebox.showinfo("Sucesso", f"Log salvo em: {file_path}")
        
    def abrir_janela_excecoes(self):
        check_path = self.entry_check.get()
        if not check_path or not os.path.exists(check_path):
            messagebox.showerror("Erro", "Selecione o arquivo Check Pre Envio primeiro para carregar os nomes.")
            return
            
        try:
            df = pd.read_excel(check_path)
            nome_col = 'Nome ' if 'Nome ' in df.columns else 'Nome'
            if nome_col not in df.columns:
                messagebox.showerror("Erro", f"Coluna de Nome ({nome_col}) não encontrada na planilha.")
                return
            profissionais = sorted([str(p).strip() for p in df[nome_col].dropna().unique() if str(p).strip() not in ['nan', '-']])
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
            selecionados = [listbox.get(i) for i in listbox.curselection()]
            self.excecoes_envio = selecionados
            self.log(f"{len(selecionados)} exceções salvas. Estes profissionais serão ignorados na geração.")
            top.destroy()
            
        tk.Button(top, text="Salvar Exceções", command=salvar_excecoes, bg="orange", font=("Arial", 10, "bold")).pack(pady=10)

    def browse_file(self, entry_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
            
    def start_etapa1(self):
        f_2468 = self.entry_2468.get()
        if not f_2468 or not os.path.exists(f_2468):
            messagebox.showerror("Erro", "Selecione o Relatório 2468.")
            return
            
        self.btn_etapa1.config(state=tk.DISABLED)
        self.log("Iniciando Motor de Cruzamento (Etapa 1)...")
        threading.Thread(target=self.process_etapa1, args=(
            f_2468, 
            self.entry_sportv.get(), 
            self.entry_premiere.get(), 
            self.entry_combate.get()
        )).start()

    def process_etapa1(self, path_2468, path_sportv, path_premiere, path_combate):
        try:
            self.log("Lendo relatórios e grades... (Isso pode demorar alguns segundos)")
            from engine_cross import run_etapa1
            out_path = run_etapa1(path_2468, path_sportv, path_premiere, path_combate)
            
            self.log(f"Etapa 1 concluída! Arquivo gerado: {out_path}")
            
            # Auto-preencher Etapa 2
            self.entry_check.delete(0, tk.END)
            self.entry_check.insert(0, out_path)
            
        except Exception as e:
            self.log(f"Erro na Etapa 1: {e}")
        finally:
            self.btn_etapa1.config(state=tk.NORMAL)
            
    def start_etapa2(self):
        check_path = self.entry_check.get()
        if not check_path or not os.path.exists(check_path):
            messagebox.showerror("Erro", "Selecione o arquivo Check Pre Envio.")
            return
            
        selected_groups = []
        if self.var_narradores.get(): selected_groups.append("Narrador")
        if self.var_coment_futebol.get(): selected_groups.append("Comentarista Futebol")
        if self.var_coment_outros.get(): selected_groups.append("Comentaristas (outros)")
        if self.var_coment_arbitragem.get(): selected_groups.append("Comentaristas Arbitragem")
        if self.var_colaboradores.get(): selected_groups.append("Colaboradores")
        if self.var_outros.get(): selected_groups.append("Outros / Desconhecidos")
            
        self.btn_etapa2.config(state=tk.DISABLED)
        self.log("Iniciando Geração de HTMLs (Etapa 2)...")
        threading.Thread(target=self.process_etapa2, args=(check_path, self.entry_contacts.get(), selected_groups)).start()

    def process_etapa2(self, check_path, contacts_path, selected_groups):
        try:
            self.log("Lendo arquivo consolidado...")
            df = pd.read_excel(check_path)
            
            # Filter empty names
            nome_col = 'Nome ' if 'Nome ' in df.columns else 'Nome'
            if nome_col not in df.columns:
                self.log(f"ERRO: Coluna '{nome_col}' não encontrada na planilha.")
                return
            df = df.dropna(subset=[nome_col])
            df = df[df[nome_col] != '']
            
            # Date format and filtering
            if 'Data' in df.columns:
                df['Data_obj'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
                
                # Check Period filter
                data_inicio = self.entry_data_inicio.get().strip()
                data_fim = self.entry_data_fim.get().strip()
                
                if data_inicio and data_fim:
                    try:
                        dt_inicio = pd.to_datetime(data_inicio, format="%d/%m/%Y")
                        dt_fim = pd.to_datetime(data_fim, format="%d/%m/%Y")
                        mask = (df['Data_obj'] >= dt_inicio) & (df['Data_obj'] <= dt_fim)
                        df = df[mask]
                        self.log(f"Filtro aplicado: mostrando eventos de {data_inicio} até {data_fim}.")
                    except Exception as e:
                        self.log(f"Erro ao processar datas de filtro. Verifique o formato DD/MM/YYYY. Detalhe: {e}")
                
                dias_pt = {
                    'Monday': 'Segunda-feira', 'Tuesday': 'Terça-feira', 'Wednesday': 'Quarta-feira',
                    'Thursday': 'Quinta-feira', 'Friday': 'Sexta-feira', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
                }
                df['Dia'] = df['Data_obj'].dt.strftime('%A').map(dias_pt)
                df['Data'] = df['Data_obj'].dt.strftime('%d/%m/%Y')
                
            if df.empty:
                self.log("Nenhum evento encontrado no período especificado.")
                return
            df = df.fillna("-")
            
            contacts_path = contacts_path.strip() if contacts_path else ""
            if not contacts_path or not os.path.exists(contacts_path):
                self.log("ERRO: Planilha de contatos é obrigatória para filtrar por grupos.")
                return
                
            self.log("Lendo grupos da planilha de contatos...")
            contacts_dict = {}
            name_to_group = {}
            try:
                df_contacts = pd.read_excel(contacts_path, sheet_name='Lista e-mails')
                if 'Grupo' in df_contacts.columns and 'Nome' in df_contacts.columns:
                    for _, r in df_contacts.iterrows():
                        n = str(r['Nome']).strip().lower()
                        g = str(r['Grupo']).strip()
                        e = str(r.get('Email', '')).strip()
                        if n and n != 'nan':
                            name_to_group[n] = g
                            if e and e != 'nan':
                                contacts_dict[n] = e
                else:
                    self.log("ERRO: Formato da aba 'Lista e-mails' inválido. Selecione a nova versão de contatos gerada (contatos_nova_versao.xlsx).")
                    return
            except Exception as e:
                self.log(f"Erro ao ler contatos: {e}")
                return
                
            import difflib
            # Aplicar filtro de grupos
            profissionais_a_gerar = []
            for prof in df[nome_col].unique():
                prof_lower = prof.lower()
                grupo_do_prof = name_to_group.get(prof_lower)
                
                # Se não encontrar direto, tenta aproximação leve
                if not grupo_do_prof:
                    matches = difflib.get_close_matches(prof_lower, name_to_group.keys(), n=1, cutoff=0.8)
                    if matches:
                        grupo_do_prof = name_to_group[matches[0]]
                        
                if not grupo_do_prof:
                    grupo_do_prof = "Outros / Desconhecidos"
                    
                if grupo_do_prof in selected_groups and prof not in self.excecoes_envio:
                    profissionais_a_gerar.append(prof)
                    
            if not profissionais_a_gerar:
                self.log("Nenhum profissional restou após os filtros (verifique grupos e exceções).")
                return
                
            df = df[df[nome_col].isin(profissionais_a_gerar)]
                
            profissionais = df[nome_col].unique()
            self.log(f"Gerando HTML para {len(profissionais)} profissionais...")
            
            output_dir = os.path.join(os.path.dirname(check_path), "escalas_geradas_html")
            os.makedirs(output_dir, exist_ok=True)
            
            for prof in profissionais:
                df_prof = df[df[nome_col] == prof]
                if not df_prof.empty:
                    self.gerar_html(prof, df_prof, output_dir)
                    
            self.log(f"Geração concluída! HTMLs salvos em: {output_dir}")
            
            # Sending mode logic
            modo_envio = self.envio_var.get()
            if modo_envio == "gerar":
                self.log("Modo: 'Somente Gerar HTML'. Nenhum email foi enviado.")
            elif modo_envio == "teste":
                self.log("Modo: 'Enviar Teste (Para mim)'. Preparando envio...")
                self.enviar_emails(output_dir, contacts_dict, teste=True)
            elif modo_envio == "oficial":
                self.log("Modo: 'Disparo Oficial'. Preparando disparo para base...")
                self.enviar_emails(output_dir, contacts_dict, teste=False)
                
        except Exception as e:
            self.log(f"Erro na Etapa 2: {e}")
        finally:
            self.btn_etapa2.config(state=tk.NORMAL)
            
    def gerar_html(self, nome, df, output_dir):
        html_template = """
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; }}
                h2 {{ color: #333; }}
                table {{ border-collapse: collapse; width: 100%; font-size: 12px; }}
                th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
            </style>
        </head>
        <body>
            <h2>Escala - {nome}</h2>
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
        </body>
        </html>
        """
        
        rows_html = ""
        for _, row in df.iterrows():
            evento = str(row.get('Atividade/Descrição', row.get('Evento/Descrição', '-'))).strip()
            
            def format_time(time_str):
                if not time_str or str(time_str).strip() in ['-', 'nan', 'NaT', '']: return '-'
                t = str(time_str).strip()
                if ' ' in t: t = t.split(' ')[-1]
                parts = t.split(':')
                return f"{parts[0]}:{parts[1]}" if len(parts) >= 2 else t
                
            inicio = format_time(row.get('Início', '-'))
            fim = format_time(row.get('Fim', '-'))
            pre = format_time(row.get('Pré', '-'))
            
            rows_html += "<tr>"
            rows_html += f"<td>{row.get('Nome ', row.get('Nome', '-'))}</td>"
            rows_html += f"<td>{row.get('Plataforma', '-')}</td>"
            rows_html += f"<td>{row.get('Data', '-')}</td>"
            rows_html += f"<td>{row.get('Dia', '-')}</td>"
            rows_html += f"<td>{pre}</td>"
            rows_html += f"<td>{inicio}</td>"
            rows_html += f"<td>{fim}</td>"
            rows_html += f"<td>{evento}</td>"
            rows_html += f"<td>{row.get('Produto (WO/Shift)', row.get('Produto', '-'))}</td>"
            rows_html += f"<td>{row.get('Local de Gravação', row.get('Local Narração', row.get('Local', '-')))}</td>"
            rows_html += f"<td>{row.get('Elenco', '-')}</td>"
            rows_html += f"<td>{row.get('Coordenador', '-')}</td>"
            rows_html += f"<td>{row.get('Produtor', '-')}</td>"
            rows_html += "</tr>"
            
        final_html = html_template.format(nome=nome, rows=rows_html)
        file_name = f"escala_{nome.replace(' ', '_')}.html"
        file_path = os.path.join(output_dir, file_name)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(final_html)
            
    def enviar_emails(self, html_dir, contacts, teste=False):
        self.log("Conectando ao Outlook e preparando rascunhos...")
        try:
            outlook = win32com.client.Dispatch('outlook.application')
        except Exception as e:
            self.log(f"Erro ao abrir o Outlook: {e}. Verifique se ele está instalado.")
            return

        import difflib
        arquivos = [f for f in os.listdir(html_dir) if f.endswith('.html')]
        self.log(f"Foram encontrados {len(arquivos)} arquivos HTML.")
        
        for arquivo in arquivos:
            nome = arquivo.replace('escala_', '').replace('.html', '').replace('_', ' ')
            
            # Buscar e-mail
            email_dest = contacts.get(nome.lower(), "")
            if not email_dest:
                matches = difflib.get_close_matches(nome.lower(), contacts.keys(), n=1, cutoff=0.7)
                if matches:
                    email_dest = contacts[matches[0]]
            
            caminho_completo = os.path.join(html_dir, arquivo)
            with open(caminho_completo, 'r', encoding='utf-8') as f:
                html_body = f.read()
                
            try:
                mail = outlook.CreateItem(0)
                if teste:
                    mail.Subject = f"[TESTE] Escala - {nome}"
                    mail.To = "" # Em teste, deixar em branco para o próprio usuário preencher se quiser
                else:
                    mail.Subject = f"Escala - {nome}"
                    mail.To = email_dest
                
                mail.HTMLBody = html_body
                mail.Display() # Abre no Outlook do usuário
            except Exception as e:
                self.log(f"Erro ao gerar draft para {nome}: {e}")
                
        self.log("Pronto! As janelas do Outlook foram abertas. Faça a conferência e envie manualmente.")

if __name__ == "__main__":
    root = tk.Tk()
    app = GeradorEscalasApp(root)
    root.mainloop()
