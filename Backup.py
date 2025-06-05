import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import pyautogui
import time
import threading
import re
from urllib.parse import quote
import os
import winsound

#Dicionário de templates de mensagens (permanece igual ao seu código atual)
mensagens = {
    "CLIENTE EM TREINAMENTO - INICIANDO O USO": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? Tudo certo? Algo a relatar? "
        "Prestamos suporte técnico rápido e ativo. Qualquer coisa, pode entrar em contato comigo, o Glaucio, "
        "ou com nosso Suporte Técnico no *(55) 9119 4370* com a (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "PROSPECTS QUE ESTAMOS BUSCANDO UM CONTATO": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* trabalhamos com um *Sistema de Gestão: Simples e Prático*. "
        "Montamos a mensalidade baseada nas ferramentas que realmente for utilizar. "
        "Valores a partir de R$ 69,90. Peça mais informações, comigo , o Glaucio. "
        "Se não quiser mais receber informações sobre nossos serviços, envie *SAIR*."
    ),
    "CLIENTE QUE UTILIZA SO MAQUININHA DE CARTÃO": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do aplicativo na maquininha de cartões? Tudo certo? Algo a relatar? "
        "Entre em contato para mais informações no (54) 9 9104 1029. "
        "Prestamos suporte técnico rápido e ativo. "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "PROSPECT QUE FOI ATÉ ORÇADO, EM NEGOCIAÇÃO": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* trabalhamos com um *Sistema de Gestão: Simples e Prático*. "
        "Já tinhamos conversado anteriormente, chegamos a falar um pouco sobre o sistema, e até foi sugerido um valor de mensalidade para usar o sistema, " 
        "e agora temos *Novidades* e *Promoções* exclusivas para você. ""Entre em contato comigo, o Glaucio! "
        "Mas Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "EMPRESAS PARCEIRAS QUE NOS INDICAM": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* reafirmamos nossa parceria e temos *Novidades* e *Promoções* "
        "exclusivas para seus Clientes nesse novo ciclo. Ferramentas como a Emissão da NF-e, e o Cupom Eletrônico interligado com as Máquinas de Cartões."
        "Entre em contato comigo, o Glaucio!, para mais informações."
        "Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "ESCRITÓRIOS CONTÁBEIS - CONTADORES": (
        ", {nome}, da empresa {empresa}, "
        "Visitei seu escritório pessoalmente, e agora estamos entrando em contato novamente para apresentar uma "
        "*Solução Prática e Simples* para seus Clientes com relação ao *Sistema de Gestão*; com um Controle de Estoque e Caixa, Além da parte Fiscal, com a Emissão da Nota Eletrônica e o Cupom Eletrônico. "
        "Entre em contato comigo, o Glaucio! "
        "Caso não queira mais receber esse tipo de informação, envie *SAIR*."
    ),
    "CLIENTES QUE UTILIZAM LIGACAO PC X SMART": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? Sabemos que utiliza a maquininha Smart, com ligação ao sistema de Emissão do Cupom. "
        "E as vezes acontecem alguns problemas, geralmente devido a oxilação da Internet, mas agora temos a opção de ligação direta nas máquinas do Banrisul, Sicredi e Stone. "
        "Se quiser trocar para essa nova forma, ela é bem mais em conta referente a valores mensais, peça mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna).  "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "APLICATIVO PROPRIO - EVOLUTI PLAY - BANRISUL": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Temos a solução para ligação direta nas máquinas de cartão SMARTs, utilizando Banrisul, Sicredi e Stone, "
        "que atendendem às exigências fiscais atuais. Interligando a emissão do Cupom Eletrônico com as maquininhas. "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
     "APLICATIVO PROPRIO - EVOLUTI PLAY - SICREDI": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Temos a solução para ligação direta nas máquinas de cartão SMARTs, utilizando Banrisul, Sicredi e Stone, "
        "que atendendem às exigências fiscais atuais. Interligando a emissão do Cupom Eletrônico com as maquininhas. "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
     "APLICATIVO PROPRIO - EVOLUTI PLAY - STONE": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Temos a solução para ligação direta nas máquinas de cartão SMARTs, utilizando Banrisul, Sicredi e Stone, "
        "que atendendem às exigências fiscais atuais. Interligando a emissão do Cupom Eletrônico com as maquininhas. "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "CLIENTES QUE USAM  O TEF CONVENCIONAL": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? Sabemos que utiliza o TEF na emissão de cupom eletrônico. "
        "Às vezes surgem problemas, principalmente a questão do recebimento dos boletos,  mas agora temos a opção de ligação direta nas máquinas de cartão as SMARTs, temos com o Banrisul, Sicredi e Stone. "
        "Se quiser trocar para essa nova forma, fale comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "CLIENTES QUE USAM BASICAMENTE A NFE": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Sabemos que utiliza principalmente a Emissão de Nota Eletrônica. "
        "Estamos sempre à disposição para ajudar. "
        "pode falar comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES QUE RARAMENTE PEDEM SUPORTE": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? "
        "Sabemos que dificilmente precisa de ajuda. "
        "Mas estamos prontos para ajudar, caso precisar de algo. Basta falar com o nosso suporte no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES QUE USAM MUITO POUCO": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? "
        "Sabemos que utiliza poucas opções e ferramentas. "
        "Mas estamos prontos para ajudar, caso precisar de algo. Basta falar com o nosso suporte no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES QUE USAM TODAS AS FERRAMENTAS": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Além de todas as opções que já utiliza, estamos sempre à disposição para ajudar. Caso precise de suporte basta falar com o nosso Whats no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES DO MEI - FERRAMENTAS SEM DOCS FISCAIS": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Além das opções que já utiliza, temos ferramentas para a parte fiscal, incluindo emissão de NF-e e o Cupom Eletrônico. e Agora com essa questão da Ligação com as maquininhas de Cartões.  "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
}

class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root
        self.drive = None
        self.root.title("Envio Automático WhatsApp v5")
        self.root.geometry("820x680+100+5")
        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica", 8, "bold"))  # <-- Aumenta fonte e coloca negrito
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))  # Cabeçalhos também

    

        # Evento para iniciar/parar o envio
        
        self.enviando = False
        self.envio_ativo = threading.Event()

        self.df = None
        self.filtered_df = None

        # Topo - Seleção do arquivo Excel e planilha
        frame_top = ttk.Frame(root, padding=10)
        frame_top.pack(fill='x')

        ttk.Label(frame_top, text="Arquivo Excel:").pack(side='left')
        self.entry_file = ttk.Entry(frame_top, width=40)
        self.entry_file.pack(side='left', padx=5)
        ttk.Button(frame_top, text="Selecionar", command=self.select_file).pack(side='left')

        ttk.Label(frame_top, text="Planilha:").pack(side='left', padx=(20, 5))
        self.combo_sheet = ttk.Combobox(frame_top, state='readonly', width=20)
        self.combo_sheet.pack(side='left')
        self.combo_sheet.bind("<<ComboboxSelected>>", self.sheet_selected)

        # Meio - Tipo de mensagem e Grupo/Sublista
        frame_middle = ttk.Frame(root, padding=10)
        frame_middle.pack(fill='x')

        tk.Label(frame_middle, text="Selecionar Tipo Escrita da Mensagem:", font=("TkDefaultFont", 10, "bold")).pack(side='left')
        self.combo_msg_type = ttk.Combobox(frame_middle, state='readonly', width=50)
        self.combo_msg_type.pack(side='left', padx=5)
        self.combo_msg_type['values'] = sorted(mensagens.keys())
        self.combo_msg_type.bind("<<ComboboxSelected>>", self.edit_template)

        ttk.Label(frame_middle, text="Grupo/Categoria:").pack(side='left', padx=(20, 5))
        self.combo_group = ttk.Combobox(frame_middle, state='readonly', width=20)
        self.combo_group.pack(side='left', padx=5)

       
        # Edição de template
        frame_edit = ttk.Frame(root, padding=50)
        frame_edit.pack(fill='x')

# Título acima do campo de texto
        tk.Label(
            frame_edit,
            text="Modelo da Mensagem a ser Enviada: PODE EDITAR ANTES DE CARREGAR NUMEROS",
            font=("Arial", 12, "bold"),
            anchor='w'
        ).pack(fill='x', padx=5, pady=(0, 5))

# Campo de texto para o template
        self.txt_template = tk.Text(frame_edit, height=5, wrap='word')
        self.txt_template.pack(fill='x', expand=True)


        # Treeview com as mensagens a serem disparadas
        frame_bottom = ttk.Frame(root, padding=10)
        frame_bottom.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(
            frame_bottom,
            columns=('Telefone', 'Mensagem'),
            show='headings'
        )
        self.tree.heading('Telefone', text='Telefone')
        self.tree.heading('Mensagem', text='Mensagens a serem Disparadas')
        self.tree.column('Telefone', width=80)
        self.tree.column('Mensagem', width=1200)
        self.tree.pack(fill='both', expand=True)

        # Botões de Ação
        frame_actions = ttk.Frame(root, padding=10)
        frame_actions.pack(fill='x')

        tk.Button(
            frame_actions,
            text="EDITAR",
            command=self.edit_message,
            bg="#b34b00",  # azul forte
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="EXCLUIR",
            command=self.delete_message,
            bg="#dc3545",  # vermelho
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="Limpar Mensagem",
            command=self.cancel_operation,
            bg="#ffc107",  # amarelo
            fg="black"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="Carregar TELEFONES",
            command=self.load_messages,
            bg="#0056b3",
            fg="white"
        ).pack(side='left', padx=10)

        # Botão ENVIAR agora chama iniciar_envio_thread
        tk.Button(
            frame_actions,
            text="Disparar Mensagens",
            command=self.iniciar_envio_thread,
            bg="#28a745",  # verde
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="Pausar Envio",
            command=self.pausar_envio,
            bg="#ff8000",  # Laranja
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="FECHAR",
            command=root.destroy,  # ou self.root.destroy se root estiver como self.root
            bg="#6c757d",  # cinza
            fg="white"
        ).pack(side='left', padx=5)

        # Contador de mensagens carregadas
        self.label_info = ttk.Label(root, text="Nenhuma mensagem carregada.")
        self.label_info.pack(anchor='w', padx=10)

    def pausar_envio(self):
        self.envio_ativo.clear()  # Desativa o sinal para pausar o envio
        self.enviando = False
        print("Envio pausado manualmente.")
        messagebox.showerror("Disparo PAUSADO"," Clicle Novamente no botão 'DISPARAR MENSAGENS' para Retomar o Envio.")

    def select_file(self):
    # Verifica se já há um caminho no campo de entrada
        current_path = self.entry_file.get()
    
    # Se o caminho já existe e o arquivo está presente, carrega diretamente
        if current_path and os.path.isfile(current_path):
            self.load_sheets(current_path)
            return

    # Caso contrário, abre o diálogo de seleção de arquivo
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)
            self.load_sheets(file_path)

    def load_sheets(self, filepath):
        try:
            xls = pd.ExcelFile(filepath)
            self.combo_sheet['values'] = xls.sheet_names
            if xls.sheet_names:
                self.combo_sheet.current(0)
                self.sheet_selected()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir arquivo Excel:\n{e}")

    def sheet_selected(self, event=None):
        filepath = self.entry_file.get()
        sheet = self.combo_sheet.get()
        try:
            # Carrega o DataFrame e padroniza colunas em minúsculo
            self.df = pd.read_excel(filepath, sheet_name=sheet)
            self.df.columns = [str(col).strip().lower() for col in self.df.columns]
            self.update_group_list()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha:\n{e}")

    def update_group_list(self):
        # Preenche Combobox de Grupo/Sublista a partir da coluna 'grupo'
        if self.df is not None and 'grupo' in self.df.columns:
            grupos = sorted(self.df['grupo'].dropna().unique())
            self.combo_group['values'] = grupos
            if grupos:
                self.combo_group.current(0)

    def edit_template(self, event=None):
        tipo = self.combo_msg_type.get()
        template = mensagens.get(tipo, "")
        self.txt_template.delete('1.0', tk.END)
        self.txt_template.insert(tk.END, template)

    def load_messages(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue um arquivo primeiro.")
            return

        tipo = self.combo_msg_type.get()
        grupo = self.combo_group.get()
        if not tipo or not grupo:
            messagebox.showwarning("Aviso", "Selecione o Tipo de Mensagem a ser Enviada.")
            return

        if 'grupo' not in self.df.columns:
            messagebox.showerror("Erro", "A planilha não possui a coluna 'grupo'.")
            return

        # Filtra clientes do grupo selecionado
        self.filtered_df = self.df[self.df['grupo'] == grupo].copy()

        # Obtém template editado
        template = self.txt_template.get("1.0", tk.END).strip()

        # Limpa treeview
        self.tree.delete(*self.tree.get_children())
        count = 0

        for _, row in self.filtered_df.iterrows():
            nome = str(row.get('nome', 'Cliente')).strip()
            empresa = str(row.get('empresa', 'Empresa')).strip()
            telefone = str(row.get('telefone', '')).strip()
            inicio = str(row.get('inicio', '')).strip()

            # Monta mensagem: <início> + <template>
            msg = f"{inicio}{template.format(nome=nome, empresa=empresa)}".strip()
            self.tree.insert('', tk.END, values=(telefone, msg))
            count += 1

        self.label_info.config(text=f"Total de mensagens carregadas a DISPARAR: {count}")

    def edit_message(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Editar", "Selecione uma mensagem para editar.")
            return

        item = selected[0]
        contato, mensagem = self.tree.item(item, 'values')

        def salvar():
            nova = txt.get("1.0", tk.END).strip()
            self.tree.item(item, values=(contato, nova))
            edit_window.destroy()

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Mensagem")
        txt = tk.Text(edit_window, height=10, width=60)
        txt.pack(padx=10, pady=10)
        txt.insert('1.0', mensagem)
        ttk.Button(edit_window, text="Salvar", command=salvar).pack(pady=5)

    def delete_message(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Excluir", "Selecione uma mensagem para excluir.")
            return

        for item in selected:
            self.tree.delete(item)

        total_restante = len(self.tree.get_children())
        self.label_info.config(text=f"Total de mensagens: {total_restante}")

    def cancel_operation(self):
        self.tree.delete(*self.tree.get_children())
        self.txt_template.delete('1.0', tk.END)
        self.label_info.config(text="Operação cancelada. Nenhuma mensagem carregada.")

    def enviar_mensagem_com_enter(self, cliente):
        try:
            telefone_formatado = cliente["telefone"]
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone_formatado}&text={quote(cliente["mensagem"])}'
            webbrowser.open(link_mensagem_whatsapp)
            
            time.sleep(10)
            pyautogui.press('enter')
            time.sleep(5)
        except Exception as e:
            print(f"Erro ao enviar para {cliente['telefone']}: {e}")
            

    def iniciar_envio_thread(self):
        if self.enviando:
            self.label_info.config(text="O envio já está em andamento.")
            return

        self.enviando = True
        self.envio_ativo.set()

        def enviar_mensagens():
            for item in self.tree.get_children():
                if not self.envio_ativo.is_set():
                    break

                contato, mensagem = self.tree.item(item, "values")
                cliente = {"telefone": contato, "mensagem": mensagem}

                self.enviar_mensagem_com_enter(cliente)

                self.tree.delete(item)  # Remove a linha da Treeview após o envio

                total_restante = len(self.tree.get_children())
                self.label_info.config(text=f"Mensagens restantes a DISPARAR: {total_restante}")

                time.sleep(110)  # Espera 2 minutos entre mensagens

            # Código final
        self.enviando = False
        self.label_info.config(text="Envio concluído.")

# Toca um som padrão do Windows (como beep)
        
        winsound.Beep(frequency=1000, duration=1000)  # frequência em Hz, duração em ms

        thread = threading.Thread(target=enviar_mensagens, daemon=True)
        thread.start()


    def parar_envio(self):
        self.envio_ativo.clear()
        print("Envio interrompido.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()
