import win32com.client
import os
import signal
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import pythoncom

class OutlookAttachmentDownloader:
    def __init__(self):
        self.interromper = False
        self.setup_gui()
        
    def setup_gui(self):
        """Configura a interface gráfica"""
        self.root = tk.Tk()
        self.root.title("Importador de Anexos do Outlook")
        
        # Criar e configurar os campos de entrada
        self.criar_campos_entrada()
        
        # Criar frame de botões
        self.criar_botoes()
        
        # Configurar barra de progresso
        self.configurar_barra_progresso()
        
        # Configurar eventos
        self.configurar_eventos()
        
    def criar_campos_entrada(self):
        """Cria os campos de entrada da interface"""
        campos = [
            ("E-mail:", "email_entry"),
            ("Pasta:", "pasta_entry"),
            ("Caminho da pasta:", "diretorio_entry"),
            ("Tipo de arquivo/Extensão:", "extensao_entry")
        ]
        
        for i, (label_text, entry_name) in enumerate(campos):
            tk.Label(self.root, text=label_text).grid(row=i, column=0, padx=10, pady=5)
            entry = tk.Entry(self.root, width=50)
            entry.grid(row=i, column=1, padx=10, pady=5)
            setattr(self, entry_name, entry)
    
    def criar_botoes(self):
        """Cria os botões da interface"""
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        tk.Button(button_frame, text="Iniciar Importação", 
                 command=self.iniciar_importacao).pack(side=tk.LEFT, padx=10)
        
        tk.Button(button_frame, text="Interromper Importação",
                 command=self.interromper_importacao).pack(side=tk.LEFT, padx=10)
    
    def configurar_barra_progresso(self):
        """Configura a barra de progresso"""
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
    
    def configurar_eventos(self):
        """Configura os eventos do teclado e sinais"""
        self.root.bind("<Return>", self.on_key_press)
        self.root.bind("<Escape>", self.on_key_press)
        signal.signal(signal.SIGINT, self.interromper_importacao)
    
    def validar_campos(self):
        """Valida os campos obrigatórios"""
        campos = {
            'email': self.email_entry.get(),
            'pasta': self.pasta_entry.get(),
            'diretorio': self.diretorio_entry.get()
        }
        
        return all(campos.values())
    
    def salvar_anexos(self, email, pasta, diretorio, extensao):
        """Função principal para salvar os anexos"""
        try:
            pythoncom.CoInitialize()
            outlook = self.conectar_outlook(email, pasta)
            folder = outlook['folder']
            
            total_items = len(folder.Items)
            self.processar_emails(folder, diretorio, extensao, total_items)
            
            self.mostrar_mensagem_final()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        finally:
            pythoncom.CoUninitialize()
    
    def conectar_outlook(self, email, pasta):
        """Estabelece conexão com o Outlook"""
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        if email not in [folder.Name for folder in namespace.Folders]:
            raise Exception(f"O perfil de e-mail '{email}' não foi encontrado.")
        
        try:
            folder = namespace.Folders[email].Folders[pasta]
            return {'outlook': outlook, 'namespace': namespace, 'folder': folder}
        except Exception:
            raise Exception(f"A pasta '{pasta}' não foi encontrada no e-mail '{email}'.")
    
    def processar_emails(self, folder, diretorio, extensao, total_items):
        """Processa os emails e salva os anexos"""
        for i, item in enumerate(folder.Items, 1):
            if self.interromper:
                break
                
            self.atualizar_progresso(i, total_items)
            
            if item.Class == 43:  # MailItem
                self.processar_anexos(item, diretorio, extensao)
    
    def processar_anexos(self, email, diretorio, extensao):
        """Processa os anexos de um email"""
        for attachment in email.Attachments:
            if self.interromper:
                break
                
            if not extensao or attachment.FileName.lower().endswith(extensao.lower()):
                caminho_completo = os.path.join(diretorio, attachment.FileName)
                attachment.SaveAsFile(caminho_completo)
    
    def mostrar_mensagem_final(self):
        """Mostra mensagem de conclusão apropriada"""
        if self.interromper:
            messagebox.showinfo("Informação", "Importação interrompida pelo usuário.")
        else:
            messagebox.showinfo("Informação", "Exportação finalizada :)")
    
    def atualizar_progresso(self, count, total):
        """Atualiza a barra de progresso"""
        self.progress_var.set(count / total * 100)
        self.root.update_idletasks()
    
    def iniciar_importacao(self):
        """Inicia o processo de importação"""
        self.interromper = False
        
        if not self.validar_campos():
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return
        
        # Coletar dados dos campos
        dados = {
            'email': self.email_entry.get(),
            'pasta': self.pasta_entry.get(),
            'diretorio': self.diretorio_entry.get(),
            'extensao': self.extensao_entry.get()
        }
        
        # Iniciar thread de processamento
        threading.Thread(target=self.salvar_anexos, kwargs=dados).start()
    
    def interromper_importacao(self, signal=None, frame=None):
        """Interrompe o processo de importação"""
        self.interromper = True
        print("Importação interrompida pelo usuário (via terminal ou botão).")
    
    def on_key_press(self, event):
        """Trata eventos de tecla"""
        if event.keysym in ["Return", "Escape"]:
            self.interromper_importacao()
    
    def executar(self):
        """Inicia a execução da aplicação"""
        self.root.mainloop()

if __name__ == "__main__":
    app = OutlookAttachmentDownloader()
    app.executar()