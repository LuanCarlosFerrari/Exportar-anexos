import win32com.client
import os
import signal
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import pythoncom
import sys

class OutlookAttachmentDownloader:
    def __init__(self):
        self.interromper = False
        self.setup_gui()
        
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Importador de Anexos do Outlook")
        
        # Configurar codificação da interface
        if sys.platform == 'win32':
            self.root.option_add('*Font', 'System 10')
            
        self.criar_campos_entrada()
        self.criar_botoes()
        self.configurar_barra_progresso()
        self.configurar_eventos()
        
    def criar_campos_entrada(self):
        campos = [
            ("E-mail:", "email_entry"),
            ("Pasta:", "pasta_entry"),
            ("Caminho da pasta:", "diretorio_entry"),
            ("Tipo de arquivo/Extensão:", "extensao_entry")
        ]
        
        for i, (label_text, entry_name) in enumerate(campos):
            tk.Label(self.root, text=label_text).grid(row=i, column=0, padx=10, pady=5, sticky='w')
            entry = tk.Entry(self.root, width=50)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            setattr(self, entry_name, entry)
    
    def criar_botoes(self):
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        tk.Button(button_frame, text="Iniciar Importação", 
                 command=self.iniciar_importacao).pack(side=tk.LEFT, padx=10)
        
        tk.Button(button_frame, text="Interromper Importação",
                 command=self.interromper_importacao).pack(side=tk.LEFT, padx=10)
    
    def configurar_barra_progresso(self):
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
    
    def configurar_eventos(self):
        self.root.bind("<Return>", self.on_key_press)
        self.root.bind("<Escape>", self.on_key_press)
        signal.signal(signal.SIGINT, self.interromper_importacao)
    
    def validar_campos(self):
        campos = {
            'email': self.email_entry.get().strip(),
            'pasta': self.pasta_entry.get().strip(),
            'diretorio': self.diretorio_entry.get().strip()
        }
        
        if not all(campos.values()):
            return False
            
        # Validar se o diretório existe
        if not os.path.exists(campos['diretorio']):
            try:
                os.makedirs(campos['diretorio'])
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o diretório: {e}")
                return False
                
        return True
    
    def salvar_anexos(self, email, pasta, diretorio, extensao):
        try:
            pythoncom.CoInitialize()
            outlook = None
            try:
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            except Exception as e:
                raise Exception(f"Erro ao conectar com Outlook: {e}")

            # Tentar encontrar a conta de e-mail
            found_account = None
            for account in outlook.Folders:
                if email.lower() in str(account).lower():
                    found_account = account
                    break
                    
            if not found_account:
                raise Exception(f"Conta de e-mail '{email}' não encontrada")

            # Tentar encontrar a pasta
            try:
                folder = found_account.Folders[pasta]
            except Exception:
                raise Exception(f"Pasta '{pasta}' não encontrada na conta {email}")

            total_items = folder.Items.Count
            self.processar_emails(folder, diretorio, extensao, total_items)
            
            if not self.interromper:
                self.mostrar_mensagem_final()
            
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        finally:
            pythoncom.CoUninitialize()
    
    def processar_emails(self, folder, diretorio, extensao, total_items):
        try:
            items = folder.Items
            for i in range(total_items):
                if self.interromper:
                    break
                    
                try:
                    item = items[i+1]  # Outlook usa índice base 1
                    if item.Class == 43:  # MailItem
                        self.processar_anexos(item, diretorio, extensao)
                    self.atualizar_progresso(i+1, total_items)
                except Exception as e:
                    print(f"Erro ao processar email {i+1}: {e}")
                    continue
        except Exception as e:
            raise Exception(f"Erro ao processar emails: {e}")
    
    def processar_anexos(self, email, diretorio, extensao):
        try:
            for attachment in email.Attachments:
                if self.interromper:
                    break
                    
                nome_arquivo = str(attachment.FileName)
                if not extensao or nome_arquivo.lower().endswith(extensao.lower()):
                    caminho_completo = os.path.join(diretorio, nome_arquivo)
                    try:
                        attachment.SaveAsFile(caminho_completo)
                    except Exception as e:
                        print(f"Erro ao salvar anexo {nome_arquivo}: {e}")
        except Exception as e:
            print(f"Erro ao processar anexos: {e}")
    
    def mostrar_mensagem_final(self):
        if self.interromper:
            messagebox.showinfo("Informação", "Importação interrompida pelo usuário.")
        else:
            messagebox.showinfo("Informação", "Exportação finalizada com sucesso!")
    
    def atualizar_progresso(self, count, total):
        self.progress_var.set((count / total) * 100)
        self.root.update_idletasks()
    
    def iniciar_importacao(self):
        self.interromper = False
        
        if not self.validar_campos():
            messagebox.showerror("Erro", "Por favor, preencha todos os campos obrigatórios.")
            return
        
        dados = {
            'email': self.email_entry.get().strip(),
            'pasta': self.pasta_entry.get().strip(),
            'diretorio': self.diretorio_entry.get().strip(),
            'extensao': self.extensao_entry.get().strip()
        }
        
        threading.Thread(target=self.salvar_anexos, kwargs=dados, daemon=True).start()
    
    def interromper_importacao(self, signal=None, frame=None):
        self.interromper = True
        print("Importação interrompida pelo usuário.")
    
    def on_key_press(self, event):
        if event.keysym in ["Return", "Escape"]:
            self.interromper_importacao()
    
    def executar(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = OutlookAttachmentDownloader()
    app.executar()