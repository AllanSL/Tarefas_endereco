import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from processamento import processar_arquivo

def escolher_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo 'Relatório Gestão de Tarefas'",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho:
        iniciar_processamento(caminho)

def iniciar_processamento(caminho):
    # Desativa botão
    botao.config(state="disabled")
    
    # Cria barra de progresso determinística
    barra_progresso = ttk.Progressbar(root, orient="horizontal", mode="determinate", maximum=100, length=300)
    barra_progresso.pack(pady=10)
    barra_progresso["value"] = 0

    # Função para atualizar progresso
    def atualizar_progresso(valor):
        barra_progresso["value"] = valor
        root.update_idletasks()

    try:
        processar_arquivo(caminho, progresso_callback=atualizar_progresso)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    finally:
        barra_progresso.destroy()
        botao.config(state="normal")

# Interface gráfica
root = tk.Tk()
root.title("Locais - Tarefas")
root.resizable(False, False)

# Define tamanho da janela
largura = 400
altura = 200

# Pega a largura e altura da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Calcula a posição x e y para centralizar
pos_x = (largura_tela // 2) - (largura // 2)
pos_y = (altura_tela // 2) - (altura // 2)

# Define geometria centralizada
root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

label = tk.Label(root, text="Clique abaixo para escolher o relatório:", font=("Arial", 12))
label.pack(pady=20)

botao = tk.Button(root, text="Selecionar arquivo", font=("Arial", 12), command=escolher_arquivo)
botao.pack(pady=10)

root.mainloop()
