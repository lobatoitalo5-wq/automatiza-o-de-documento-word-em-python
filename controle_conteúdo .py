import tkinter as tk
from tkinter import filedialog
from docx import Document

# Cria a janela oculta do Tkinter
root = tk.Tk()
root.withdraw()

# Abre o explorador de arquivos para escolher o documento word
caminho_arquivo = filedialog.askopenfilename(
    title="Selecione um documento",
    filetypes=[("Todos os arquivos", "*.*"), ("Arquivos de Texto", "*.txt")]
)

# Mostra o caminho do arquivo selecionado no terminal
if caminho_arquivo:
    print(f"Arquivo escolhido: {caminho_arquivo}")
else:
    print("Nenhum arquivo foi selecionado.")

#Carregar o documento modelo
doc = Document(caminho_arquivo)


# 2. Definir os dados que vão substituir os marcadores
dados = {
    "{{nome}}": input("Qual seu nome?"),
    "{{cargo}}": input("Qual cargo? "),
    "{{data}}": input("informe a data: "),
}


# 3. Função para substituir o texto nos parágrafos
def substituir_texto(doc, dados):
  for paragrafo in doc.paragraphs:
    for chave, valor in dados.items():
      if chave in paragrafo.text:
        inline = paragrafo.runs
        for i in range(len(inline)):
          if chave in inline[i].text:
            inline[i].text = inline[i].text.replace(chave, valor)

# Executar a substituição
substituir_texto(doc, dados)

# 4. Salvar o novo documento preenchido
doc.save("documento_preenchido.docx")
print("Documento criado com sucesso!")

