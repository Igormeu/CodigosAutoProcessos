import os

caminho_pasta = "\\\\apolo\\Governanca\\PROCESSOS\\MAPEAMENTO DE PROCESSOS\\CODIGOS E AUTOMACOES\\Codigos.PY\\Codigos.PY\\Atualizar_BI"
arquivos = os.listdir(caminho_pasta)

# Filtrar apenas arquivos (ignorando pastas)
arquivos = [arquivo for arquivo in arquivos if os.path.isfile(os.path.join(caminho_pasta, arquivo))]

# Exibir os arquivos
print("Arquivos encontrados:")
for arquivo in arquivos:
    print(arquivo)