import PyPDF2
import re
import pandas as pd

def extrair_texto_pdf(caminho_pdf):
    with open(caminho_pdf, 'rb') as arquivo_pdf:
        leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
        texto_completo = ""
        for pagina in range(len(leitor_pdf.pages)):
            pagina_atual = leitor_pdf.pages[pagina]
            texto_completo += pagina_atual.extract_text()
        return texto_completo

def extrair_informacoes(texto):
    numero_nota = re.findall(r"Número da Nota\s*(\d+)",texto)
    valor = re.findall(r"VALOR TOTAL DA NOTA\s*=\s*R\$\s*([\d\.]+,\d{2})", texto)  
    datas = re.findall(r"Data\s*e\s*Hora\s*de\s*Emissão\s*(\d{2}/\d{2}/\d{4})", texto)  
    mes_competencia = re.findall(r"Mês\s*de\s*Competência\s*da\s*Nota\s*Fiscal:\s*(\d{2}/\d{4})",texto)
    return numero_nota, valor, datas, mes_competencia


def coletar_informacoes (Lmes_competencia,Ldatas,Lvalor,Lnumero_nota,doc_ref,empresa,max):
    for i in range(max):
        if i+1 <= 9:
            nome_arquivo = f'{empresa} 0{i+1} - TAXA.pdf'
        else:
            nome_arquivo = f'{empresa} {i+1} - TAXA.pdf'
        
        caminho_pdf = f"Geral/{nome_arquivo}"
        
        try:
            texto = extrair_texto_pdf(caminho_pdf)
            numero_nota,valor,datas,mes_competencia = extrair_informacoes(texto)
            doc_ref.append(nome_arquivo)
            Lnumero_nota.append(numero_nota)
            Lvalor.append(valor)
            Ldatas.append(datas)
            Lmes_competencia.append(mes_competencia)
        except(FileNotFoundError):
            continue 
    
loja = {
    "1": ["EFA",70],
    "2": ["FROSTY",100],
    "3": ["EMHD",50],
    "4": ["FEN", 5]
}

# Exemplo de uso:
doc_ref = []
Lmes_competencia = []
Lnumero_nota = []
Lvalor = []
Ldatas = []

for i in range(4):
    empresa = loja[f"{i+1}"][0]
    max = loja[f"{i+1}"][1]
    
    print(empresa,max)
    coletar_informacoes(Lmes_competencia,Ldatas,Lvalor,Lnumero_nota,doc_ref,empresa,max)

print(doc_ref)
print(Lnumero_nota)
print(Lvalor) 
print(Ldatas)
print(Lmes_competencia)

Lnumero_nota = [item[0] for item in Lnumero_nota]
Lvalor = [item[0] for item in Lvalor]
Ldatas = [item[0] for item in Ldatas]
Lmes_competencia = [item[0] for item in Lmes_competencia]

df = pd.DataFrame({
    'Arquivo PDF': doc_ref,
    'Número da Nota': Lnumero_nota,
    'Valor': Lvalor,
    'Data de Emissão': Ldatas,
    'Mês de Competência': Lmes_competencia
})

df.to_excel('relatorio.xlsx', index=False)



