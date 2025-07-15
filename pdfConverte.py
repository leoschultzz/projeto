import os
import re
import pandas as pd
import pytesseract
from openpyxl import load_workbook
from pdf2image import convert_from_path

# CONFIGURAÇÕES
pasta_pdfs = 'documentos/'
poppler_path = r'C:\Poppler\Library\bin'
tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Atribui caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# REGEXS
regex_nome = re.compile(r'o lado,\s*(.*?)(?:,|\n)', re.IGNORECASE)
regex_valor = re.compile(r'R\$\s*([\d\.,]+)', re.IGNORECASE)

# Lista final de dados
dados_extraidos = []

# Processa os pdfs, tirando fotos das paginas e depois lendo as fotos
def executar():
    global dados_extraidos
    for arquivo in os.listdir(pasta_pdfs):
        if arquivo.lower().endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_pdfs, arquivo)
            print(f"Processando: {arquivo}")

            try:
                paginas = convert_from_path(caminho_pdf, dpi=300, poppler_path=poppler_path)
                texto_total = ""

                for i, imagem in enumerate(paginas):
                    texto = pytesseract.image_to_string(imagem)
                    texto_total += texto + "\n"

                nome_base = os.path.splitext(arquivo)[0]
                os.makedirs("txts", exist_ok=True)

                nome_match = regex_nome.search(texto_total)
                nome = nome_match.group(1).strip() if nome_match else 'NÃO ENCONTRADO'

                valores = list(regex_valor.finditer(texto_total))

                if valores:
                    inicio = 0
                    for idx, match in enumerate(valores, start=1):
                        fim = match.end()
                        trecho = texto_total[inicio:fim]
                        inicio = fim

                        if len(valores) == 1:
                            caminho_txt = os.path.join("txts", f"{nome_base}.txt")
                        else:
                            caminho_txt = os.path.join("txts", f"{nome_base}_{idx}.txt")

                        with open(caminho_txt, 'w', encoding='utf-8') as f:
                            f.write(trecho.strip())

                        valor_formatado = match.group(1).strip().rstrip(',')
                        dados_extraidos.append({
                            'Arquivo': arquivo,
                            'Nome': nome,
                            'Dívida': valor_formatado,
                            'Número da Dívida': idx
                        })
                else:
                    caminho_txt = os.path.join("txts", f"{nome_base}_1.txt")
                    with open(caminho_txt, 'w', encoding='utf-8') as f:
                        f.write(texto_total.strip())

                    dados_extraidos.append({
                        'Arquivo': arquivo,
                        'Nome': nome,
                        'Dívida': 'NÃO ENCONTRADO',
                        'Número da Dívida': 1
                    })

            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

    # Cria planilha
    df = pd.DataFrame(dados_extraidos)

    # Salvar o DataFrame normalmente
    caminho_planilha = 'resultado_dados.xlsx'
    df.to_excel(caminho_planilha, index=False)

    # Abrir com openpyxl para ajustar largura das colunas
    wb = load_workbook(caminho_planilha)
    ws = wb.active

    # Ajustar largura das colunas
    ws.column_dimensions['A'].width = 30  # Coluna "Arquivo"
    ws.column_dimensions['B'].width = 30  # Coluna "Nome"
    ws.column_dimensions['C'].width = 15  # Coluna "Dívida"
    ws.column_dimensions['D'].width = 20  # Coluna "Número da Dívida"

    # Salvar novamente
    wb.save(caminho_planilha)

    # # Gerar arquivo .txt com os dados
    # with open('resultado_dados.txt', 'w', encoding='utf-8') as txt:
    #     for item in dados_extraidos:
    #         txt.write(f"Arquivo: {item['Arquivo']}\n")
    #         txt.write(f"Nome: {item['Nome']}\n")
    #         txt.write(f"Dívida: R$ {item['Dívida']}\n")
    #         txt.write("-" * 30 + "\n")

    print("✅ Arquivo .txt gerado com sucesso!")

    print("✅ Planilha gerada com colunas ajustadas!")

    print("✅ Planilha gerada com sucesso!")

if __name__ == "__main__":
    executar()
