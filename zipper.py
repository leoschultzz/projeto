import os
import shutil

# Definir os caminhos exatos
PASTA_BASE_PROGRAMA = os.path.dirname(os.path.abspath(__file__))
PASTA_DOCUMENTOS_ORIGEM = "C:\\Users\\Admin\\Downloads\\os Coisa\\confissao\\app\\documentos"
PASTA_DOCXS_GERADOS_ORIGEM = os.path.join(PASTA_BASE_PROGRAMA, "docxs_gerados")
PASTA_PDFS_FORMATADOS_ORIGEM = os.path.join(PASTA_BASE_PROGRAMA, "pdfs_formatados")
# PASTA_TXTS_ORIGEM = os.path.join(PASTA_BASE_PROGRAMA, "txts") # Linha removida/comentada para não incluir txts


def executar():
    """
    Compacta as pastas 'documentos', 'docxs_gerados', e 'pdfs_formatados'
    em um arquivo ZIP chamado 'gerados.zip' na PASTA_BASE_PROGRAMA.
    Se a pasta 'gerados' temporária ou o arquivo 'gerados.zip' já existirem,
    eles serão removidos antes da criação.
    """
    print("—————Iniciando a compactação das pastas de saída em 'gerados.zip'—————\n")

    pastas_a_compactar = [
        PASTA_DOCUMENTOS_ORIGEM,
        PASTA_DOCXS_GERADOS_ORIGEM,
        PASTA_PDFS_FORMATADOS_ORIGEM,
        # PASTA_TXTS_ORIGEM # Linha removida/comentada para não incluir txts
    ]

    # Cria uma pasta temporária para copiar os conteúdos antes de zipar,
    # para garantir que a estrutura dentro do ZIP seja limpa e a mesma para todos os casos.
    pasta_temporaria_gerados = os.path.join(PASTA_BASE_PROGRAMA, "temp_gerados_para_zip")
    if os.path.exists(pasta_temporaria_gerados):
        print(f"  ℹ️ Removendo pasta temporária existente: '{os.path.basename(pasta_temporaria_gerados)}'.")
        shutil.rmtree(pasta_temporaria_gerados)
    os.makedirs(pasta_temporaria_gerados)

    itens_copiados = False
    for pasta_origem in pastas_a_compactar:
        if os.path.exists(pasta_origem) and os.path.isdir(pasta_origem):
            if os.listdir(pasta_origem): # Verifica se a pasta não está vazia
                nome_pasta = os.path.basename(pasta_origem)
                destino_copia = os.path.join(pasta_temporaria_gerados, nome_pasta)
                try:
                    shutil.copytree(pasta_origem, destino_copia)
                    print(f"  ✅ Conteúdo da pasta '{nome_pasta}' copiado para '{os.path.basename(pasta_temporaria_gerados)}'.")
                    itens_copiados = True
                except Exception as e:
                    print(f"  ⚠️ Erro ao copiar a pasta '{nome_pasta}': {e}")
            else:
                print(f"  ℹ️ Pasta '{os.path.basename(pasta_origem)}' está vazia, não será incluída no ZIP.")
        else:
            print(f"  ℹ️ Pasta '{os.path.basename(pasta_origem)}' não encontrada, não será incluída no ZIP.")

    if not itens_copiados:
        print("\n⚠️ Nenhuma pasta com conteúdo foi encontrada para compactar. 'gerados.zip' não será criado.")
        shutil.rmtree(pasta_temporaria_gerados) # Limpa a pasta temp vazia
        print(f"  ✅ Pasta temporária '{os.path.basename(pasta_temporaria_gerados)}' removida.")
        print("\n—————Compactação concluída: Nenhum ZIP criado.—————")
        return

    # Define o nome e caminho do arquivo ZIP de saída
    nome_zip_final = "gerados"
    caminho_zip_final = os.path.join(PASTA_BASE_PROGRAMA, nome_zip_final)

    # Remove o arquivo ZIP existente se houver
    if os.path.exists(f"{caminho_zip_final}.zip"):
        print(f"  ℹ️ Removendo arquivo ZIP existente: '{nome_zip_final}.zip'.")
        os.remove(f"{caminho_zip_final}.zip")

    try:
        # Cria o arquivo ZIP a partir da pasta temporária
        shutil.make_archive(caminho_zip_final, 'zip', root_dir=PASTA_BASE_PROGRAMA, base_dir=os.path.basename(pasta_temporaria_gerados))
        print(f"\n✅ Arquivo '{nome_zip_final}.zip' criado com sucesso em '{PASTA_BASE_PROGRAMA}'.")
    except Exception as e:
        print(f"\n❌ Erro ao criar o arquivo ZIP: {e}")
    finally:
        # Sempre remove a pasta temporária
        if os.path.exists(pasta_temporaria_gerados):
            shutil.rmtree(pasta_temporaria_gerados)
            print(f"  ✅ Pasta temporária '{os.path.basename(pasta_temporaria_gerados)}' removida.")

    print("\n—————Compactação das pastas de saída concluída!—————")

# Para teste direto do zipper.py (opcional)
if __name__ == "__main__":
    print("Este script foi projetado para ser chamado pelo main.py.")
    print("Para um teste, certifique-se de que as pastas de origem existam e contenham arquivos.")
    executar()