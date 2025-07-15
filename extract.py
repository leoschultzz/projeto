import os
import shutil
import unicodedata

# Caminho da pasta principal que contém as subpastas com os PDFs
PASTA_PRINCIPAL_EXTRACT = "extract"

def remover_acentos(texto):
    """Função auxiliar para remover acentos do nome do arquivo."""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )

def executar():
    """
    Função principal para organizar os PDFs movendo-os de subpastas
    diretamente para a PASTA_PRINCIPAL_EXTRACT (C:\\Users\\Admin\\Downloads\\extract).
    Remove subpastas vazias após a movimentação.
    """
    print("—————Iniciando organização dos PDFs na pasta 'extract'—————\n")

    if not os.path.exists(PASTA_PRINCIPAL_EXTRACT):
        print(f"⚠️ Pasta de origem '{PASTA_PRINCIPAL_EXTRACT}' não encontrada. Nenhuma operação de extract será realizada.")
        return

    pdfs_movidos = 0
    subpastas_removidas = 0

    # Percorrer todas as subpastas
    for nome_subpasta in os.listdir(PASTA_PRINCIPAL_EXTRACT):
        caminho_subpasta = os.path.join(PASTA_PRINCIPAL_EXTRACT, nome_subpasta)

        # Verificar se é pasta
        if os.path.isdir(caminho_subpasta):
            print(f"  Processando subpasta: {nome_subpasta}")
            # Listar arquivos dentro da subpasta
            for arquivo in os.listdir(caminho_subpasta):
                if arquivo.lower().endswith(".pdf"):
                    nome_sem_acentos = remover_acentos(arquivo)
                    origem = os.path.join(caminho_subpasta, arquivo)
                    destino = os.path.join(PASTA_PRINCIPAL_EXTRACT, nome_sem_acentos)

                    # Evitar sobrescrita renomeando se necessário
                    if os.path.exists(destino):
                        base, ext = os.path.splitext(nome_sem_acentos)
                        contador = 1
                        while os.path.exists(destino):
                            novo_nome = f"{base}_{contador}{ext}"
                            destino = os.path.join(PASTA_PRINCIPAL_EXTRACT, novo_nome)
                            contador += 1
                        print(f"    ℹ️ Arquivo '{arquivo}' já existe em '{os.path.basename(PASTA_PRINCIPAL_EXTRACT)}'. Renomeado para '{os.path.basename(destino)}'.")

                    try:
                        shutil.move(origem, destino)
                        print(f"    ✅ Movido: '{arquivo}' para '{os.path.basename(PASTA_PRINCIPAL_EXTRACT)}'.")
                        pdfs_movidos += 1
                    except Exception as e:
                        print(f"    ⚠️ Erro ao mover '{arquivo}' de '{nome_subpasta}': {e}")

            # Remover subpasta se estiver vazia
            try:
                if not os.listdir(caminho_subpasta):
                    os.rmdir(caminho_subpasta)
                    print(f"  ✅ Subpasta '{nome_subpasta}' removida pois estava vazia.")
                    subpastas_removidas += 1
                else:
                    print(f"  ℹ️ Subpasta '{nome_subpasta}' não está vazia. Conteúdo restante: {os.listdir(caminho_subpasta)}")
            except Exception as e:
                print(f"  ⚠️ Erro ao tentar remover subpasta '{nome_subpasta}': {e}")

    if pdfs_movidos > 0:
        print(f"\n✅ Total de {pdfs_movidos} PDF(s) organizados na pasta '{os.path.basename(PASTA_PRINCIPAL_EXTRACT)}'.")
    else:
        print("\nℹ️ Nenhun PDF encontrado em subpastas para organizar.")
    print("\n—————Organização da pasta 'extract' concluída!—————")

# Exemplo de uso se este script fosse executado diretamente (para testes)
if __name__ == "__main__":
    executar_extract()