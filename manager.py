import os
import shutil
from datetime import datetime
from tqdm import tqdm

# Caminhos exatos (invariados)
PASTA_BASE_PROGRAMA = os.path.dirname(os.path.abspath(__file__))
PASTA_EXTRACT_ORIGEM = "extract"
PASTA_DOCUMENTOS_DESTINO = "documentos"

# Lista de arquivos e pastas para serem gerenciados (invariado)
ARQUIVOS_EXCEL_PARA_DELETAR = [
    os.path.join(PASTA_BASE_PROGRAMA, "dados_atualizados.xlsx"),
    os.path.join(PASTA_BASE_PROGRAMA, "resultado_dados.xlsx")
]
ARQUIVOS_TEMP_PARA_DELETAR = [
    os.path.join(PASTA_BASE_PROGRAMA, "temp_pag.pdf.png")
]
PASTAS_PARA_ARQUIVAR_E_LIMPAR = [
    os.path.join(PASTA_BASE_PROGRAMA, "docxs_gerados"),
    os.path.join(PASTA_BASE_PROGRAMA, "pdfs_formatados"),
    os.path.join(PASTA_BASE_PROGRAMA, "txts"),
    PASTA_DOCUMENTOS_DESTINO
]
PASTA_LOGS = os.path.join(PASTA_BASE_PROGRAMA, "logs")

# Garante que a pasta logs exista no início do manager
os.makedirs(PASTA_LOGS, exist_ok=True)
# Garante que a pasta documentos de destino exista (será limpa depois)
os.makedirs(PASTA_DOCUMENTOS_DESTINO, exist_ok=True)


def _deletar_arquivo_se_existe(arquivo_path, nome_amigavel):
    """Deleta um arquivo se ele existir, informando ao usuário."""
    if os.path.exists(arquivo_path):
        try:
            os.remove(arquivo_path)
            print(f"  ✅ Arquivo '{nome_amigavel}' deletado com sucesso.")
        except OSError as e:
            print(f"  ⚠️ Erro ao deletar o arquivo '{nome_amigavel}': {e}")
    else:
        print(f"  ℹ️ Arquivo '{nome_amigavel}' não encontrado, pulando deleção.")


def _arquivar_e_limpar_pasta(pasta_path):
    """
    Move o conteúdo de uma pasta para uma subpasta datada dentro da pasta de logs,
    mantendo a estrutura da pasta original, e depois limpa a pasta original.
    """
    if not os.path.exists(pasta_path) or not os.path.isdir(pasta_path):
        print(f"  ℹ️ Pasta '{os.path.basename(pasta_path)}' não encontrada para arquivar/limpar.")
        return

    if not os.listdir(pasta_path):
        print(f"  ℹ️ Pasta '{os.path.basename(pasta_path)}' está vazia, nada para arquivar.")
        return

    # Cria a pasta de log com data e hora atual para esta execução completa
    # E dentro dela, uma subpasta com o nome original da pasta que está sendo arquivada
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_log_execucao = os.path.join(PASTA_LOGS, f"execucao_{timestamp}") # Pasta para esta rodada de logs
    os.makedirs(pasta_log_execucao, exist_ok=True)

    # O destino dos itens será uma subpasta dentro da pasta de execução
    # com o nome da pasta original (ex: logs/execucao_2025-06-16_13-45-00/docxs_gerados/)
    pasta_log_destino_especifico = os.path.join(pasta_log_execucao, os.path.basename(pasta_path))
    os.makedirs(pasta_log_destino_especifico, exist_ok=True)

    print(f"  🔄 Arquivando conteúdo de '{os.path.basename(pasta_path)}' para '{os.path.basename(pasta_log_execucao)}/{os.path.basename(pasta_path)}'...")
    movidos_count = 0
    for item in os.listdir(pasta_path):
        origem_item = os.path.join(pasta_path, item)
        destino_item = os.path.join(pasta_log_destino_especifico, item) # Move para a subpasta específica
        try:
            shutil.move(origem_item, destino_item)
            movidos_count += 1
        except shutil.Error as e:
            print(f"  ⚠️ Erro ao mover '{item}' para logs: {e}")
        except OSError as e:
            print(f"  ⚠️ Erro do sistema ao mover '{item}' para logs: {e}")

    if movidos_count > 0:
        print(f"  ✅ {movidos_count} item(ns) de '{os.path.basename(pasta_path)}' arquivado(s).")
    else:
        print(f"  ℹ️ Nenhum item arquivado de '{os.path.basename(pasta_path)}'.")

    # Após mover para o log, garante que a pasta original esteja vazia ou recria-a
    try:
        if os.listdir(pasta_path):
            print(
                f"  ⚠️ Alguns itens permaneceram em '{os.path.basename(pasta_path)}' após arquivamento. Tentando remover diretamente.")
            shutil.rmtree(pasta_path)
        else:
            os.rmdir(pasta_path)
        print(f"  ✅ Pasta original '{os.path.basename(pasta_path)}' limpa.")
        # Recria a pasta se ela for uma das que precisam existir para o fluxo
        if pasta_path in [
            os.path.join(PASTA_BASE_PROGRAMA, "docxs_gerados"),
            os.path.join(PASTA_BASE_PROGRAMA, "pdfs_formatados"),
            os.path.join(PASTA_BASE_PROGRAMA, "txts"),
            PASTA_DOCUMENTOS_DESTINO
        ]:
            os.makedirs(pasta_path, exist_ok=True)
            print(f"  ✅ Pasta '{os.path.basename(pasta_path)}' recriada para uso futuro.")

    except OSError as e:
        print(f"  ⚠️ Erro ao limpar/recriar a pasta '{os.path.basename(pasta_path)}': {e}")


def _mover_extract_para_documentos_final():
    """
    Move TODOS os arquivos e subpastas da pasta 'extract' (PASTA_EXTRACT_ORIGEM)
    para a pasta 'documentos' (PASTA_DOCUMENTOS_DESTINO no app).
    Esta é a última etapa de movimentação, com barra de progresso.
    """
    print("\n—————Iniciando a operação de RECORTAR e COLAR da pasta 'extract' para 'documentos'—————\n")

    # Debugging Critical (invariado)
    print(f"  DEBUG: Verificando o conteúdo da pasta de origem: '{PASTA_EXTRACT_ORIGEM}'")
    try:
        items_in_extract = os.listdir(PASTA_EXTRACT_ORIGEM)
        if not items_in_extract:
            print(f"  DEBUG: A pasta '{PASTA_EXTRACT_ORIGEM}' está VAZIA no momento da movimentação. NADA PARA MOVER.")
            print("\n—————Movimentação final concluída: PASTA DE ORIGEM JÁ ESTAVA VAZIA.—————")
            return
        else:
            print(f"  DEBUG: Conteúdo encontrado em '{PASTA_EXTRACT_ORIGEM}': {items_in_extract}")
    except FileNotFoundError:
        print(f"  DEBUG: A pasta de origem '{PASTA_EXTRACT_ORIGEM}' NÃO FOI ENCONTRADA. IMPOSSIBILITADO DE MOVER.")
        print("\n—————Movimentação final concluída: PASTA DE ORIGEM NÃO ENCONTRADA.—————")
        return
    except Exception as e:
        print(f"  DEBUG: Erro inesperado ao listar o conteúdo de '{PASTA_EXTRACT_ORIGEM}': {e}")
        print("\n—————Movimentação final concluída: ERRO AO LISTAR PASTA DE ORIGEM.—————")
        return

    # Garante que a pasta de destino exista (invariado)
    os.makedirs(PASTA_DOCUMENTOS_DESTINO, exist_ok=True)
    print(f"  Pasta de destino '{PASTA_DOCUMENTOS_DESTINO}' verificada/criada.")

    arquivos_movidos = 0
    # Envolver o loop com tqdm para a barra de progresso (invariado)
    for item in tqdm(items_in_extract, desc="Recortando arquivos", unit="item"):
        origem = os.path.join(PASTA_EXTRACT_ORIGEM, item)
        destino = os.path.join(PASTA_DOCUMENTOS_DESTINO, item)

        try:
            # Verifica se o item já existe no destino (invariado)
            if os.path.exists(destino):
                if os.path.isfile(origem):
                    base, ext = os.path.splitext(item)
                    contador = 1
                    while os.path.exists(destino):
                        novo_nome = f"{base}_{contador}{ext}"
                        destino = os.path.join(PASTA_DOCUMENTOS_DESTINO, novo_nome)
                        contador += 1
                    tqdm.write(
                        f"    ℹ️ Item '{item}' (arquivo) já existe no destino. Renomeado para '{os.path.basename(destino)}' para evitar sobrescrita.")
                elif os.path.isdir(origem):
                    tqdm.write(
                        f"    ⚠️ Item '{item}' (pasta) já existe em '{os.path.basename(PASTA_DOCUMENTOS_DESTINO)}'. Não movido para evitar fusão/sobrescrita complexa.")
                    continue

            # Executa a movimentação (recorte) (invariado)
            shutil.move(origem, destino)
            tqdm.write(
                f"    ✅ RECORTADO: '{item}' de '{os.path.basename(PASTA_EXTRACT_ORIGEM)}' e COLADO em '{os.path.basename(PASTA_DOCUMENTOS_DESTINO)}'.")
            arquivos_movidos += 1

        except shutil.Error as e:
            tqdm.write(f"    ❌ ERRO DE RECORTE (shutil): '{item}'. Detalhes: {e}")
        except OSError as e:
            tqdm.write(f"    ❌ ERRO DE RECORTE (OS): '{item}'. Detalhes: {e}")
        except Exception as e:
            tqdm.write(f"    ❌ ERRO INESPERADO: '{item}'. Detalhes: {e}")

    # Print final e CLARO sobre o resultado da movimentação (invariado)
    print("\n----------------------------------------------------------------------------------------------------")
    if arquivos_movidos > 0:
        print(
            f"🎉 SUCESSO FINAL: {arquivos_movidos} item(ns) foram RECORTADOS de '{PASTA_EXTRACT_ORIGEM}' e COLADOS em '{PASTA_DOCUMENTOS_DESTINO}'.")
    else:
        print(
            f"🔴 ATENÇÃO FINAL: NENHUM item foi recortado da pasta '{PASTA_EXTRACT_ORIGEM}' para '{PASTA_DOCUMENTOS_DESTINO}'.")
        print("    Isso pode ocorrer se a pasta de origem já estava vazia ou se ocorreram erros durante a movimentação.")
    print("----------------------------------------------------------------------------------------------------")

    # Mensagem final, sem tentar remover a pasta (invariado)
    print(f"\n  ℹ️ A pasta '{os.path.basename(PASTA_EXTRACT_ORIGEM)}' não será removida, conforme solicitado.")
    print("\n—————Movimentação final concluída!—————")


def executar():
    """
    Função principal que será chamada pela main.py para executar
    as operações de gerenciamento de arquivos (limpeza e log),
    incluindo a movimentação final como "Parte 4".
    """
    print("—————Iniciando gerenciamento de arquivos (limpeza e log)—————\n")

    print("  1. Deletando arquivos Excel temporários...")
    for arquivo_path in ARQUIVOS_EXCEL_PARA_DELETAR:
        _deletar_arquivo_se_existe(arquivo_path, os.path.basename(arquivo_path))

    print("\n  2. Deletando arquivos temporários restantes...")
    for arquivo_path in ARQUIVOS_TEMP_PARA_DELETAR:
        _deletar_arquivo_se_existe(arquivo_path, os.path.basename(arquivo_path))

    print("\n  3. Arquivando e limpando pastas de resultados anteriores...")
    for pasta_path in PASTAS_PARA_ARQUIVAR_E_LIMPAR:
        _arquivar_e_limpar_pasta(pasta_path)

    print("\n  4. Recortando arquivos da pasta 'extract' para 'documentos' (movimentação final)...")
    _mover_extract_para_documentos_final()

    print("\n—————Gerenciamento de arquivos concluído (todas as etapas)!—————")