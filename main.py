from pdfConverte import executar as executar_pdfConverte
from pdfGerar import executar as executar_pdfGerar
from confissaoGerar import executar as executar_confissaoGerar
from extract import executar as executar_extract
from manager import executar as executar_manager
from zipper import executar as executar_zipper
from downloader import executar as executar_downloader

if __name__ == "__main__":
    # Solicita o valor do IGP-M
    # while True:
    #     try:
    #         igpm_input = input("\nDigite o valor do IGP-M (ex: 0.068634): ").replace(',', '.').strip()
    #         igpm = float(igpm_input)
    #         break
    #     except ValueError:
    #         print("⚠️ Valor inválido. Tente novamente usando ponto como separador decimal.")

    igpm = 0.0617006

    # Pergunta se deseja censurar(função antiga)
    # while True:
    #     censura_input = input("\nDeseja aplicar censura nos nomes? (S/N): ").strip().upper()
    #     if censura_input == 'S':
    #         censurar = True
    #         break
    #     elif censura_input == 'N':
    #         censurar = False
    #         break
    #     else:
    #         print("⚠️ Responda apenas com S ou N.")

    # Solicita o hyperlink para a pasta compartilhada no drive
    # while True:
    #     drive_folder_link = input(
    #         "\nDigite o LINK COMPLETO ou o ID da pasta do Google Drive para download (certifique-se de que está pública): ").strip()
    #     if drive_folder_link:
    #         break
    #     else:
    #         print("⚠️ O link ou ID da pasta do Google Drive não pode ser vazio. Por favor, tente novamente.")

    drive_folder_link = 'https://drive.google.com/drive/folders/1LH-VIWeuDF23TWjht555iVnJyUyjDow-?hl=pt-br'
    destination_folder = 'extract'

    print("\n\n—————Pré-Etapas—————\n\n")
    executar_downloader(drive_folder_link, destination_folder)
    executar_extract()
    executar_manager()

    print("\n\n—————Etapa 1: Convertendo PDFs—————\n\n")
    executar_pdfConverte()

    print("\n\n—————Etapa 2: Gerando PDFs atualizados—————\n\n")
    executar_pdfGerar(igpm)

    print("\n\n—————Etapa 3: Gerando documentos .docx—————\n\n")
    executar_confissaoGerar(False) # Passando false porque está sem o input para agilizar a execução

    print("\n\n—————Etapa 4: Gerando documentos .docx—————\n\n")
    executar_zipper()

    print("\n🎉 Processo finalizado com sucesso!")