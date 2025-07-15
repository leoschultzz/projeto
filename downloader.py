import os
import gdown
import re

class GoogleDriveDownloader:
    def __init__(self):
        pass

    def _extract_folder_id(self, google_drive_link): # Extrai o ID da pasta de um hyperlink do Google Drive.
        match = re.search(r'/folders/([a-zA-Z0-9_-]+)', google_drive_link)
        if match:
            return match.group(1)
        return None

    def download_folder_contents(self, google_drive_link_or_id, destination_folder): # Baixa todos os arquivos da pasta do Google Drive.
        folder_id = self._extract_folder_id(google_drive_link_or_id)
        if not folder_id:
            folder_id = google_drive_link_or_id

        os.makedirs(destination_folder, exist_ok=True)
        print(f"\nIniciando download dos arquivos da pasta do Google Drive para '{destination_folder}' usando gdown...")

        try:
            gdown.download_folder(id=folder_id, output=destination_folder, quiet=False, use_cookies=False)
            print(f"\n✅ Download da pasta do Google Drive (ID: {folder_id}) concluído com sucesso!")

        except Exception as e:
            print(f"❌ Ocorreu um erro ao baixar a pasta com gdown: {e}")
            print("Verifique se o ID ou link da pasta está correto e se a pasta está compartilhada publicamente (com a opção 'Qualquer pessoa com o link pode visualizar').")



def executar(google_drive_link_or_id, destination_folder): # Função principal para executar o download da pasta do Google Drive.
    # Instancia a classe GoogleDriveDownloader
    downloader = GoogleDriveDownloader()
    # Chama o método para baixar o conteúdo
    downloader.download_folder_contents(google_drive_link_or_id, destination_folder)