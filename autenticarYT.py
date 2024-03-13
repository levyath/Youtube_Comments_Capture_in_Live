from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import openpyxl
import time
import pytz
from datetime import datetime

class Api_Youtube():


    def __init__(self):

        # Defina as suas credenciais aqui
        API_KEY = "sua chave"
        VIDEO_ID = input("Entre com a URL da Live: ")

        # Extrair o ID do vídeo da URL
        VIDEO_ID = VIDEO_ID.split('=')[-1]
        
        # Crie uma instância do serviço da API do YouTube
        youtube = build('youtube', 'v3', developerKey=API_KEY)

        # Faça uma solicitação para obter o ID do live chat associado ao vídeo
        response = youtube.videos().list(
            part='liveStreamingDetails',
            id=VIDEO_ID
        ).execute()

        # Extraia o ID do live chat
        live_chat_id = response['items'][0]['liveStreamingDetails']['activeLiveChatId']

        # Carregar ou criar um arquivo Excel para armazenar os comentários
        try:
            workbook = openpyxl.load_workbook('youtube_comments.xlsx')
        except FileNotFoundError:
            workbook = openpyxl.Workbook()

        # Selecionar a primeira planilha no arquivo (se não houver, criar uma)
        sheet = workbook.active
        if sheet.title != 'Comments':
            sheet.title = 'Comments'
            sheet.append(['Comment_ID', 'Published_At', 'Comment_Text'])

        # Definir uma lista para armazenar IDs de comentários já vistos
        seen_comments = []

        # Loop infinito para capturar novos comentários em tempo real
        while True:
            # Use o ID do live chat para capturar os comentários em tempo real
            response = youtube.liveChatMessages().list(
                liveChatId=live_chat_id,
                part='snippet',
                maxResults=200  # Defina o número máximo de comentários a serem retornados
            ).execute()

            # Iterar sobre os comentários
            for item in response['items']:
                try:
                    comment_id = item['id']
                    comment_text = item['snippet']['textMessageDetails']['messageText']
                    comment_published_at = item['snippet']['publishedAt']

                    # Verificar se o ID do comentário já foi visto
                    if comment_id not in seen_comments:
                        # Adicionar o ID do comentário à lista de comentários vistos
                        seen_comments.append(comment_id)

                        # Converte o horário para o fuso que conhecemos (Por padrão chega em formato ISO 8601)
                        # Converta o horário UTC para o fuso horário local
                        utc_time = datetime.strptime(comment_published_at, "%Y-%m-%dT%H:%M:%S.%f%z")
                        local_timezone = pytz.timezone('America/Sao_Paulo')  # Defina o fuso horário local desejado
                        local_time = utc_time.astimezone(local_timezone)

                        # Adicionar os detalhes do comentário ao arquivo Excel
                        sheet.append([comment_id, local_time.strftime("%Y-%m-%d %H:%M:%S"), comment_text])
                except KeyError:
                    # Se a chave 'textMessageDetails' estiver ausente, o item não é um comentário de texto
                    # Você pode lidar com isso aqui, se desejar
                    pass

            # Aguardar alguns segundos antes de verificar novamente por novos comentários
            # Aguardar 3 segundos antes de verificar novamente
            time.sleep(3)

            # Salvar o arquivo Excel após cada iteração para garantir que os dados não sejam perdidos em caso de falha
            workbook.save('youtube_comments.xlsx')
