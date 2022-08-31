import logging
from logging import config
import yaml
import os.path
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys
from object import getMail
from object import spreadSheet

def gmail_init():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(tokenPath):
        creds = Credentials.from_authorized_user_file(tokenPath, gmailScope)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
               credentialsPath, gmailScope)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(tokenPath, 'w') as token:
            token.write(creds.to_json())
 
    service = build('gmail', 'v1', credentials=creds)
    return service

# ラベルID取得
def gmail_display_label(service):
    log.debug("ラベルID取得 - 開始")
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        log.debug("ラベルが見つかりませんでした。")
    else:
        for label in labels:
            log.debug(label)
    log.debug("ラベルID取得 - 完了")
    sys.exit()

# logger設定
config.fileConfig("C:\\programing\\python\\rpa\\getGmailInfo\\resource\\config\\logging.conf")
log = logging.getLogger('main')

# パラメータファイル
with open("C:\\programing\\python\\rpa\\getGmailInfo\\resource\\config\\parameters.yaml", encoding="utf-8") as file:
    para = yaml.safe_load(file)

    # Google Gmail 認証
    try:
        log.debug("Google Gmail 認証 - 開始")
        # API設定
        gmailScope = ['https://www.googleapis.com/auth/gmail.readonly']

        # token.jsonを定義（実体ファイルは最初はなくてもOK。なければ作る処理を書くので）
        tokenPath = para["COMMON"]["TOKEN_PATH"]

        # 認証情報設定
        credentialsPath = para["COMMON"]["CREDENTIALS_PATH"]

        # Gmail APIの初期化実行
        service = gmail_init()

        log.debug("Google Gmail 認証 - 完了")
    except Exception as e:
        log.error("Google Gmail 認証 - 失敗")
        log.exception(e)
        sys.exit()

    # ラベルID取得(通常は実行しない)
    # gmail_display_label(service)

    # Gmail
    gmail = getMail.GetMail()

    # Gmailのメッセージを取得
    try:
        log.debug("Gmailのメッセージを取得 - 開始")
        message_list = gmail.mail_get_messages_body(service, para["G_MAIL"]["LABEL_ID"])

        if(0 == message_list):
            sys.exit()

        log.debug("Gmailのメッセージを取得 - 完了")
    except Exception as e:
        log.error("Gmailのメッセージを取得 - 失敗")
        sys.exit()

    # Google SpreadSheets認証
    try:
        log.debug("Google SpreadSheets認証 - 開始")
        # API設定
        sheetsScope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

        # 認証情報設定
        credentials = ServiceAccountCredentials.from_json_keyfile_name(para["COMMON"]["CERTIFICATION_KEY"], sheetsScope)
        gc:any = gspread.authorize(credentials)

        # スプレッドシート取得
        work_book = gc.open_by_key(para["COMMON"]["SPREADSHEET_KEY"])
        log.debug("Google SpreadSheets認証 - 完了")
    except Exception as e:
        log.error("Google SpreadSheets認証 - 失敗")
        log.exception(e)
        sys.exit()

    # SpreadSheet
    spread_sheets = spreadSheet.SpreadSheet(
        work_book, 
        para["SPREAD_SHEET"]["PRINT_SHEET_NAME"]
    )

    # メール内容貼付
    try:
        log.debug("メール内容貼付 - 開始")
        spread_sheets.printMessageList(message_list)
        log.debug("メール内容貼付 - 完了")
    except Exception as e:
        log.error("メール内容貼付 - 失敗")
        sys.exit()