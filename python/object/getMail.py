import logging
import base64

# Gmailクラスを定義
class GetMail:

    def __init__(self) -> None:
        # logger設定
        self.log = logging.getLogger('main').getChild("GetMail")

    # Gmailのメッセージを取得
    def mail_get_messages_body(self, service, labelIdsValue):
        try:
            message_list = []
        
            # メッセージの一覧を取得
            messages = service.users().messages()
            msg_list = messages.list(userId='me', labelIds=labelIdsValue).execute()

            # メッセージの件数を取得
            result_size = msg_list['resultSizeEstimate']
            self.log.debug("取得件数：" + str(result_size) + "件")

            # 取得したメールの件数が０件の場合は終了
            if(0 == result_size):
                self.log.debug("新規メールはありませんでした")
                return result_size
            
            message_list = self.mail_get_messages_body_content(messages, msg_list, message_list)
            
            
            while 'nextPageToken' in msg_list:
                page_token = msg_list['nextPageToken']
                msg_list = messages.list(userId='me', labelIds=labelIdsValue, pageToken=page_token).execute()
                message_list = self.mail_get_messages_body_content(messages, msg_list, message_list)
            
            return message_list
        except Exception as e:
            self.log.exception(e)
            raise e

    # Gmailのコンテンツを取得
    def mail_get_messages_body_content(self, messages, msg_list, message_list):
     
        # 各message内容確認
        for message_id in msg_list["messages"]:
            try:
                
                # 各メッセージ詳細
                msg = messages.get(userId='me', id=message_id['id']).execute()
                            
                # 本文
                if(msg["payload"]["body"]["size"]!=0):
                    decoded_bytes = base64.urlsafe_b64decode(
                        msg["payload"]["body"]["data"])
                    decoded_message = decoded_bytes.decode("UTF-8")
                else:
                    #メールによっては"parts"属性の中に本文がある場合もある
                    decoded_bytes = base64.urlsafe_b64decode(
                        msg["payload"]["parts"][0]["body"]["data"])
                    decoded_message = decoded_bytes.decode("UTF-8")
                
                # 氏名
                name = self.find_name(decoded_message)

                # 内容
                content = self.find_content(decoded_message)

                # 担当者
                manager = self.find_manager(decoded_message)

                # メニュー
                menu = self.find_menu(decoded_message)

                # 時間
                time = self.find_time(decoded_message)

                # 日程
                schedule = self.find_schedule(decoded_message)
        
                message_list.append([name, content, manager, menu, time, schedule])
            except Exception as e:
                self.log.exception(e)
                
        return message_list
    
    # 本文の中から氏名を取得
    def find_name(self, body):
        try:
            target = '氏名： '
            target_end = '<br>内容'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 4 : idx_end]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e

    # 本文の中から内容を取得
    def find_content(self, body):
        try:
            target = '内容： '
            target_end = '<br>担当者'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 4 : idx_end]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e

    # 本文の中から担当者を取得
    def find_manager(self, body):
        try:
            target = '担当者： '
            target_end = '<br>メニュー'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 5 : idx_end]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e
    
    # 本文の中からメニューを取得
    def find_menu(self, body):
        try:
            target = 'メニュー: '
            target_end = '<br>日程'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 6 : idx_end]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e

    # 本文の中から時間を取得
    def find_time(self, body):
        try:
            target = '時間： '
            target_end = '分'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 4 : idx_end + 1]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e
    
    # 本文の中から日程を取得
    def find_schedule(self, body):
        try:
            target = '日程： '
            target_end = '<br>時間'
            idx = body.find(target)
            idx_end = body.find(target_end)
            r = body[idx + 4 : idx_end]
            return r
        except Exception as e:
            self.log.exception(e)
            raise e
