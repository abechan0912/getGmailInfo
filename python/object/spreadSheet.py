import logging


# 採用者情報管理クラスを定義
class SpreadSheet:

    # コンストラクタを定義
    def __init__(self, work_book, print_sheet_name) -> None:
        # ワークブック
        self.work_book = work_book

        # 貼付用ワークシート名
        self.print_sheet_name = print_sheet_name

        # logger設定
        self.log = logging.getLogger('main').getChild("SpreadSheet")

    # データ貼り付け
    def printMessageList(self, message_list) -> None:
        try:
            # 空文字削除
            message_list = self.deleteEmptyString(message_list)
        except Exception as e:
            self.log.exception(e)
            raise e
        
        try:
            # 貼り付ける前にシートをクリア
            self.work_book.worksheet(self.print_sheet_name).clear()
            self.work_book.values_append(self.print_sheet_name, {"valueInputOption" : "USER_ENTERED"}, {"values" : message_list})
        except Exception as e:
            self.log.exception(e)
            raise e

    # 空文字削除
    def deleteEmptyString(self, message_list) -> None:
        try:
            self.log.debug("空文字削除 - 開始")
            for i in range(len(message_list)):
                table = str.maketrans({'\u3000': '',' ': '','\t': ''})
                message_list[i][0] = message_list[i][0].translate(table)
            self.log.debug("空文字削除 - 完了")
            return message_list
        except Exception as e:
            self.log.debug("空文字削除 - 失敗")
            raise e