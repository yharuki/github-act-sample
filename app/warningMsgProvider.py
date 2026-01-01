# coding: utf-8

__author__ = "Yuji Haruki (modifier) / Kohei, Watanabe <kohei.watanabe3@brother.co.jp> (original)"
__version__ = "2.1.0"
__date__ = "5 June 2024"

from enum import Enum


class MainAppStatus(Enum):
    WARNING_CODE_1 = 1
    WARNING_CODE_2 = 2
    WARNING_CODE_3 = 3
    ERROR_CODE_1 = 101
    ERROR_CODE_2 = 102
    ERROR_CODE_3 = 103
    ERROR_CODE_4 = 104

class ExOpStatus(Enum):
    WARNING_CODE_1 = 1001
    WARNING_CODE_2 = 1002
    WARNING_CODE_3 = 1003
    ERROR_CODE_1 = 1101
    ERROR_CODE_2 = 1102
    ERROR_CODE_3 = 1103

class MdOpStatus(Enum):
    WARNING_CODE_1 = 10001
    WARNING_CODE_2 = 10002
    WARNING_CODE_3 = 10003
    WARNING_CODE_4 = 10004
    WARNING_CODE_5 = 10005
    WARNING_CODE_6 = 10006
    ERROR_CODE_1 = 10101
    ERROR_CODE_2 = 10102
    ERROR_CODE_3 = 10103
    ERROR_CODE_4 = 10104
    ERROR_CODE_5 = 10105
    ERROR_CODE_6 = 10106
    ERROR_CODE_7 = 10107
    ERROR_CODE_8 = 10108
    ERROR_CODE_9 = 10109


class WarningMsgProvider:
    def __init__(self):
        self.warning_target_fp = ""
        self.error_target_fp = ""

    def setTargetFP(self, file_path):
        self.warning_target_fp = file_path
        self.error_target_fp = file_path

    def buildMsg(self, code, line_num="", arg1="", arg2=""):
        msg = ""

        if self.warning_target_fp:
            msg += "\n"
            msg += "→ " + self.warning_target_fp + " について以下を確認してください"+ "\n"
            msg += "\n"
            self.warning_target_fp = ""

        ### MdToExcel.py 関連の警告とエラー

        # MdToExcel.py の警告メッセージは
        # ユーザーとの対話に使用ものが主であるため
        # コードの可読性を維持する目的で
        # 当該モジュールに直接記述する
        if code == MainAppStatus.WARNING_CODE_1.value:
            pass # Reserved
        elif code == MainAppStatus.WARNING_CODE_2.value:
            pass # Reserved
        elif code == MainAppStatus.WARNING_CODE_3.value:
            pass # Reserved
        elif code == MainAppStatus.ERROR_CODE_1.value:
            msg += "【 エラー 】" + "\n"
            msg += "設定ファイル（config.yaml）が見つかりません" + "\n"
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MainAppStatus.ERROR_CODE_2.value:
            msg += "【 エラー 】" + "\n"
            msg += "指定できるファイルの拡張子は以下のいずれかのみです" + "\n"
            msg += "（双方向の変換は同時にできません）" + "\n"
            msg += "\n"
            msg += "・Markdown -> Excel 変換：md" + "\n"
            msg += "・Excel -> Markdown 変換：xlsm" + "\n"
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MainAppStatus.ERROR_CODE_3.value:
            pass # Reserved


        ### excel_operator.py 関連の警告とエラー

        if code == ExOpStatus.WARNING_CODE_1.value:
            pass # Reserved
        elif code == ExOpStatus.WARNING_CODE_2.value:
            pass # Reserved
        elif code == ExOpStatus.WARNING_CODE_3.value:
            pass # Reserved
        elif code == ExOpStatus.ERROR_CODE_1.value:
            msg += "【 エラー 】" + "\n"
            msg += self.error_target_fp + " に保存できません" + "\n"
            msg += "ファイルを開いていませんか？" + "\n"
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == ExOpStatus.ERROR_CODE_2.value:
            msg += "【 エラー 】" + "\n"
            msg += "予期しないエラーが発生しました（管理者に報告してください）" + "\n"
            msg += "\n"
        elif code == ExOpStatus.ERROR_CODE_3.value:
            pass # Reserved


        ### markdown_operator.py 関連の警告とエラー

        if code == MdOpStatus.WARNING_CODE_1.value:
            pass # reserved
        elif code == MdOpStatus.WARNING_CODE_2.value:
            msg += line_num + "行目: 無効な記述があります" + "\n"
            msg += "            " + arg1
        elif code == MdOpStatus.WARNING_CODE_3.value:
            msg += line_num + "行目: 無効な記述があります" + "\n"
            msg += "            " + "セル番号 " + arg1 + ": " + arg2
        elif code == MdOpStatus.WARNING_CODE_4.value:
            msg += "            " + "セル番号 " + arg1 + ": " + arg2
        elif code == MdOpStatus.WARNING_CODE_5.value:
            msg += line_num + "行目: テスト観点のレベル（# の数）が間違っていませんか？" + "\n"
        elif code == MdOpStatus.ERROR_CODE_1.value:
            msg += "【 エラー 】" + "\n"
            msg += "Markdownファイル（.md）が見つかりません" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_2.value:
            msg += "【 エラー 】" + "\n"
            msg += "Markdownファイルの名前が不適切です" + "\n"
            msg += "次の点を確認して修正してください" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += "・ファイル名が 31 文字以内であること" + "\n"
            msg += "・次の使用できない文字が含まれていないこと" + "\n"
            msg += (
                "　コロン(:)、円記号(\)、スラッシュ(/)、疑問符(?)、アスタリスク(*)、左角かっこ([)、右角かっこ(])"
                + "\n"
            )
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_3.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト実施環境の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += "・同じ名前のテスト実施環境が存在していませんか？" + "\n"
            msg += "・記述したエリアが ``` で囲まれていますか？" + "\n"
            msg += (
                "・テスト実施環境の名前が # （テスト観点の記号）で始まっていませんか？"
                + "\n"
            )
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_4.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト項目の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += "・ヘッダー行（下記）を正しく認識できませんでした" + "\n"
            msg += "\n"
            msg += "　1,2,3,4,5,6,番号,環境,準備,手順,確認,備考・・" + "\n"
            msg += "\n"
            msg += "処理を中止しました\n" + "\n"
        elif code == MdOpStatus.ERROR_CODE_5.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト項目の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += (
                line_num
                + "行目: テスト観点、またはテスト項目の番号が記載されていません"
                + "\n"
            )
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_6.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト項目の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += line_num + "行目: 「手順」または「確認」が空白になっています" + "\n"
            msg += "            " + "セル番号 " + arg1 + ": " + arg2 + "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_7.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト項目の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += (
                line_num
                + "行目: 改行記述の（親となる）先頭行は、リスト or 番号付きリストである必要があります"
                + "\n"
            )
            msg += "            " + "セル番号 " + arg1 + ": " + arg2 + "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_8.value:
            msg += "【 エラー 】" + "\n"
            msg += "テストのタイトル行がありません" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += "・イコールの記号で定義されるタイトル行が必要です" + "\n"
            msg += "\n"
            msg += "処理を中止しました" + "\n"
        elif code == MdOpStatus.ERROR_CODE_9.value:
            msg += "【 エラー 】" + "\n"
            msg += "テスト項目の記述に誤りがあります" + "\n"
            msg += "\n"
            msg += "ファイル名：" + self.error_target_fp + "\n"
            msg += "\n"
            msg += (
                line_num
                + "行目: この行の直前に「手順」もしくは「確認」が空白の項目があります"
            )
            msg += "\n"
            msg += "処理を中止しました" + "\n"
            pass
        return msg
