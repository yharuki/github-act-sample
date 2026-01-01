# coding: utf-8

__author__ = "Yuji Haruki (modifier) / Kohei, Watanabe <kohei.watanabe3@brother.co.jp> (original)"
__version__ = "2.1.0"
__date__ = "5 June 2024"

import re
import sys
import os
import pandas as pd
import warnings
from enum import Enum
from typing import Union
from excel_operator import col_num_to_excel_col_name
from warningMsgProvider import MdOpStatus, WarningMsgProvider

warning_msg_provider = WarningMsgProvider()
warnings.simplefilter(action="ignore", category=pd.errors.PerformanceWarning)


# Markdown で記述した行を識別するためのクラス
class MarkdownLine(Enum):
    FREE_AREA = 0
    SUMMARY_AREA = 1
    TEST_ENV_FRAME_AREA = 2
    TEST_ITEMS_AREA = 3


# Excel の行と列を識別するためのクラス
class ExcelRow(Enum):
    SUMMARY = 0
    TEST_ENV_FRAME = 1
    TEST_HEADER = 2
    TEST_ITEMS = 3


class ExcelCol(Enum):
    LV1 = 0
    LV6 = 1
    NUMBER = 2
    ENVIRONMENT = 3
    PRECONDITION = 4
    STEPS = 5
    EXPECTED = 6
    NOTES = 7
    TEST_INTENTION = 8


# 「手順」や「番号リスト」の番号を再割り当てするクラス
#   Markdown の番号リストを `1.` のみで記述している場合に連番に変換する
class ListNumConverter:
    def __init__(self, config_md: dict):
        self.conf_md = config_md
        self.number_list_patterns = []
        self.number_list_patterns_str = []
        self.list_num_counter = []

        # 変換の対象とする行の正規表現
        self.number_list_patterns.append(
            "(" + self.conf_md["mark_for_read"]["steps"].replace("[0-9]+", "1") + ")"
        )
        for v in self.conf_md["aux_mark"]["nested"]["number_list_lv"]:
            self.number_list_patterns.append("(" + v.replace("[0-9]+", "1") + ")")
        self.number_list_patterns_str = "|".join(self.number_list_patterns)

        # 番号カウンターの初期化
        self.reset()

    def conv(self, line: str, nest_lv: int) -> str:
        line = line.replace("1", str(self.list_num_counter[nest_lv]), 1)
        self.list_num_counter[nest_lv] += 1
        return line

    def reset(self):
        self.list_num_counter = [1] * len(self.number_list_patterns)

    def renumbering(self, current_nest_lv, previous_nest_lv):
        for i in range(len(self.list_num_counter)):
            if current_nest_lv < i and i <= previous_nest_lv:
                self.list_num_counter[i] = 1

def load_md(input_path: str):
    try:
        warning_msg_provider.setTargetFP(input_path)
        input_file = open(input_path, "r", encoding="utf-8")
        return input_file
    except FileNotFoundError:
        msg = warning_msg_provider.buildMsg(MdOpStatus.ERROR_CODE_1.value)
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)


def append_df(
    df: pd.DataFrame, current_item_dict: dict, item_counter: dict, config_md: dict
) -> pd.DataFrame:
    # 項目のナンバリングとカウンター更新
    k = current_item_dict["mark"]
    if k in item_counter:
        current_item_dict[k] = item_counter[k]

        isIncremented = False
        for (
            kk
        ) in (
            item_counter
        ):  # Memo: Python 3.6 より dictionary 型のキーの順序は保証されている
            if isIncremented:
                if kk != "number":  # number は通し番号なのでリセットしない
                    item_counter[kk] = 1
            else:
                if k == kk:
                    item_counter[kk] += 1
                    isIncremented = True

    # 行追加
    list = [[v for _, v in current_item_dict.items()]]
    df_append = pd.DataFrame(data=list, columns=df.columns.to_list())
    _df = pd.concat([df, df_append], ignore_index=True)

    # 初期化
    for k in current_item_dict:
        current_item_dict[k] = ""

    return _df


def check_if_append_df(current_item_dict: dict) -> Union[bool, str]:
    steps = current_item_dict["steps"]
    expected = current_item_dict["expected"]
    environment = current_item_dict["environment"]
    precondition = current_item_dict["precondition"]
    notes = current_item_dict["notes"]

    if steps and expected:
        return True
    elif (environment or precondition or notes) and (not steps or not expected):
        return "Error"
    elif (steps and not expected) or (not steps and expected):
        return "Error"
    else:
        return False


def convert_md_to_df(
    input_path: str, config_md: dict
) -> tuple[pd.DataFrame, str, str, list, list, list]:
    """
    Args:
        input_path:        入力ファイルパス
        config_md:         マークダウン部分に関する設定

    Returns:
        df:                データフレーム型テスト項目書
        sheet_name:        Excelのシート名
        product_categorie:      製品カテゴリの略称
        summary:           概要欄の入力文章
        test_env_frame     テスト環境枠
        warning:           Markdownの記述、その他に関する警告
    """

    cur_mark = ""
    prev_mark = ""
    cur_nest_lv = 0
    prev_nest_lv = 0
    title_detected = False
    prev_test_viewpoint_lv = 0
    prev_line = ""
    line_feed = {"flag": False, "indent": 0}
    md_line_section = MarkdownLine.FREE_AREA
    bullet_point_mark = "・"
    item_counter = {
        "lv1": 1,
        "lv2": 1,
        "lv3": 1,
        "lv4": 1,
        "lv5": 1,
        "lv6": 1,
        "number": 1,
    }
    lstNumConverter = ListNumConverter(config_md)

    # テスト項目表用の空データフレーム
    df = pd.DataFrame(columns=[k for k, _ in config_md["col_name"].items()])
    current_item_dict = {k: "" for k, _ in config_md["col_name"].items()}
    # シート名
    sheet_name = ""
    # 製品カテゴリの略称
    product_categorie = ""
    # 上記表の上に記載する概要文章用の空リスト
    summary = []
    # テスト環境枠用の空リスト
    test_env_frame = []
    # 警告メッセージ格納用（Excel に変換されないデータなどの警告）
    warning = []

    input_file = load_md(input_path)

    def resetLstNum():
        lstNumConverter.reset()

    def reset_line_feed_flags():
        line_feed["flag"] = False
        line_feed["indent"] = 0

    def get_sheet_name(md_file_path):
        warning_msg_provider.setTargetFP(md_file_path)
        s_name = os.path.splitext(os.path.basename(md_file_path))[0]

        sheet_name_err = False
        invalid_chars = [
            ":",
            "\\",
            "/",
            "?",
            "*",
            "[",
            "]",
            "：",
            "￥",
            "／",
            "？",
            "＊",
            "［",
            "］",
        ]
        if any(char in s_name for char in invalid_chars):
            sheet_name_err = True
        elif len(s_name) > 31:
            sheet_name_err = True
        elif s_name == "":
            sheet_name_err = True

        if sheet_name_err:
            msg = warning_msg_provider.buildMsg(MdOpStatus.ERROR_CODE_2.value)
            print(msg)
            input("何かキーを押してください...")
            sys.exit(1)
        else:
            return s_name

    sheet_name = get_sheet_name(input_path)

    for i, line in enumerate(input_file):

        # タイトル行
        if re.match(config_md["mark_for_read"]["title"], line):
            # タイトル行 - テスト観点行間は概要欄とする
            md_line_section = MarkdownLine.SUMMARY_AREA
            title_detected = True
            product_categorie = prev_line.strip()

        # テスト環境枠
        elif re.match(config_md["mark_for_read"]["test_env_frame"], line):
            if md_line_section == MarkdownLine.SUMMARY_AREA:
                md_line_section = MarkdownLine.TEST_ENV_FRAME_AREA
            elif md_line_section == MarkdownLine.TEST_ENV_FRAME_AREA:
                md_line_section = MarkdownLine.SUMMARY_AREA
        # メモ欄（Excel 変換の対象外）
        elif md_line_section == MarkdownLine.FREE_AREA:
            # ファイルの先頭 - タイトル行間はメモ欄とする
            prev_line = line
            continue

        # テスト観点行
        elif line.startswith("#"):
            # テスト環境枠の記載エリアが閉じられていない場合はエラーとする
            if md_line_section == MarkdownLine.TEST_ENV_FRAME_AREA or len(
                test_env_frame
            ) != len(set(test_env_frame)):
                msg = warning_msg_provider.buildMsg(MdOpStatus.ERROR_CODE_3.value)
                print(msg)
                input("何かキーを押してください...")
                sys.exit(1)

            # テスト項目エリア開始時
            elif md_line_section == MarkdownLine.SUMMARY_AREA:
                if len(test_env_frame) == 0:
                    test_env_frame.append("")
                # データフレームにテスト環境枠の列追加
                for i in range(len(test_env_frame)):
                    for name in [k for k, _ in config_md["col_name_res_area"].items()]:
                        tmp_name = name + "_" + str(i + 1)
                        df[tmp_name] = ""
                        current_item_dict[tmp_name] = ""

            reset_line_feed_flags()

            cur_mark = prev_mark = ""
            cur_nest_lv = prev_nest_lv = 0
            lstNumConverter.reset()
            md_line_section = MarkdownLine.TEST_ITEMS_AREA

            # Lv1 - Lv6
            for k, v in config_md["mark_for_read"].items():
                if re.match(v, line):
                    # このテスト観点の直前で生成したテスト項目行があれば追加
                    res = check_if_append_df(current_item_dict)
                    if res == "Error":
                        line_num = i + 1
                        msg = warning_msg_provider.buildMsg(
                            MdOpStatus.ERROR_CODE_9.value, str(line_num)
                        )
                        print(msg)
                        input("何かキーを押してください...")
                        sys.exit(1)
                    elif res:
                        df = append_df(df, current_item_dict, item_counter, config_md)

                    # テスト観点行追加（lv6 の空白見出しの場合は、テスト観点行を追加しない）
                    is_lv6_with_empty_content = re.match(
                        config_md["mark_for_read"]["lv6"] + "*$", line
                    )
                    if not is_lv6_with_empty_content:
                        current_item_dict["mark"] = k
                        current_item_dict["environment"] = (
                            re.sub(v, "", line).replace("\n", "").lstrip()
                        )
                        df = append_df(df, current_item_dict, item_counter, config_md)

                        # テスト観点のレベルが1つ飛ばして上がったとき警告する
                        cur_test_viewpoint_lv = v.count("#")
                        if cur_test_viewpoint_lv - prev_test_viewpoint_lv >= 2:
                            line_num = i + 1
                            warning.append(
                                warning_msg_provider.buildMsg(
                                    MdOpStatus.WARNING_CODE_5.value, str(line_num)
                                )
                            )
                        prev_test_viewpoint_lv = cur_test_viewpoint_lv

                    break

        # 概要欄 ※ 項目表内に表示する情報ではないため `df` とは別に `summary` にデータを格納していく
        elif md_line_section == MarkdownLine.SUMMARY_AREA:
            summary.append(line)

        # テスト環境枠
        elif md_line_section == MarkdownLine.TEST_ENV_FRAME_AREA:
            tmp_str = line.strip()
            if tmp_str != "":
                test_env_frame.append(line.replace("\n", "").strip())

        # テスト項目行
        else:
            cell_data = ""
            current_item_dict["mark"] = "number"

            # 前提・手順・確認・備考
            if re.match(config_md["mark_for_read"]["environment"], line):
                cur_mark = "environment"
                resetLstNum()
                cell_data = (
                    bullet_point_mark
                    + re.sub(config_md["mark_for_read"][cur_mark], "", line).lstrip()
                )
            elif re.match(config_md["mark_for_read"]["precondition"], line):
                cur_mark = "precondition"
                resetLstNum()
                cell_data = (
                    bullet_point_mark
                    + re.sub(config_md["mark_for_read"][cur_mark], "", line).lstrip()
                )
            elif re.match(config_md["mark_for_read"]["steps"], line):
                cur_mark = "steps"
                if cur_mark != prev_mark:
                    lstNumConverter.reset()
                cur_nest_lv = 0
                lstNumConverter.renumbering(cur_nest_lv, prev_nest_lv)
                prev_nest_lv = cur_nest_lv
                cell_data = lstNumConverter.conv(line, cur_nest_lv)
            elif re.match(
                config_md["mark_for_read"]["expected"], line
            ) and not re.match(config_md["mark_for_read"]["notes"], line):
                # 実施判定 の初期値設定
                for i in range(len(test_env_frame)):
                    tmp_name = "test_intention_" + str(i + 1)
                    if current_item_dict[tmp_name] == "":
                        current_item_dict[tmp_name] = config_md["test_intention"][
                            "inclusion_word"
                        ]

                cur_mark = "expected"
                resetLstNum()
                cell_data = (
                    bullet_point_mark
                    + re.sub(config_md["mark_for_read"][cur_mark], "", line).lstrip()
                )
            elif re.match(config_md["mark_for_read"]["notes"], line):
                cur_mark = "notes"
                resetLstNum()

                # 実施 or 省略の判定
                omission_str_all_test_env = "- [x] "
                omission_word = config_md["test_intention"]["omission_word"]
                specified_test_env_omission_idx = []

                if line.startswith(omission_str_all_test_env):
                    for idx, env_name in enumerate(test_env_frame):
                        omission_str_specified_test_env = (
                            omission_str_all_test_env + env_name
                        )
                        # テスト環境の指定がある場合は、その環境のみ省略
                        if line.startswith(omission_str_specified_test_env):
                            specified_test_env_omission_idx.append(idx + 1)
                    # テスト環境の指定がない場合はすべて省略
                    if not specified_test_env_omission_idx:
                        for idx, env_name in enumerate(test_env_frame):
                            current_item_dict["test_intention_" + str(idx + 1)] = (
                                omission_word
                            )
                    else:
                        # 以下のようなケースで省略を指定したテスト環境に、意図しないものまで含まれるケースがあるので、最後尾のみ取得する
                        # `- [x] 環境10` としたが、 環境1まで省略の対象になってしまった..
                        current_item_dict[
                            "test_intention_" + str(specified_test_env_omission_idx[-1])
                        ] = omission_word

                cell_data = (
                    bullet_point_mark
                    + re.sub(config_md["mark_for_read"][cur_mark], "", line).lstrip()
                )
            elif re.match(config_md["mark_for_read"]["caption"], line):
                pass
            elif re.match(config_md["mark_for_read"]["separator"], line):
                pass
            # 空白行
            elif re.match("^\n", line):
                if cur_mark:
                    cell_data = "\n"
            # 入れ子のリスト
            elif re.match("^(    ){1,}([0-9]+. |- |\+ |\* )", line):
                nested = config_md["aux_mark"]["nested"]
                for key in nested.keys():
                    for idx, val in enumerate(nested[key]):
                        cur_nest_lv = idx + 1
                        if re.match(val, line):

                            lstNumConverter.renumbering(cur_nest_lv, prev_nest_lv)

                            if key == "points_list_lv":
                                cell_data = bullet_point_mark + re.sub(
                                    "^(    ){1,}[-\+\*] ", "", line
                                )
                            elif key == "number_list_lv":
                                tmp_line = lstNumConverter.conv(line, cur_nest_lv)
                                cell_data = re.sub("^(    ){1,}", "", tmp_line)
                            line_feed["indent"] = config_md["aux_mark"][
                                "nested_list_indent_lv"
                            ][idx]
                            prev_nest_lv = cur_nest_lv
                            total_len = line_feed["indent"] + len(cell_data)
                            cell_data = cell_data.rjust(total_len)
            # 改行して挿入　※１つ前の行末に半角スペースが2つ以上あり 且つ、上記条件に一致しない行
            elif line_feed["flag"]:
                cell_data = line.lstrip()
                total_len = line_feed["indent"] + len(cell_data) + 2
                cell_data = cell_data.rjust(total_len)
            # 上記以外の無効データ（Excelに変換されないもの）について警告
            else:
                line_num = i + 1
                warning.append(
                    warning_msg_provider.buildMsg(
                        MdOpStatus.WARNING_CODE_2.value,
                        str(line_num),
                        line.replace("\n", ""),
                    )
                )

            if cur_mark:
                prev_mark = cur_mark

                # 行末の改行シンボル（半角スペース）チェック
                if cell_data.endswith("  \n"):
                    line_feed["flag"] = True
                    cell_data = cell_data.rstrip() + "\n"
                elif cell_data == "\n":
                    pass
                else:
                    reset_line_feed_flags()

                # 空情報のリストは無効にする
                if re.match("^" + bullet_point_mark + "( |\\n)*$", cell_data):
                    cell_data = ""

                # 1行分の情報を追加
                if cell_data != "":
                    current_item_dict[cur_mark] += cell_data

        prev_line = line

    # タイトル行がない場合はエラーとする
    if not title_detected:
        msg = warning_msg_provider.buildMsg(MdOpStatus.ERROR_CODE_8.value)
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)

    # ファイル終了時点の最後の項目を追加
    res = check_if_append_df(current_item_dict)
    if res == "Error":
        msg = warning_msg_provider.buildMsg(
            MdOpStatus.ERROR_CODE_9.value, str("最終")
        )
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)
    elif res:
        df = append_df(df, current_item_dict, item_counter, config_md)

    if check_if_append_df(current_item_dict):
        df = append_df(df, current_item_dict, item_counter, config_md)
    return df, sheet_name, product_categorie, summary, test_env_frame, warning


def convert_df_to_md(
    df: pd.DataFrame, config_md: dict, output_fn: str, sheet_pos_order: int, product_categorie: str
) -> list:
    """
    convert_df_to_md()により生成されたデータフレームを Markdown に変換します

    Args:
        df:                convert_excel_to_df()により生成されたデータフレーム
        config_md:         yamlで定義している設定
        output_fn:         出力先のファイル
        sheet_pos_order    シートの並び順
        product_categorie  製品カテゴリー


    Returns:
        None
    """

    # 警告メッセージ格納用（Markdown に変換されないデータなどの警告）
    warning = []
    warning_target_fp = os.path.splitext(output_fn)[0] + " シート"
    warning_msg_provider.setTargetFP(warning_target_fp)

    # DataFrame を List に変換
    sheet_list = []
    for index, row in df.iterrows():
        sheet_list.append(row.values.tolist())

    # ヘッダー情報（ヘッダー行の探索用）
    header_data = []
    for k, v in config_md["col_name"].items():
        if k != "mark":
            header_data.append(v)

    # 各データの位置情報（行番号 or 列番号）を取得していく
    row_idx = [-1] * len(ExcelRow.__members__)
    col_idx = [-1] * len(ExcelCol.__members__)
    row_idx[ExcelRow.SUMMARY.value] = 0

    col_idx_start = -1
    col_idx_test_env_frame = []
    for r_idx, row_data in enumerate(sheet_list):
        col_idx_start = [
            i
            for i, x in enumerate(row_data)
            if row_data[i : i + len(header_data)] == header_data
        ]
        if col_idx_start:
            row_idx[ExcelRow.TEST_HEADER.value] = r_idx
            row_idx[ExcelRow.TEST_ENV_FRAME.value] = r_idx - 1
            row_idx[ExcelRow.TEST_ITEMS.value] = r_idx + 1
            col_idx[ExcelCol.LV1.value] = col_idx_start[0]
            col_idx[ExcelCol.LV6.value] = col_idx[ExcelCol.LV1.value] + 5
            col_idx[ExcelCol.NUMBER.value] = col_idx[ExcelCol.LV6.value] + 1
            col_idx[ExcelCol.ENVIRONMENT.value] = col_idx[ExcelCol.NUMBER.value] + 1
            col_idx[ExcelCol.PRECONDITION.value] = (
                col_idx[ExcelCol.ENVIRONMENT.value] + 1
            )
            col_idx[ExcelCol.STEPS.value] = col_idx[ExcelCol.PRECONDITION.value] + 1
            col_idx[ExcelCol.EXPECTED.value] = col_idx[ExcelCol.STEPS.value] + 1
            col_idx[ExcelCol.NOTES.value] = col_idx[ExcelCol.EXPECTED.value] + 1
            col_idx[ExcelCol.TEST_INTENTION.value] = col_idx[ExcelCol.NOTES.value] + 1

            # テスト環境枠の開始位置取得
            for c_idx, data in enumerate(row_data):
                if data == config_md["col_name_res_area"]["test_intention"]:
                    col_idx_test_env_frame.append(c_idx)

            break

    if any(r == -1 for r in row_idx) or any(c == -1 for c in col_idx):
        msg = warning_msg_provider.buildMsg(MdOpStatus.ERROR_CODE_4.value)
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)

    # Markdown ファイルに書き込んでいく
    arr_md_str = []
    # タイトル 及び シート配置順
    # arr_md_str.append(os.path.splitext(output_fn)[0])
    arr_md_str.append(product_categorie)
    arr_md_str.append(config_md["mark_for_write"]["title"] * sheet_pos_order)

    row_section = ExcelRow.SUMMARY.value
    prev_is_test_row = None
    for r_idx, row_data in enumerate(sheet_list):
        # 行のカテゴリを識別
        if r_idx == row_idx[ExcelRow.TEST_ENV_FRAME.value]:
            row_section = ExcelRow.TEST_ENV_FRAME.value
        elif r_idx == row_idx[ExcelRow.TEST_HEADER.value]:
            row_section = ExcelRow.TEST_HEADER.value
        elif r_idx == row_idx[ExcelRow.TEST_ITEMS.value]:
            row_section = ExcelRow.TEST_ITEMS.value

        md_str_at_row, warning_at_row, is_test_row = getRowDataConvertedToMarkdown(
            r_idx,
            row_data,
            row_section,
            prev_is_test_row,
            col_idx,
            col_idx_test_env_frame,
            config_md,
        )
        prev_is_test_row = is_test_row

        arr_md_str.extend(md_str_at_row)

        if len(warning_at_row):
            for w in warning_at_row:
                warning.extend(w)

    if os.path.exists(output_fn):
        print("\n保存先のファイルが既に存在します " + "(" + output_fn + ")")
        while True:
            user_input = input("→ 上書きしますか? (y/n): ").lower()
            if user_input == 'y':
                print("")
                break
            elif user_input == 'n':
                print(output_fn + " の書き込みをスキップしました\n")
                return []
            else:
                print("→ 'y' または 'n' いずれかのキーを押してください")

    with open(output_fn, mode="w", encoding="utf-8") as f:
        for line in arr_md_str:
            f.write(line + "\n")

    return warning


def getRowDataConvertedToMarkdown(
    row_index: int,
    row_data: list,
    row_section: int,
    prev_is_test_row: bool,
    col_idx: list,
    col_idx_test_env_frame: list,
    config_md: dict,
) -> tuple[list, list, bool]:

    def get_omission_test_env_words(row_data, col_idx_test_env_frame) -> list:
        res = []
        ##### vvv
        # 逆変換するとき実施環境枠は、ひとまず最初の枠のみ変換する仕様とした
        # http://ghe.nanao.co.jp/SQG/Tools_ST/issues/54#issuecomment-292293
        #
        if row_data[col_idx_test_env_frame[0]] == "省略":
            res.append("- [ ] ")
        ##### ^^^
        return res

    md_str_at_row = []
    warning_at_row = []
    tmp_warning_msg = []
    is_test_row = False

    col_lv1_idx = col_idx[ExcelCol.LV1.value]
    col_lv6_idx = col_idx[ExcelCol.LV6.value]

    # 各カテゴリごとの処理
    if row_section == ExcelRow.SUMMARY.value:
        summary = ""
        for c_idx, cell_data in enumerate(row_data):
            if c_idx < col_lv1_idx:
                continue

            if cell_data != "":
                if summary == "":
                    summary = cell_data.rstrip("\n")
                else:
                    cell_num = col_num_to_excel_col_name(c_idx + 1) + str(row_index + 1)
                    if len(tmp_warning_msg) == 0:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_3.value,
                            str(row_index + 1),
                            str(cell_num),
                            str(cell_data),
                        )
                    else:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_4.value,
                            "",
                            str(cell_num),
                            str(cell_data),
                        )
                    tmp_warning_msg.append(m)

        if summary != "":
            md_str_at_row.append(summary)

    elif row_section == ExcelRow.TEST_ENV_FRAME.value:
        test_env_frame = []
        for c_idx, cell_data in enumerate(row_data):
            if c_idx < col_lv1_idx:
                continue

            if c_idx in col_idx_test_env_frame:
                test_env_frame.append(cell_data)
            else:
                if cell_data:
                    cell_num = col_num_to_excel_col_name(c_idx + 1) + str(row_index + 1)
                    if len(tmp_warning_msg) == 0:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_3.value,
                            str(row_index + 1),
                            str(cell_num),
                            str(cell_data),
                        )
                    else:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_4.value,
                            "",
                            str(cell_num),
                            str(cell_data),
                        )
                    tmp_warning_msg.append(m)

        if len(test_env_frame):
            md_str_at_row.append("")
            md_str_at_row.append(config_md["mark_for_write"]["test_env_frame"])

            ##### vvv
            # 逆変換するとき実施環境枠は、ひとまず最初の枠のみ変換する仕様とした
            # http://ghe.nanao.co.jp/SQG/Tools_ST/issues/54#issuecomment-292293
            #
            # md_str_at_row.extend(test_env_frame)
            md_str_at_row.extend(test_env_frame[:1])
            ##### ^^^

            md_str_at_row.append(config_md["mark_for_write"]["test_env_frame"])
            md_str_at_row.append("")

    elif row_section == ExcelRow.TEST_HEADER.value:
        pass

    elif row_section == ExcelRow.TEST_ITEMS.value:
        mandatory_columns = []
        optional_columns = []
        test_viewpoint_lv_idx = -1

        is_test_row = True if row_data[col_idx[ExcelCol.NUMBER.value]] else False

        # 列に対する入力データの条件フラグを設定 - 必須（mandatory_columns）or 任意（optional_columns) or 無効
        # テスト観点行
        if not is_test_row:
            for idx, x in enumerate(row_data[0 : col_lv6_idx + 1], 0):
                if idx < col_lv1_idx:
                    continue
                if x:
                    mandatory_columns.append(idx)
                    test_viewpoint_lv_idx = idx - col_lv1_idx
                    break
            if not mandatory_columns:
                msg = warning_msg_provider.buildMsg(
                    MdOpStatus.ERROR_CODE_5.value, str(row_index + 1)
                )
                print(msg)
                input("何かキーを押してください...")
                sys.exit(1)

            optional_columns.append(col_idx[ExcelCol.ENVIRONMENT.value])

        # テスト項目行
        else:
            mandatory_columns.append(col_idx[ExcelCol.NUMBER.value])
            optional_columns.append(col_idx[ExcelCol.ENVIRONMENT.value])
            optional_columns.append(col_idx[ExcelCol.PRECONDITION.value])
            mandatory_columns.append(col_idx[ExcelCol.STEPS.value])
            mandatory_columns.append(col_idx[ExcelCol.EXPECTED.value])
            optional_columns.append(col_idx[ExcelCol.NOTES.value])
            ##### vvv
            # 逆変換するとき実施環境枠は、ひとまず最初の枠のみ変換する仕様とした
            # http://ghe.nanao.co.jp/SQG/Tools_ST/issues/54#issuecomment-292293
            #
            mandatory_columns.append(col_idx[ExcelCol.TEST_INTENTION.value])
            ##### ^^^

        if is_test_row and prev_is_test_row:
            md_str_at_row.append(config_md["mark_for_write"]["test_rows"]["lv6"])

        # 各列の入力データを MarkDown に変換しながら配列に格納していく
        for c_idx, cell_data in enumerate(
            row_data[: col_idx[ExcelCol.NOTES.value] + 1]
        ):
            if c_idx < col_lv1_idx:
                continue

            if c_idx in mandatory_columns:
                if cell_data != "":
                    md_str_at_row.extend(
                        convCellToMDStrLst(
                            row_index,
                            cell_data,
                            c_idx,
                            col_idx,
                            config_md,
                            test_viewpoint_lv_idx,
                        )
                    )
                else:
                    cell_num = col_num_to_excel_col_name(c_idx + 1) + str(row_index + 1)
                    msg = warning_msg_provider.buildMsg(
                        MdOpStatus.ERROR_CODE_6.value,
                        str(row_index + 1),
                        str(cell_num),
                        str(cell_data),
                    )
                    print(msg)
                    input("何かキーを押してください...")
                    sys.exit(1)

            elif c_idx in optional_columns:
                if c_idx == col_idx[ExcelCol.NOTES.value]:
                    # 「実施判定」の `省略` をここで反映する
                    omission_test_env_words = get_omission_test_env_words(
                        row_data, col_idx_test_env_frame
                    )
                    md_str_at_row.extend(
                        convCellToMDStrLst(
                            row_index,
                            cell_data,
                            c_idx,
                            col_idx,
                            config_md,
                            test_viewpoint_lv_idx,
                            omission_test_env_words,
                        )
                    )
                else:
                    md_str_at_row.extend(
                        convCellToMDStrLst(
                            row_index,
                            cell_data,
                            c_idx,
                            col_idx,
                            config_md,
                            test_viewpoint_lv_idx,
                        )
                    )

            else:
                if cell_data != "":
                    cell_num = col_num_to_excel_col_name(c_idx + 1) + str(row_index + 1)
                    if len(tmp_warning_msg) == 0:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_3.value,
                            str(row_index + 1),
                            str(cell_num),
                            str(cell_data),
                        )
                    else:
                        m = warning_msg_provider.buildMsg(
                            MdOpStatus.WARNING_CODE_4.value,
                            "",
                            str(cell_num),
                            str(cell_data),
                        )
                    tmp_warning_msg.append(m)

    warning_at_row.append(tmp_warning_msg)

    return md_str_at_row, warning_at_row, is_test_row


def convCellToMDStrLst(
    row_index: int,
    cell_data: str,
    c_idx: int,
    col_idx: list,
    config_md: dict,
    test_viewpoint_lv_idx=-1,
    omission_test_env_words=None,
) -> list:

    def markReplacer(cell_data, mark, nested_lst_mark, nested_num_mark) -> list:

        result = []

        lines = cell_data.split("\n")

        if not list(filter(lambda item: item.strip() != "", lines)):
            # セルが空白のときは記号のみ追加
            result.append(mark)
        else:
            newline_pos_idx = -1
            for idx, line in enumerate(lines):
                is_empty_line = False

                line = line.rstrip()

                if re.match("^・", line):
                    result.append(re.sub("^・", mark, line))
                elif re.match("  ・", line):
                    result.append(re.sub("  ・", "    " + nested_lst_mark, line))
                elif re.match("    ・", line):
                    result.append(re.sub("    ・", "        " + nested_lst_mark, line))
                elif re.match("      ・", line):
                    result.append(
                        re.sub("      ・", "            " + nested_lst_mark, line)
                    )
                elif re.match("^[0-9]+", line):
                    result.append(re.sub("^[0-9]+. *", mark, line))
                elif re.match("  [0-9]+", line):
                    result.append(re.sub("  [0-9]+. *", "    " + nested_num_mark, line))
                elif re.match("    [0-9]+", line):
                    result.append(
                        re.sub("    [0-9]+. *", "        " + nested_num_mark, line)
                    )
                elif re.match("      [0-9]+", line):
                    result.append(
                        re.sub("      [0-9]+. *", "            " + nested_num_mark, line)
                    )
                # 改行 + 空白文字のみの行（空白行）
                elif re.match("^\\s*?$(\\r\\n|\\r|\\n)?", line):
                    is_empty_line = True
                    result.append("")
                else:
                    if newline_pos_idx >= 0:
                        # 改行元の末尾に半角スペース2個追加
                        result[newline_pos_idx] = result[newline_pos_idx] + "  "
                    else:
                        # 改行元が不定な改行記述があった場合は処理中止
                        cell_num = col_num_to_excel_col_name(c_idx + 1) + str(
                            row_index + 1
                        )
                        msg = warning_msg_provider.buildMsg(
                            MdOpStatus.ERROR_CODE_7.value,
                            str(row_index + 1),
                            str(cell_num),
                            str(cell_data),
                        )
                        print(msg)
                        input("何かキーを押してください...")
                        sys.exit(1)

                    result.append(line)

                if not is_empty_line:
                    newline_pos_idx = idx

        # 省略項目の処理
        if omission_test_env_words:
            ##### vvv
            # 逆変換するとき実施環境枠は、ひとまず最初の枠のみ変換する仕様とした
            # http://ghe.nanao.co.jp/SQG/Tools_ST/issues/54#issuecomment-292293
            #
            pattern = re.compile(r"^- \[\s\] ")
            result = [re.sub(pattern, "- [x] ", s) for s in result]
            ##### ^^^

        if result and result[-1].strip() == "":
            del result[-1]

        return result

    md_str_at_col = []
    repraced_cell_data = []
    mark = config_md["mark_for_write"]["test_rows"]

    # テスト観点(lv1 - lv6)
    if test_viewpoint_lv_idx >= 0:
        # テスト観点行の記号とテスト観点の記述を結合する
        if c_idx == col_idx[ExcelCol.ENVIRONMENT.value]:
            header_symbol = mark["lv" + str(test_viewpoint_lv_idx + 1)]
            md_str_at_col.append(header_symbol + cell_data)
            test_viewpoint_lv_idx = -1

    # No.
    elif c_idx == col_idx[ExcelCol.NUMBER.value]:
        pass

    # 環境
    elif c_idx == col_idx[ExcelCol.ENVIRONMENT.value]:
        md_str_at_col.append(mark["environment_caption"])
        repraced_cell_data = markReplacer(
            cell_data, mark["environment"], mark["environment"], mark["steps"]
        )

    # 準備
    elif c_idx == col_idx[ExcelCol.PRECONDITION.value]:
        md_str_at_col.append(mark["precondition_caption"])
        repraced_cell_data = markReplacer(
            cell_data, mark["precondition"], mark["precondition"], mark["steps"]
        )

    # 手順
    elif c_idx == col_idx[ExcelCol.STEPS.value]:
        md_str_at_col.append(mark["steps_caption"])
        repraced_cell_data = markReplacer(
            cell_data, mark["steps"], mark["expected"], mark["steps"]
        )

    # 確認
    elif c_idx == col_idx[ExcelCol.EXPECTED.value]:
        md_str_at_col.append(mark["expected_caption"])
        repraced_cell_data = markReplacer(
            cell_data, mark["expected"], mark["expected"], mark["steps"]
        )

    # 備考
    elif c_idx == col_idx[ExcelCol.NOTES.value]:
        md_str_at_col.append(mark["notes_caption"])
        repraced_cell_data = markReplacer(
            cell_data, mark["notes"], mark["expected"], mark["steps"]
        )
        repraced_cell_data.append("---")
        repraced_cell_data.append("")

    # 実施判定
    elif c_idx == col_idx[ExcelCol.TEST_INTENTION.value]:
        pass

    md_str_at_col.extend(repraced_cell_data)

    return md_str_at_col
