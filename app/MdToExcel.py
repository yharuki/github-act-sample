# coding: utf-8

"""
Markdownで書かれたテスト項目書をエクセルファイルに変換します。

Usage:
    MdToExcel.py [-f] <file>... [-m]

Options:
    -f, --file             入力ファイルパス

Requirements:
    - pandas
    - openpyxl 3.0.0 or higher
    - PyYAML 5.0.0 or higher
    - docopt

Notes:
    - 変換するMarkdownは以下の形式で記述してください

    # 観点1
    ## 観点2
    ### 観点3
    #### 観点4
    ##### 観点5
    ###### 観点6
    > 環境
    + 環境
    > 準備
    * 準備
    > 手順
    1. 手順
    1. 手順
    > 確認
    - 確認
    > 備考
    - [ ] 備考

"""

__author__ = "Yuji Haruki (modifier) / Kohei, Watanabe <kohei.watanabe3@brother.co.jp>"
__version__ = "2.1.0"
__date__ = "5 June 2024"

import os
import sys

try:
    import yaml
    import openpyxl
    from docopt import docopt
except ModuleNotFoundError as e:
    print("This program requires pandas/docopt/openpyxl>=3.0.0.")
    input()
    sys.exit(1)
assert (
    yaml.__version__ >= "5.0.0"
), "This program requires PyYAML>=5.0.0.\n$ pip install pyyaml==5.3.1"
assert (
    openpyxl.__version__ >= "3.0.0"
), "This program requires openpyxl>=3.0.0.\b$ pip install openpyxl==3.0.5"

from markdown_operator import convert_md_to_df, convert_df_to_md
from excel_operator import convert_df_to_excel, convert_excel_to_df
from warningMsgProvider import MainAppStatus, WarningMsgProvider
warning_msg_provider = WarningMsgProvider()


def resourcePath(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(filename)


def load_config() -> dict:
    try:
        with open(
            resourcePath("resources/config.yaml"), "r", encoding="utf-8_sig"
        ) as f:
            config = yaml.load(f, Loader=yaml.FullLoader)
    except FileNotFoundError:
        msg = warning_msg_provider.buildMsg(MainAppStatus.ERROR_CODE_1.value)
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)

    return config


def sort_by_specified_order(files) -> []:
    order_index = []
    for file in files:
        with open(file, "r", encoding="utf-8") as f:
            for line in f:
                count = 0
                for char in line:
                    if char == "=":
                        count += 1
                    else:
                        break
                if count > 0:
                    order_index.append(count)
                    break

    if len(files) != len(order_index):
        return files
    else:
        # (order_index, files) のタプルのリストを作成する
        lst = list(zip(order_index, files))

        # インデックスでソートし、同じインデックスの場合はファイルパスでソートする
        sorted_lst = sorted(lst, key=lambda x: (x[0], x[1]))

        # ソートされたリストからファイルパスのみを取り出す
        sorted_files = [x[1] for x in sorted_lst]
        return sorted_files

def isValidName(fn):
    for char in ["<", ">", ":", '"', "/", "\\", "|", "?", "*"]:
        if char in fn:
            print("")
            print(f"ファイル名に禁止文字 '{char}' が含まれています!")
            return False
    return True

def main():
    args = docopt(__doc__)
    files = args["<file>"]
    excel_book_save_names = []
    convert_type = -1
    # -1:単一ファイル変換 → 単一ファイル保存
    #  0:複数ファイル変換 → 単一ファイル保存（複数シート展開）
    #  1:複数ファイル変換 → 複数ファイル保存
    #  2:逆変換（Excel to Markdown）

    config = load_config()
    print("")
    print("MdToExcel ver." + __version__ + " 起動")
    print("")

    # 引数に指定されたファイルの種類からオペレーション選択（変換 or 逆変換）

    md_file_cnt = excel_file_cnt = other_file_cnt = 0
    for file in files:
        _, ext = os.path.splitext(file)
        if ext == ".md":
            md_file_cnt += 1
        elif ext == ".xlsm" or ext == ".xlsx":
            excel_file_cnt += 1
        else:
            other_file_cnt += 1

    if other_file_cnt or (md_file_cnt and excel_file_cnt):
        msg = warning_msg_provider.buildMsg(MainAppStatus.ERROR_CODE_2.value)
        print(msg)
        input("何かキーを押してください...")
        sys.exit(1)

    # Markdown -> Excel 変換処理
    if md_file_cnt:
        if len(files) > 1:
            print("複数のファイルが指定されました")
            print("")
            while True:
                print("次のいずれかの処理（0 or 1）を選択してください")
                print("")
                print("  0 → 1つの Excelブック に 複数のシート で展開する")
                print("  1 → 複数の Excelブック に 1シート ずつ展開する")
                print("")
                convert_type = input("数字を入力 : ")
                if convert_type == "0" or convert_type == "1":
                    break
                else:
                    continue

            files = sort_by_specified_order(files)

            if convert_type == "0":
                while True:
                    print("")
                    save_name = input(
                        "保存するExcelブックのファイル名（拡張子を除く）を指定してください : "
                    )
                    if save_name == "":
                        continue
                    else:
                        if not isValidName(save_name):
                            continue 
                        else:
                            break 
                excel_book_save_names.append(save_name)

            elif convert_type == "1":
                for file in files:
                    excel_book_save_names.append(
                        os.path.splitext(os.path.basename(file))[0]
                    )

        elif len(files) == 1:
            excel_book_save_names.append(
                os.path.splitext(os.path.basename(files[0]))[0]
            )

        dfs = []
        sheet_names = []
        product_categories = []
        summaries = []
        test_env_frames = []
        warnings = []
        print("")
        for file in files:
            print("Markdownファイル読み込み中 : " + file)
            df, sheet_name, product_categorie, summary, test_env_frame, warning = convert_md_to_df(
                file, config_md=config["md"]
            )
            dfs.append(df)
            sheet_names.append(sheet_name)
            product_categories.append(product_categorie)
            summaries.append(summary)
            test_env_frames.append(test_env_frame)
            warnings.extend(warning)

        for i in range(len(excel_book_save_names)):
            # prefer Docker `/app/tmp` if it exists; otherwise use a local `./tmp` directory
            docker_tmp = '/app/tmp'
            local_tmp = os.path.join(os.getcwd(), 'tmp')
            target_tmp = docker_tmp if os.path.isdir(docker_tmp) else local_tmp
            os.makedirs(target_tmp, exist_ok=True)
            output_fn = os.path.join(target_tmp, excel_book_save_names[i] + ".xlsm")
            print("Excelファイル書き込み中 : " + output_fn)

            tmp_dfs, tmp_sheet_names, tmp_product_categories, tmp_summaries, tmp_test_env_frames = (
                [],
                [],
                [],
                [],
                [],
            )
            if len(excel_book_save_names) == 1:
                tmp_dfs = dfs
                tmp_sheet_names = sheet_names
                tmp_product_categories = product_categories
                tmp_summaries = summaries
                tmp_test_env_frames = test_env_frames
            else:
                tmp_dfs.append(dfs[i])
                tmp_sheet_names.append(sheet_names[i])
                tmp_product_categories.append(product_categories[i])
                tmp_summaries.append(summaries[i])
                tmp_test_env_frames.append(test_env_frames[i])

            convert_df_to_excel(
                tmp_dfs,
                tmp_sheet_names,
                tmp_product_categories,
                tmp_summaries,
                tmp_test_env_frames,
                config_excel=config["excel"],
                input_path=resourcePath(
                    "resources/" + config["excel"]["template_file_name"]
                ),
                output_fn=output_fn,
                merge_cells=False,  # この機能は不要なため非サポートとしておく（将来的に削除したい）
            )

        if len(warnings):
            print("")
            print("【 警告 】")
            # print("Excel に変換されなかった行があります")
            # print("記述に間違いがないか確認してください")
            for msg in warnings:
                print(msg)
            print("")
            input("何かキーを押してください...")

        print("完了")

    # Excel -> Markdown 変換処理
    elif excel_file_cnt:
        dfs_lst = []
        product_categories_lst = []
        warnings = []
        print("")
        for file in files:
            print("Excelファイル読み込み中 : " + file)
            tmp_dfs, product_categorie = convert_excel_to_df(file)
            dfs_lst.append(tmp_dfs)
            product_categories_lst.append(product_categorie)

        for idx, dfs in enumerate(dfs_lst):
            sheet_pos_order = 0
            for sheet_name, df in dfs.items():
                file_name = sheet_name + ".md"
                print("Markdownファイル書き込み中 : " + file_name)
                sheet_pos_order += 1
                warning = convert_df_to_md(df, config["md"], file_name, sheet_pos_order, product_categories_lst[idx])
                if warning:
                    warnings.extend(warning)

        if len(warnings):
            print("")
            print("【 警告 】")
            print("Markdown に変換されなかった情報があります")
            print("記述に問題がないか確認してください")
            for msg in warnings:
                print(msg)
            print("")
            input("何かキーを押してください...")

        print("完了")


if __name__ == "__main__":
    main()
