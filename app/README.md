# MarkdownToExcel

![イメージ図](./img/img0.dio.svg)

<br>

- [MarkdownToExcel](#markdowntoexcel)
  - [概要](#概要)
  - [使い方](#使い方)
    - [作業環境準備](#作業環境準備)
    - [テスト項目書作成](#テスト項目書作成)
    - [Excelに変換](#excelに変換)
    - [いろいろな書き方](#いろいろな書き方)
      - [改行する](#改行する)
      - [入れ子のリストにする](#入れ子のリストにする)
      - [入れ子の中に入れ子をつくる](#入れ子の中に入れ子をつくる)
      - [空白行を入れる](#空白行を入れる)
      - [入れ子のリスト と 改行 と 空白行 の組み合わせ](#入れ子のリスト-と-改行-と-空白行-の組み合わせ)
      - [項目の「実施判定」を `省略` にする](#項目の実施判定を-省略-にする)
      - [テスト観点の一覧（目次）を作成する](#テスト観点の一覧目次を作成する)
    - [Excel から Markdown に逆変換](#excel-から-markdown-に逆変換)
      - [逆変換の仕様](#逆変換の仕様)
        - [基本仕様](#基本仕様)
        - [詳細仕様](#詳細仕様)
  - [その他](#その他)
    - [利点比較 (Excel と Markdown)](#利点比較-excel-と-markdown)
  - [開発者向け](#開発者向け)
    - [開発環境](#開発環境)
    - [構成](#構成)
    - [実行](#実行)
    - [実行ファイル(`exe`)のビルド](#実行ファイルexeのビルド)
    - [デバッグ](#デバッグ)

<br>

## 概要
[`Markdown` 形式](https://qiita.com/tbpgr/items/989c6badefff69377da7)で書かれたテスト項目書を `Excel` 形式に変換するツールです  
逆変換も可能です  
  テキスト形式ファイルでテスト設計ができるので、`GitHub` でのレビューがやりやすくなります(詳しくは[こちら](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/MarkdownToExcel/README.md#%E5%88%A9%E7%82%B9%E6%AF%94%E8%BC%83-excel-%E3%81%A8-markdown))



## 使い方


### 作業環境準備
1.以下から最新版のツール `MdToExcel.exe` を任意の場所に保存　　

`\\eizo-fsv\eec\EEC03_share\EEC038_share\30_検証関連\03_検証用ツール\2_検証用\MdToExcel`  

<br>

### Excelに変換
1. 後述の要領で作成した `Markdown` 形式のテスト項目書を `MdToExcel.exe` にドラッグアンドドロップすると、 `Excel` 形式に変換されます

    ![image](http://ghe.nanao.co.jp/storage/user/110/files/33f2bd8f-0afd-4dd6-b887-801e1b3ef647)

    :memo: 複数の Markdown ファイルを同時に変換することもできます

<br>

#### テスト項目書作成

1. 任意のリポジトリにmdファイルを作成する
    ```   
    例）フロントパネルのテスト.md
    ```

1. 上記ファイルを `Visual Studio Code` で開く
    - フォルダを右クリックし `Codeで開く` で `Visual Studio Code` を起動する
    - 起動したら作成したファイルを選択する


1. 以下の「テストの書き方」を参考に `Markdown` 形式でテスト項目書を作成していく
    - 補足
        - :memo: Markdown の [プレビュー表示](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/know-how/VSCode/README_vscode_install.md#markdown-%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%E3%83%97%E3%83%AC%E3%83%93%E3%83%A5%E3%83%BC%E8%A1%A8%E7%A4%BA)をしながら作業することをおすすめします
        - :memo: 下記のテンプレをコピペして利用できます

          <details>
          <summary> テスト項目書のテンプレート </summary>
      
          ```
          BP
          ===
          
          ■ 概要欄
          テストの概要、観点、注意事項、準備条件など
          
          # 観点1
          ## 観点2
          ### 観点3
          #### 観点4
          ##### 観点5
          ###### 観点6
          > 環境
          + テスト環境
          > 準備
          * 準備手順
          * 準備手順
          > 手順
          1. テスト手順
          1. テスト手順
          1. テスト手順
          > 確認
          - 確認手順
          - 確認手順
          > 備考
          - [ ] 備考
          ---
          
          ###### 観点6
          > 環境
          + テスト環境
          > 準備
          * 準備手順
          * 準備手順
          > 手順
          1. テスト手順
          1. テスト手順
          1. テスト手順
          > 確認
          - 確認手順
          - 確認手順
          > 備考
          - [ ] 備考
          ---
          
          #### 観点4
          > 準備
          * 準備手順
          > 手順
          1. テスト手順
              - 入れ子のリスト
              - 入れ子のリスト
          1. テスト手順
              1. 入れ子の番号リスト
              1. 入れ子の番号リスト
              1. 入れ子の番号リスト
          1. テスト手順
          > 確認
          - 確認手順
          - 確認手順
              1. 入れ子の番号リスト
                  1. 入れ子の番号リスト
                      - 入れ子のリスト
                      - 入れ子のリスト
                  1. 入れ子の番号リスト
                      - 入れ子のリスト
                      - 入れ子のリスト
              1. 入れ子の番号リスト
              1. 入れ子の番号リスト
          - 確認手順
          ---
          
          ###### 
          > 準備
          * 準備手順
          > 手順
          1. テスト手順
          1. テスト手順
          > 確認
          - 確認手順
          - 確認手順
          > 備考
          - [x] 〇〇の理由で省略
          ---
          
          ###### 
          > 準備
          * 準備手順
          > 手順
          1. テスト手順
          1. テスト手順
          > 確認
          - 確認手順
          - 確認手順
          > 備考
          - [ ] 備考
          ```
      
          </details>

    - テストの書き方  

      （💡 見えにくい場合はブラウザの表示を拡大してください）

      ![テストの書き方](./img/img2.dio.svg)

<br>



### いろいろな書き方

<br>

#### 改行する

- １つの手順の中で、改行して続けて何かを記載したいとき、  
行末に半角スペース2個記述すると改行できます

  <details>
  <summary> サンプル </summary>

  ![image](http://ghe.nanao.co.jp/storage/user/110/files/2360763b-df22-402a-a64a-d05392d68eb3)

  </details>

<br>

#### 入れ子のリストにする
- 各種手順の記載において先頭に半角スペース4個いれると、入れ子で記述できます

  <details>
  <summary> サンプル </summary>

  ![image](http://ghe.nanao.co.jp/storage/user/110/files/ac0706f0-5d3b-4502-8a72-18291991b3fe)

  - 半角スペース4つの後ろに記述できる記号は以下のとおりです
    |列|使用できる記号|備考|
    |---|---|---|
    |環境|`+` `-` `1.`|`1.` は番号付きリストになります|
    |準備|`*` `-` `1.`|同上|
    |手順|`-` `1.`|同上|
    |確認|`-` `1.`|同上|
    |備考|`-` `1.`|同上|
  
  - 入れ子の深さは３段階まであり、それぞれの深度に合わせて半角スペースの数は4個 or 8個 or 12個 となります
   
      ![image](http://ghe.nanao.co.jp/storage/user/110/files/cbef4ae9-c45e-4e15-89f0-3737976ab062)

  </details>

<br>

#### 入れ子の中に入れ子をつくる
- 入れ子の中を、さらに入れ子で記述できます
- 番号付きリストと普通のリストを織り交ぜることもできます

  <details>
  <summary> サンプル </summary>

    ![image](http://ghe.nanao.co.jp/storage/user/110/files/75cdf917-2bf1-4da4-88b7-1776359081a1)

  </details>

<br>

#### 空白行を入れる
- 以下のように空白行をいれることができます

  <details>
  <summary> サンプル </summary>

  ![image](http://ghe.nanao.co.jp/storage/user/110/files/7cb9d2d1-1f4f-4eb6-8776-d8512bc6b303)

  </details>

<br>

#### 入れ子のリスト と 改行 と 空白行 の組み合わせ
- 上記を組み合わせて記述した例です

  ![image](http://ghe.nanao.co.jp/storage/user/110/files/32cb6656-5315-49bc-aea9-fd00c5ad10ff)

<br>

#### 項目の「実施判定」を `省略` にする
- 備考欄の見出しを以下のようにすることで、省略項目にできる  
（省略理由などの文章記述は、あっても無くてもどちらでもOKです）

    ![image](http://ghe.nanao.co.jp/storage/user/110/files/71c2c341-b3d5-4bca-a465-517088780ee3)

    :warning: `Excel` の `条件付き書式設定` までは適用されないので、`Excel` 変換後に `Excel` 側で該当のマクロを実行する必要があります

<br>

#### テスト観点の一覧（目次）を作成する
- `Markdown` ファイルの先頭に目次を作成しておくと、いろいろと便利です
    - やり方は[こちら](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/know-how/VSCode/README_vscode_install.md#markdown-%E3%81%AE%E7%9B%AE%E6%AC%A1%E3%82%92%E8%87%AA%E5%8B%95%E7%94%9F%E6%88%90) を参照
    - リンクを `Ctrl`を押しながらクリックすると飛べます(下の図の左側の黒色文字)

    ![image](http://ghe.nanao.co.jp/storage/user/110/files/e0a40a0d-f930-4578-8e85-24bdf98e2b54)

### Excel から Markdown に逆変換
- Excel形式に変換したテスト項目書を再び `MdToExcel.exe` にドラッグアンドドロップすると、 `Markdown` 形式に逆変換されます

    ![image](http://ghe.nanao.co.jp/storage/user/110/files/7db7aec3-d6cb-49d8-a55d-9f29b2d60251)

    :memo: 複数の Excel ファイルを同時に逆変換することもできます
      
#### 逆変換の仕様

##### 基本仕様
- 逆変換をすると Excel テスト項目書のシート名が、Markdown テスト項目書のファイル名になります
- Excel テスト項目に記載の「・」で始まる記述が Markdown のリスト記述 「+」「*」「-」 「- [ ] 」に変換されます
- Excel テスト項目に記載の「1. 」や「2. 」の番号で始まる記述が Markdown の番号付きリストに変換されます
- （⚠️v2.0.2 以降）Excel 「表紙」シートのセル `A1` に記載の製品カテゴリー略称が、Markdown の `=` 記号の上に記述されます 
    ![img](./img/img7.dio.svg)

##### 詳細仕様

- テスト手順に入れ子の記述を含む場合
  <details>
  
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/b958ca60-97ce-47fd-9dcb-29b83dfcfffa)
  
  </details>

<br>

- テスト手順に改行を含む場合
  <details>
    
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/676f8331-87d8-4d2d-bf94-7b5b3ce144a0)
  
  </details>

<br>

- テスト手順に空白行を含む場合
  <details>
     
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/08c8c049-858a-4081-b52e-00e9711acfcd)
  
  </details>

<br>

- 複数シート・複数ファイルの場合

  <details>
  
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/f57e419b-1b72-4df8-a496-283e36bb5425)
  
  </details>

<br>

- Excel > Markdown 変換時の「警告」と「エラー」について  
Excel のテスト項目書において、不要なセルなどに記述があった場合、変換時に警告やエラーを表示します


  <details>
  <summary>セルに記述があると警告が表示されるエリア</summary>

  ![image](http://ghe.nanao.co.jp/storage/user/110/files/04522d12-778a-4d15-867c-f617122a5b83)

  </details>

  <br>

  <details>
  <summary>セルに記述があるとエラーになるエリア</summary>
  
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/5d54c88b-6e5a-44f6-8bbc-b82605ef6294)
  
  </details>

  <br>

  <details>
  <summary>入力必須のエリア（記述がない場合にエラー表示）</summary>
  
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/0d17ad9c-8b8a-4371-b7c0-6c76cdaf98f1)
  
  </details>

  <br>

  <details>
  <summary>入力任意のエリア（記述があった場合に警告表示）</summary>
  
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/06570c08-cef4-4290-9ac2-ab08c65c0e95)
  
  </details>

<br>

## その他
### 利点比較 (Excel と Markdown)
GitHub Enterprise 管理での下記テスト運用における利点と欠点
- Excel 運用
- Markdown + Visual Studio Code 運用 ※実施は Excel

<br>

**結論**

- 一般的なテスト項目書であれば、**慣れれば** Markdown 方式のほうが（レビューイ　・レビューアの作業全体でみて）効率がよくなり、成果物の品質も向上する  
  さらにその後の熟練度に比例して効率は向上する
  - ただし、以下についてある程度のスキルアップは必要
    - Markdown 記法の習得
    - Visual Studio Code （テキストエディター）のキーボード操作や検索などの機能活用
    - GitHub Enterprise のプルリク機能の利用
- Markdown はGitHubでの **差分管理** とテキストエディターでの **一括検索** (他のテスト項目書も含めた文書全体の検索) ができることの恩恵がとても大きい

| 項目          | Excel | Markdown + Visual Studio Code | 備考                                                                                                                                                                                                                                                                                                                                                          |
| ----------- | :---: | :---------------------------: | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 使いやすさ　・操作性   | ◯     | △ → (◯)                       | Excel<br>　・使い慣れている<br>　・直感的操作が可能<br>Markdown<br>　・テキストエディタに慣れていない<br>　・熟練すると◯                                                                                                                                                                                                                                                                                   |
| 学習コスト       | ◎    | △                             | Excel<br>　・既に熟練している<br>　・さらにキーボード操作の割合を増やして作業スピードを速くしたい場合は、ある程度の期間は練習が必要<br>Markdown<br>　・Markdown記法とVisual Studio Code の2つを学ぶ必要がある（難しくはないが練習して慣れるまでにある程度の期間が必要）                                                                                                                                                                                                                                                                                             |
| 機能拡張性       | △     | ◎                             | Excel<br>　・マクロで機能拡張できるがプログラミング技術が必要<br>Markdown<br>　・Markdownに関連するVsCodeの便利な拡張機能を誰でも簡単にインストールしてすぐに使える                                                                                                                                                                                                                                                            |
| テスト設計作業     | ◯     | ◯ → (◎)                             | Excel<br>　・特に問題ないが、強いて言えばセルの編集モード切り替えが煩わしい<br>　・テスト項目書内を検索できるがファイル単体でしかできない<br>Markdown<br>　・テスト手順を記述する作業において[テンプレート](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/know-how/VSCode/README_vscode_expansion.md#テスト項目を自動入力できるように設定する)を利用すれば、特に Excel で作成するのと工数的には変わらないが、慣れは必要<br>　・Visual Studio Code 上で他のテスト項目書も含めて検索ができるので、他のテスト項目書で書いたことのある手順などを流用したい時にすぐに探せる<br>　・テスト項目が縦一列に並ぶので可読性の点ではExcel に劣る<br>　・総じて熟練すると◎|
| レビュー指摘　・修正作業 | △     | ◎                             | Excel<br>　・修正箇所をGitで差分表示できないため、修正箇所を手作業で指し示す必要がある<br>　（番号指定や色を変えるなど)<br>　・レビュー指摘する側も同様<br>　・修正時に指摘箇所との番号がずれると面倒で気づくのに時間をロスする<br>　・少しの修正でも Excel ファイルを開く必要がある<br>（レビューアー、レビューイの双方）<br>Markdown<br>　・修正箇所をGitで差分表示でき、且つそこに直接レビューコメントを書ける<br>　・GitHub の Suggestion機能（提案機能）を使えば、複雑な指示を含む指摘を要したり、逆にどうでもよい誤字などを、レビューアーが直接テスト項目を修正してレビューイに提案でき、レビューイはそれを直接取り込める<br>　・ファイルを開かなくても直接 GHE上でレビューすることもできる |
| テスト実施作業     | ◎     | (△)                             | Excel<br>　・自動集計ができる<br>　・表形式なのでテスト項目の可読性がよく実施しやすい<br>Markdown<br>　・自動集計ができない（これを回避するためにこのツールを開発した ※設計は Markdown、実施は Excel で運用）                                                                                                                                                                                                                                                            |
| テスト観点俯瞰     | 〇     | ◯                             | Excel<br>　・マクロ操作でテスト観点（色付き帯）だけを見ることができる<br>Markdown<br>　・ Visual Studio Code の目次機能で目次を一度作成しておけば、いつでもテスト観点を閲覧できる                                                                                                                                                                                                                                    |
| テスト項目の閲覧性   | ◯     | △                             | Excel<br>　・テスト項目が表形式になっているので見やすい<br>Markdown<br>　・テスト項目が縦一列にならんでいるので、Excelよりは見にくい                                                                                                                                                                                                                                                                             |
| 項目書のカスタマイズ性 | ？     | ？                             | Excel<br>　・フォーマットを変えやすいが、ルール化があいまいだと逆にいろんな種類のフォーマットが増えてカオスになりやすい<br>Markdown<br>　・Markdown 記法の範疇でしか表現できないので、カスタマイズ性は低い（ただしMdToExcel変換ツールを改良すれば対応できる）が逆にフォーマットを統一しやすい                                                                                                                                                                                                       |
| レビュー記録の品質   | △     | ◎                             | Excel<br>　・手作業なのでヒューマンエラーが起こりやすく、検証報告書のレビューもそれによるロスが多くなる<br>　・テスト項目修正時に項目が増えたりすると、レビュー記録にある指摘箇所の番号と実際の指摘箇所の場所がずれることがあり、齟齬が生じたまま成果物となってしまう<br>　・指摘が書いてある場所と、指摘先が別シートなので経緯を追いにくい<br>Markdown<br>　・レビュー記録の管理は GHE がやってくれるので、ヒューマンエラーの入り込む余地はほとんどない<br>　・GHEのプルリク機能を利用し、テスト項目修正の差分に直接指摘を書き出しているので、テスト項目が増えても指摘箇所がずれることはない<br>　・指摘箇所 - 指摘部分 - 指摘コメント - 修正コメントが、１つの画面上に時系列で一列に並ぶので経緯を追いやすい |
| Git管理       | △     | ◎                             | Excel<br>　・差分管理できないので修正履歴を残す場合は、Excel（バイナリファイル）を毎回まるごとコミットすることになり、GHEのストレージを圧迫する<br>Markdown<br>　・差分管理できるのでGHEのストレージを圧迫しない  |

<br>

## 開発者向け

### 開発環境

- Python 3.6 or higher
  - 必要なライブラリは[requirements.txt](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/MarkdownToExcel/requirements.txt) を参照。  
    以下のコマンドで一括インストール可能。
    ```
    $ pip install -r requirements.txt
    ```

### 構成

主なファイルは以下の通り。

```
.
|-- dist                        # ビルド先フォルダ
|     |-- MdToExcel.exe         # 変換処理の実行ファイル
|     |-- test_spec_sample.md   # MardDown テスト項目書サンプル
|     |-- test_spec_sample.xlsx # 上記を Excel に変換したもの
|-- documents                   # 参考資料フォルダ
|-- resource                    # リソースフォルダ
|     |-- config.yaml           # 変換処理の設定ファイル
|     |-- st_template.xlsm      # Excel テスト項目書テンプレート
|-- excel_operator.py           # excel関係の処理 
|-- markdown_operator.py        # markdown関係の処理
|-- MdToExcel.py                # MAIN
|-- MdToExcel.spec              # ビルド用設定ファイル
|-- warningMsgProvider.py       # 変換時の警告・エラーメッセージの定義ファイル
|-- README.md                   # 説明
|-- requirements.txt            # 利用ライブラリ一覧
```

### 実行
以下のコマンドで変換処理を実行できます。  
```
$ python MdToExcel.py -f {テスト項目書のファイルパス}
```

### 実行ファイル(`exe`)のビルド
`MdToExcel.py` をビルドして `exe` 化します。  
これを利用することで `Python` がインストールされていない環境上でも実行できるようになります。
```
$ pyinstaller MdToExcel.spec
```
📔 `$ pyinstaller --onefile MdToExcel.py` でもビルド可能ですが、この場合リソースフォルダの内容が含まれません。  
 そのため、上記のように`spec` ファイルを利用してビルドしてください。（[参考](http://ghe.nanao.co.jp/SQG/Tools_ST/blob/master/MarkdownToExcel/documents/Pyinstaller%20%E3%81%A7%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9%E3%82%92%E5%90%AB%E3%82%81%E3%81%9Fexe%E3%82%92%E4%BD%9C%E6%88%90%E3%81%99%E3%82%8B%20-%20Qiita.pdf)）

### デバッグ
- `Visual Studio Code` の拡張機能に、`Python` をインストールする。
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/0579824c-f1cd-4c5e-b93f-e8eb63f06147)

- `.vscode/launch.json`を開いて、Excelに変換する Markdownファイルを指定します。（`args`のところ）
    ```
    {
        "version": "0.2.0",
        "configurations": [
            {
                "cwd": "${fileDirname}",    // カレントディレクトリをファイルがある位置に移動
                "args": ["-f", "StreamGateway.md"],       // 入力した引数つきで実行 ※カレントにある `StreamGateway.md` を Excel に変換する場合
                "name": "Python: Current File",
                "type": "python",
                "request": "launch",
                "program": "${file}",    // デバッグ実行するPythonファイル（エディタ上で選択されているファイルが該当するので MdToExcel.py を選択しておく）
                "console": "integratedTerminal"
            }
        ]
    }
    ```
- 上記で指定したMarkdownファイルを今からデバッグ実行するメインプログラム `MdToExcel.py`と同じ場所に配置します。

- `Visual Studio Code` 上で `MdToExcel.py` を開き、任意のコードにブレークポイントを設定する。  
 （画像は JavaScript のコードになっていますが、Python でも同様の手順です）
  ![image](http://ghe.nanao.co.jp/storage/user/110/files/9c17e0a1-9033-417f-8edb-12637af8c76d)

    :warning: `MdToExcel.py` 以外のモジュールにも当然ブレークポイントは設定できますが、必ずエディタ上で `MdToExcel.py` が選択されている状態でデバッグ実行を開始してください。

- 参考資料  

  <details>
    <summary> Visual Studio Code　～Pythonのデバッグ方法～ </summary>
    
    _出典: https://recruit.crdoti.co.jp/2022/12/visual-studio-code-python.html_
    ![image](http://ghe.nanao.co.jp/storage/user/110/files/847e9257-ab22-4bdd-a866-1d577b479e43)

  </details>



