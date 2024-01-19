# NMA_MetaInsight_long_format_csv_autocreate_file
Network Meta-analysisを実施するためのソフトウェアの一つにMetaInsight (https://crsu.shinyapps.io/MetaInsight/) があります。これには、Long format形式か、Wide format形式のcsvファイルをアップロードする必要があります。ただし、データを手入力してこれらの形式のcsvファイルを、対象のアウトカムの数だけ作成するのは手間がかかります。よって、より直感的でデータの入力がしやすいフォーマットを作成し、入力データに基づいてLong format形式のcsvファイルに変換するマクロ有効ブックを、ChatGPT 4.0によるリファクタリングを経て実装しました（以下，入力ツール）。入力ツールによりLong format形式のcsvをボタン操作で作成できます。

## 初期設定
マクロ有効ブック内の「開発」タブを選択し、「Visual Basic」を選択し、「ツール (T)」の「参照設定」から「Selenium Type Library」を選択して「Selenium」を有効にしてください。

## 注意事項
入力ツールは、基本的に「InputSheet」を操作し、データを入力していきます。黄色のセルは基本的にExcelの数式が入っていないセルなので、状況に応じて入力値を変更することができます。色のついていないセルは基本的にExcelの数式が入っているためセルを編集しないようにしてください。また、全てのシートに対して行や列の挿入を行うとデータが壊れる可能性がありますので、避けてください。


## 基本の使い方（NMAで比較する群が8群を超えない場合の使用法）
## ※8群を超える場合には、READMEファイルの「NMAで比較する群の総数が8群を超える場合の使い方」も合わせて参照してください）
はじめに、「SortSheet」に配置された3つのコマンドボタンで群の分類の種類を指定することができます。

### 「SortSheet」による群 (arm) の分類
MetaInsightに出力される際の各群のラベル名は、「SortSheet」のC3セル～C10セルの入力値になります。「classification into single or combined groups」ボタンをクリックしたときは、D2セル、E2セル、F2セルから入力値を変更し、「classified into the usual 8 groups」ボタン、「classification by Threshold」ボタンをクリックしたときは、C3セルからC10セルから直接入力値を変更できます。

「classification into single or combined groups」ボタンをクリックすると、D2セル、E2セル、F2セルに入力された介入の単独群、併用群に振り分けることができます。この状態ではD2セル、E2セル、F2セルの入力値を変更することができ、それに応じてC列のarmのラベルが変わります。各armに対応するPEEP、VTの値を、「InputSheet」に入力することで間接的にarmを指定します。
「classified into the usual 8 groups」ボタンをクリックすると、C3セルからC10セルのarmに分類することができます。各armに対応するPEEP、VTの値を、「InputSheet」に入力することで間接的にarmを指定します。

「classification by Threshold」ボタンをクリックすると、IndexAを4段階、IndexBを2段階に分類するための「Threshold settings」フォームが立ち上がり、「IndexA」タブ、「IndexB」タブの順に閾値を設定できます。「IndexA」シートにおいて、IndexAを表（B4セルからF9セルまでの範囲）の閾値に従って4段階に区分します。

まず「E7セルの値」を入力し、リストから「less than」もしくは「or below」を選択します。これにより、「Low」の上限の閾値を決めます。同様に、「E8セルの値」を入力し、リストから「less than」もしくは「or below」を選択します。これにより、「Intermediate」の上限の閾値を決めます。「IndexInput」ボタンを押すとフォームの情報が「IndexA」シートの表に反映されます。同様に、「IndexB」タブにおいても、「E8セルの値」を入力し、リストから「less than」もしくは「or below」を選択します。そして、「IndexInput」ボタンを押すと、フォームの情報が「InputB」シートの表に反映されます。各論文中で報告された2つのパラメータの値を「IndexA」、「IndexB」として捉え、「InputSheet」シートのIndexA、IndexBに入力していくと「Threshold settings」フォームで設定した閾値に応じて自動的にarmが分類されるようになっています。

### 「InputSheet」での入力
「InputSheet」の3行目のうち、黄色のセルは名前を変更でき、Z列以降でアウトカム名の変更ができます。「InputSheet」では表をダブルクリックすることで、入力フォーム (Input Form) が表示されます。そのフォームを使ってデータを入力していくことができます。現在入力中の行は水色で示されるようになっています。基本的には「InputSheet」においては、セルの値を変更していいセルには色が塗られており、逆に色が塗られていないセルは数式が入力されているためセルの値を変更しないようにします。また、タブの切り替えは「Ctrl+Tab」で一つ右に、「Ctrl+Shift+Tab」で一つ左に移動させることができます。

#### 「information」タブ
Input Formの一番左のタブ（デフォルトでは「Information」）は、研究の基本情報を入力できます。「StudyNo」は「InputSheet」の上から何番目の研究なのかを番号で示しています。また、「Authors」、「PMID」、「Year」、「Country」、「Research Period」は、それぞれ論文の筆頭著者、PubMed ID、論文の出版年、出版国、研究期間を入力できる。MetaInsightでのNMAに必須の入力事項は、「StudyNo」、「Authors」となります。

また、「Back」、「Next」のボタンを押すと、選択中のセルを一つ上、一つ下に移動させることができます。そして、「Update」ボタンを押すと、フォームで入力した内容を更新できます。なお、フォームで入力し直した際に「Update」ボタンを押さずに選択中のセルを移動させてしまうとシートに値が更新されないため、フォームで入力値を変更したら必ず「Update」ボタンを押すように注意する必要があります。さらに、「Outcome_Add」ボタンを押すと、「Outcome Setting」というフォームが表示され、「InputSheet」の最右端に任意のアウトカムを追加できます。対象とするアウトカムが2値データのアウトカムならリストから「Dichotomous」、連続量のアウトカムなら「Continuous」を選択します。また、対象とするアウトカム名を、「outcome_name」と書かれたテキストボックスに入力し、「Outcome_Insort」ボタンをクリックすることでアウトカムを追加できます。「Add」ボタンを押すことで、文字の入力された最終行の1つ下の行に移動でき、新たな論文の入力ができます。ここで、表の最終行まで入力されている場合には、その1行下に新たに自動で入力欄が作られるようになっています。「RowDelete」ボタンは、「InputSheet」内の表中の選択行を削除することができます。

#### 「Strategies」タブ
Input Formの左から2番目のタブ（デフォルトでは「Strategies」）は、研究のPEEP、VT、各群のサンプルサイズを入力できます。「treatment1」、「treatment2」、「treatment3」という「介入1」、「介入2」、「介入3」といった最大3群比較の研究を入力できます。入力フォームはtreatment1の項目を縦に見て「PEEP1」、「VT1」、「Patients (n1)」と並んでおり、1つ目の群のPEEP、VT、サンプルサイズを表しています。また。添え字の数字は、どの群に属するかを表しています。不明な値がある場合には、空白のままにする。「PEEP」、「VT」を入力の上、Informationタブにある「更新」ボタンをクリックすると、「treatment」に加えて、各群の名前 (arm) が更新される仕組みになっています。

#### 「outcome」タブ
Input Formの左から3番目以降のタブは、連続量のアウトカムか、2値のアウトカムかで入力できる項目が異なります。共通して入力できる項目として、各アウトカムごとのサンプルサイズ (n) があります。「Strategies」タブで入力した「Patients (n)」との違いは、臨床試験中に生じたプロトコール逸脱者を考慮するかどうかです。臨床試験中に生じた脱落者を考慮した場合、もともとプロトコール段階で割り振っていた各群のサンプルサイズと異なる可能性が考えられます。基本的には「Strategies」タブで入力した「Patients (n)」と一致していることが大半であるが、「Strategies」タブで入力した「Patients (n)」と異なる場合には、「n」に異なる数値を入力する必要があります。「Change n of treatment」にチェックを入れた上で「n」の数値を変更することができます。チェックを入れた状態で「Update」ボタンを押すと、選択中のアウトカムの「input」シートの表の「Patients (n)」のセルが赤くなり、もともと入っていた数式は値貼り付けによって上書きされます。もし誤ってチェックを入れた状態で「Update」ボタンを押してしまった場合には、誤って入力してしまった赤いセルの列のうち、数式の入っている色の塗りつぶされていないセルをコピーして、赤くなったセルに貼り付けることで元に戻すことができます。

##### 連続量のアウトカム (Continuous Outcome)
連続量のアウトカムのoutcomeタブでは、「μ (mean)」、「±SD」に、平均値と標準偏差を持つアウトカムを入力することができます。ここでは、不明な値が有ればNRを入力し、空欄にしないようにしてください。

##### 2値のアウトカム (Dichotomous Outcome)
2値のアウトカムのoutcomeタブでは、「event」に、イベント数を入力することができます。ここでは、不明な値があれば「NR」を入力し、空欄にしないようにしてください。

#### 「create sheets for each outcome」ボタン
全てのデータが入力できたら、InputSheetの表の下にある「create sheets for each outcome」ボタンを押して、MetaInsightで用いられるLong formatの形式に変換します。ボタンを押すと、「Before executing this macro, please save all necessary data. When the macro is executed, the original data will be lost. Have you saved your data?」という警告が表示されるが、「OK」をクリックして先へ進みます。「complete」というメッセージが出て「OK」を押したら、inputシートよりも左側のシートに「（アウトカム名） table」というシートが出現します。なお、これらのシートは、MetaInsightのLong Format形式になったシートの集合です。ここで、各tableシートのA列のA2セル以降にはコメントがついており、inputシートの「StudyNo」、「Authors」、「PMID」、「Year」が並んでおり情報を確認できます。

#### 「csv file output」ボタン
「csv file output」ボタンは、InputSheetよりも左にあるシートをまとめてcsvファイルに変換する機能を持ちます。ボタンを押すと、「Executing this macro will result in the loss of data in the xlsm file. Please save the xlsm file and copy the file before creating the csv file. The csv file output should be done after copying the xlsm file. Have you copied it?」という警告が表示されるため、警告に従って一度xlsmファイルを保存し、ファイルのコピーで複製する。再度、複製したファイルにて同じ操作を行い、警告に対し、「OK」をクリックして先へ進みます。「Scope to csv」という案内が表示されるので、すべてのアウトカムをcsvで出力する場合は、startは「1」、endは「（アウトカムの数）」を入力します。そうすると、「Specify the file to save」というダイアログボックスが表示され任意の場所に保存できます。デフォルトでは、csvのファイル名は「（アウトカム名） table.csv」として保存されるようになっており、保存先に同じファイル名のcsvファイルがあった場合には、「（アウトカム名） table (2).csv」、「（アウトカム名） table (3).csv」と番号が振られるようになっています。

#### その他のボタン
「Link_List create」のボタンを押すと、「Link_List」シートに、その論文のPubMedのサイトにアクセスするためのURLを自動作成できます。また、「Connecting PMIDs with ORs」ボタンは、inputシートに入力された文献全てのPMIDをORでつないだ検索式を作成できます。「Link_List」シートのA2セルにその検索式が表示されます。ただし、Chromeを自動操作するため、自動操作中はブラウザを勝手に操作しないように注意します。また、このORでつながっている検索式は、SRのために作った検索式が調査対象の全ての文献を含むかどうかを確かめる際に必要となります。なお、文献検索のために自作した検索式は、少なくとも同じCQを持つ先行のSRで収集された論文全てが拾えるように作る必要があります。具体的には、「先行のSRで収集された各論文のPMIDをORでつないだ検索式」と「自作した検索式」をANDでつないだ検索式の検索結果の総数が、「SRで収集された論文数」と一致しているか否かで網羅的検索ができているかを確認できます。論文数が一致しない場合は、自作した検索式を作りなおし、再度同じ工程を踏み、論文数が一致するまで検索式を試行錯誤しながら作っていきます。

「delete the table on the left side of InputSheet」ボタンは、InputSheetよりも左に作成された「（アウトカム名） table」のシートの集合をまとめて消去できます。「hide uncolored columns」、「undo hiding of columns」は、InputSheet上の色のついていないセルを非表示にしたり、再表示したりすることができます。

## NMAで比較する群の総数が8群を超える場合の使い方
お手数ですが、NMAの比較する群の総数が8群を超える場合には、InputSheetのW～Y列（「Strategies」欄の「arm1」から「arm3」の列）の6行目以降に振り分ける群の名前を直接入力してください。W～Y列の6行目以降にもともと入っているセルの数式は削除して上書きする形で群の名前を入力してください。ただし、InputSheetのセルをダブルクリックした際に表示されるフォーム「Input Form」の「Strategies」タブの「IndexA」、「IndexB」に入力しても群名が変更されることはなくなりますが、「Patient (n)」のデータは入力値を変更すると、InputSheetに反映されます。

## ユーザーフォームのレイアウト
ユーザーフォームのラベル、テキストボックス、コンボボックス、コマンドボタンの配置や、コントロール名などは、スライドのファイル名「UserFormLayout.pptx」から確認できます。

## ライセンス
For this program (except Module4), this code was created by chubukeita and refactored by ChatGPT 4.0.

Copyright (c) chubukeita, subject to MIT License.

More information about the new license can be found at the following link: 

https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

For Module4, the code for this module is taken from https://github.com/yamato1413/WebDriverManager-for-VBA.

著作権 (c) yamato1413, MITライセンスに従います。

ライセンスの全文は以下のリンクから参照できます: https://github.com/yamato1413/WebDriverManager-for-VBA/blob/main/LICENSE

変更箇所: https://github.com/yamato1413/WebDriverManager-for-VBA/blob/main/WebDriverManager4SeleniumBasic.bas および README の '// SeleniumBasic' の Sampleコードをそのまま使用しています。
