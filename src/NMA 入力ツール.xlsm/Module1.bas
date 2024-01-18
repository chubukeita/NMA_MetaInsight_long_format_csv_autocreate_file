Attribute VB_Name = "Module1"
Option Explicit
Sub create()
    Application.ScreenUpdating = False
    
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("このマクロを実行する前に、必要なデータは全て保存してください。マクロを実行した場合、元データはなくなります。データは保存しましたか？", vbCritical + vbOKCancel, "警告アイコン")
    If rc = vbCancel Then
        Exit Sub
    End If

    Call outcome_Format
    
    Call leftsheet_delete
    
    Call outcome_sheet
    
    Call continuous_or_dichotomous
    
    Application.ScreenUpdating = True
    
    Call シート削除
    
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("input")
    wsInput.Activate
    
    MsgBox "完了"
End Sub

Sub outcome_Format()
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet
    Dim i As Long, k As Long, startCol As Long, lastCol As Long, Colwidth As Long
    Dim outcomeType As String
    
    ' ContinuousOutcomeの場合、結合したセルの列数が9
    ' DichotomousOutcomeの場合、結合したセルの列数が12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 12
    DichotomousWide = 9
    
    ' 入力シートを設定
    Set wsInput = Worksheets("input")

    ' "Strategies"がある列の位置を特定し、最終列を取得
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' 出力シートを探し、存在すれば削除
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOutcome_Format = Worksheets("outcome_format")
    If Not wsOutcome_Format Is Nothing Then
        wsOutcome_Format.Delete
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 新しいシートを作成して名前を設定
    Set wsOutcome_Format = Worksheets.Add(After:=Worksheets("リンク一覧"))
    wsOutcome_Format.name = "outcome_format"

    ' 出力シートにヘッダーを設定
    With wsOutcome_Format
        .Cells(2, 1).Value = "No"
        .Cells(2, 2).Value = "type"
        .Cells(2, 3).Value = "outcome"
    End With
    
    i = 3 ' データ行の開始位置

    ' アウトカムデータを処理
    For k = startCol To lastCol
        Colwidth = wsInput.Cells(3, k).MergeArea.Columns.Count
        
        With wsOutcome_Format
            ' No列に連番を設定
            .Cells(i, 1) = i - 2
            
            ' type列にアウトカムのタイプを設定
            Select Case Colwidth
                Case ContinuousWide
                    outcomeType = "Continuous"
                Case DichotomousWide
                    outcomeType = "Dichotomous"
                Case Else
                    outcomeType = ""
            End Select
            .Cells(i, 2) = outcomeType
            
            ' outcome列にアウトカム名を設定
            .Cells(i, 3) = wsInput.Cells(3, k).Value
        End With
        
        i = i + 1
        k = k + Colwidth - 1 ' マージされたセルの分だけインクリメント
    Next k
    Call Study_No_Assign
End Sub
Sub FindStrategiesAndSetColumns(wsInput As Worksheet, ByRef startCol As Long, ByRef lastCol As Long)
    Dim foundCell As Range

    ' "Strategies"の検索と列の特定
    Set foundCell = wsInput.Rows(4).Find("Strategies")
    If Not foundCell Is Nothing Then
        startCol = foundCell.Column + foundCell.MergeArea.Columns.Count
        lastCol = wsInput.Cells(5, wsInput.Columns.Count).End(xlToLeft).Column
    Else
        MsgBox "Strategies not found."
        Exit Sub
    End If
End Sub
Sub Study_No_Assign()
    Dim wsInput As Worksheet
    Dim i As Long, imax As Long
    
    ' 入力シートのセット
    Set wsInput = Worksheets("input")
    
    ' 最終行の取得
    imax = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).Row
    
    ' "Study No"の列の更新
    With wsInput
        For i = 6 To imax
            .Cells(i, 2).FormulaR1C1 = "=SUBTOTAL(3,R6C3:R" & i & "C3)"
        Next i
    End With
End Sub
Sub leftsheet_delete_alert()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("inputシートよりも左側のシートを削除します。よろしいですか？", vbCritical + vbOKCancel, "警告アイコン")
    If rc = vbCancel Then
        Exit Sub
    End If
    
    Call leftsheet_delete
End Sub

Sub leftsheet_delete()
    
    Dim ws As Worksheet, Target As String

    'このシートより左にあるシートが削除対象
    Target = "input"
    
    Application.DisplayAlerts = False
    'シートをループ
    For Each ws In Worksheets
        '「Target」シートじゃなければシート削除
        If ws.name <> Target Then
            ws.Delete
        Else
            '「Target」シートが出現したらループを抜ける
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True
    
End Sub

Sub outcome_sheet()
    
    ' 変数の宣言
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet, ws As Worksheet
    Dim k As Long, i As Long, j As Long
    Dim Colwidth As Long, FirstCol As Long, startCol As Long, lastCol As Long
    Dim maxInputRow As Long, maxOutcome_Format_Row As Long
    
    ' シートの設定
    Set wsInput = Worksheets("input")
    Set wsOutcome_Format = Worksheets("outcome_format")
    
    ' inputシートの最終行と最終列を取得
    With wsInput
        maxInputRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' outcome_formatシートの最終行を取得
    With wsOutcome_Format
        maxOutcome_Format_Row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    ' Strategiesが存在する列の取得
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' 既存のアウトカムシートを削除
    Application.DisplayAlerts = False
    For k = startCol To lastCol
        Colwidth = wsInput.Cells(3, k).MergeArea.Columns.Count
        Dim name As String
        name = wsInput.Cells(3, k).Value
        If Len(name) > 0 Then
            On Error Resume Next
            Worksheets(name).Delete
            On Error GoTo 0
        End If
        k = k + Colwidth - 1
    Next k
    Application.DisplayAlerts = True
    
    FirstCol = startCol
    
    ' 新しいアウトカムシートを追加
    For i = 1 To maxOutcome_Format_Row - 2
        Set ws = Worksheets.Add(Before:=Worksheets("2×4表"))
        With wsInput
            ' 共通の列をコピー
            .Range(.Cells(1, 1), .Cells(maxInputRow, FirstCol - 1)).Copy ws.Cells(1, 1)
            ' 各アウトカムのデータをコピー
            Colwidth = .Cells(3, startCol).MergeArea.Columns.Count
            .Range(.Cells(1, startCol), .Cells(maxInputRow, startCol + Colwidth - 1)).Copy ws.Cells(1, FirstCol)
        End With
        ws.name = wsOutcome_Format.Cells(i + 2, 3).Value ' アウトカム名でシートを命名
        startCol = startCol + Colwidth ' 次のアウトカム列へ
    Next i
    
End Sub

Sub continuous_or_dichotomous()

    ' 変数の宣言
    Dim j As Long, wide As Long, startCol As Long, lastCol As Long
    Dim name As String
    Dim wsInput As Worksheet
    
    ' ContinuousOutcomeの場合、結合したセルの列数が9
    ' DichotomousOutcomeの場合、結合したセルの列数が12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 12
    DichotomousWide = 9

    ' "input" シートを設定
    Set wsInput = Worksheets("input")

    ' "Strategies" セルを検索して、データ処理を開始する列を取得
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' startCol から lastCol までの列を繰り返し処理
    For j = startCol To lastCol
        ' 現在の列の幅（マージされたセルの数）と名前を取得
        With wsInput
            wide = .Cells(3, j).MergeArea.Columns.Count
            name = .Cells(3, j).Text
        End With
        
        ' 対応するワークシートをアクティブにし、参照を設定
        Worksheets(name).Activate
        
        ' 幅に応じて適切な処理を呼び出し
        If wide = ContinuousWide Then
            Call MetaInsightdataLONG_Continuous
        ElseIf wide = DichotomousWide Then
            Call MetaInsightdataLONG_Dichotomous
        End If
        
        ' 次のマージされたセルブロックに移動
        j = j + wide - 1
    Next j
End Sub
Sub MetaInsightdataLONG_Continuous()
    ' シート処理用の変数
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long
    
    ' シートへの参照を設定
    Set wsInput = Worksheets("input")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' 新しいシートの設定
    With wsNew
        .Range("A1:E1").Value = Array("Study", "T", "N", "Mean", "SD")
        ix = 2
    End With
    
    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With

    ' 研究データ用の変数 (4行目は"Strategies"で結合されたセルを含む)
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' データを横から縦に転置
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 4
                ' 基本的なstudyDataを抽出
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' 処理するデータがなくなったら終了
    
                ' 基本的なstudyDataからコメント文字列を作成
                Dim commentStr As String
                commentStr = Join(studyData, " ")
    
                ' データを書き込み、新しいシートにコメントを追加
                With wsNew
                    .Cells(ix, 1).Value = studyData(1) & " " & studyData(3) ' Study Name & Year
                    .Cells(ix, 1).AddComment commentStr
                    .Cells(ix, 2).Value = wsActive.Cells(i, j).Value      ' T (Treatment)
                    .Cells(ix, 3).Value = wsActive.Cells(i, j + 3).Value  ' N (Sample size)
                    .Cells(ix, 4).Value = wsActive.Cells(i, j + 1).Value  ' Mean
                    .Cells(ix, 5).Value = wsActive.Cells(i, j + 2).Value  ' SD
                    ix = ix + 1
                End With
            Next j
        Next i
    End With
    
     ' 列4の"NR"または空白を含む行をAutoFilterを使って削除
    Application.DisplayAlerts = False ' アラートダイアログを無効にする
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=4, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    
    ' 列5の"NR"または空白を含む行をAutoFilterを使って削除
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=5, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' アラートダイアログを再有効にする
    
    ' B1を除くB2から始まる2列目を右揃えに設定
    With wsNew.Range("B2:E" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).Row)
        .HorizontalAlignment = xlRight
    End With
End Sub
Sub MetaInsightdataLONG_Dichotomous()
    ' シート処理用の変数
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long

    ' シートへの参照を設定
    Set wsInput = Worksheets("input")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' 新しいシートの設定
    With wsNew
        .Range("A1:D1").Value = Array("Study", "T", "R", "N")
        ix = 2
    End With


    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' 研究データ用の変数 (4行目は"Strategies"で結合されたセルを含む)
    Call FindStrategiesAndSetColumns(wsActive, startCol, lastCol)

    ' データを横から縦に転置
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 3
                ' 基本的なstudyDataを抽出
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' 処理するデータがなくなったら終了

                ' 基本的なstudyDataからコメント文字列を作成
                Dim commentStr As String
                commentStr = Join(studyData, " ")

                ' データを書き込み、新しいシートにコメントを追加
                With wsNew
                    .Cells(ix, 1).Value = studyData(1) & " " & studyData(3) ' Study Name & Year
                    .Cells(ix, 1).AddComment commentStr
                    .Cells(ix, 2).Value = wsActive.Cells(i, j).Value      ' T (Treatment)
                    .Cells(ix, 3).Value = wsActive.Cells(i, j + 1).Value  ' R (Responders)
                    .Cells(ix, 4).Value = wsActive.Cells(i, j + 2).Value  ' N (Sample size)
                    ix = ix + 1
                End With
            Next j
        Next i
    End With
    
    ' 列3の"NR"または空白を含む行をAutoFilterを使って削除
    Application.DisplayAlerts = False ' アラートダイアログを無効にする
    With wsNew
        .AutoFilterMode = False
        .Range("A1:D1").AutoFilter Field:=3, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' アラートダイアログを再有効にする
    
    ' B1を除くB2から始まる2列目を右揃えに設定
    With wsNew.Range("B2:B" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).Row)
        .HorizontalAlignment = xlRight
    End With
End Sub

' 新しいシートを追加するか、既存のものを取得するヘルパー関数
Function AddOrGetSheet(name As String, Optional BeforeSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next ' エラーが発生しても次のコード行に進む
    Set ws = Worksheets(name)
    On Error GoTo 0 ' エラーハンドリングを標準に戻す
    If ws Is Nothing Then ' シートが存在しない場合、新しいシートを追加する
        Set ws = Worksheets.Add(Before:=BeforeSheet)
        ws.name = name
    End If
    Set AddOrGetSheet = ws ' シートを返す
End Function
Sub シート削除()
    Const START_SHEET_NAME As String = "input"  ' 削除開始シート
    Const END_SHEET_NAME As String = "2×4表"    ' 削除終了シート

    Dim startIndex As Long, endIndex As Long, i As Long, temp As Long

    ' シートインデックスを取得
    startIndex = Sheets(START_SHEET_NAME).Index
    endIndex = Sheets(END_SHEET_NAME).Index

    ' 開始と終了のインデックスを適切に設定
    If startIndex > endIndex Then
        temp = startIndex
        startIndex = endIndex
        endIndex = temp
    End If

    ' シートの削除
    Application.DisplayAlerts = False
    For i = endIndex - 1 To startIndex + 1 Step -1
        Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True
End Sub
Sub make_csv()
    Dim ans1 As Long, ans2 As Long, i As Long, j As Long
    Dim wsOutcome_Format As Worksheet
    Dim fileSaveName As Variant
    Dim newFileName As String, fileNamePath As String, fileNameOnly As String
    
    If Not ConfirmDataSave() Then Exit Sub
    
    Call outcome_Format
    
    Set wsOutcome_Format = Worksheets("outcome_format")
    Dim outcnt As Long
    With wsOutcome_Format
        outcnt = .Cells(.Rows.Count, 1).End(xlUp).Row - 2
    End With

    ans1 = GetSheetNumber("開始", outcnt)
    If ans1 = 0 Then Exit Sub
    
    ans2 = GetSheetNumber("終了", outcnt)
    If ans2 = 0 Then Exit Sub
    
    For i = ans1 To ans2
        newFileName = ThisWorkbook.Sheets(i).name
        fileSaveName = GetSaveAsFileName(newFileName)
        If fileSaveName = False Then Exit Sub
        
        ThisWorkbook.Sheets(i).Copy
        With ActiveWorkbook
            If FileExists(CStr(fileSaveName)) Then
                fileSaveName = GetUniqueFileName(CStr(fileSaveName))
            End If
            .SaveAs fileSaveName, FileFormat:=xlCSV, CreateBackup:=False
            .Close False
        End With
    Next i
End Sub

Function ConfirmDataSave() As Boolean
    Dim response As VbMsgBoxResult
    response = MsgBox("このマクロを実行すると、xlsmファイルのデータがなくなります。" & _
                      "csvファイル作成前にxlsmファイルを保存してください。" & _
                      "csvファイル出力はxlsmをコピーした後に行ってください。" & _
                      "コピーしましたか？", vbCritical + vbOKCancel, "警告")
    ConfirmDataSave = (response = vbOK)
End Function

Function GetSheetNumber(prompt As String, outcnt As Long) As Long
    Dim result As Variant
    result = InputBox("左から数えて、" & prompt & "シートは何番目ですか？" & _
                      "現在のアウトカム数は" & outcnt & "です。" & _
                      "特定のアウトカムだけを指定する場合は、同じ数字を入力してください。", _
                      "csvにする範囲", "")
    If StrPtr(result) = 0 Then
        GetSheetNumber = 0
    Else
        GetSheetNumber = CLng(result)
    End If
End Function

Function GetSaveAsFileName(newFileName As String) As Variant
    GetSaveAsFileName = Application.GetSaveAsFileName( _
                        InitialFileName:=newFileName & ".csv", _
                        FileFilter:="CSVファイル(*.csv),*.csv", _
                        FilterIndex:=1, _
                        Title:="保存ファイルの指定")
End Function

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Function GetUniqueFileName(filePath As String) As String
    Dim fileNameOnly As String, fileNamePath As String, fileExtension As String
    Dim counter As Long
    fileNameOnly = Dir(filePath)
    fileNamePath = Replace(filePath, fileNameOnly, "")
    fileExtension = ".csv"
    fileNameOnly = Replace(fileNameOnly, fileExtension, "")
    
    counter = 2
    Do While FileExists(filePath)
        filePath = fileNamePath & fileNameOnly & " (" & counter & ")" & fileExtension
        counter = counter + 1
    Loop
    GetUniqueFileName = filePath
End Function
Sub HideColumns()
    Dim j As Long, imax As Long, jmax As Long

    ' 最終行と最終列を取得
    With ActiveSheet
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
        jmax = .Cells(5, .Columns.Count).End(xlToLeft).Column
        
        ' 対象列の処理
        For j = 2 To jmax
            ' 色で塗りつぶされていない列を非表示にする
            .Columns(j).Hidden = (.Range(.Cells(6, j), .Cells(imax, j)).Interior.ColorIndex = xlNone)
        Next j
    End With
End Sub
Sub AppearColumns()
    Dim ws1 As Worksheet
    Set ws1 = ActiveSheet
    ws1.Columns.Hidden = False
End Sub
Sub PMID_search()
    Call PMID_Create
'    Call PMID_OPEN
End Sub
Sub PMID_Create()
    Dim wsInput As Worksheet, wsLink_List As Worksheet
    Dim i As Long, imax As Long
    Dim PMID As String, url As String
    
    Application.ScreenUpdating = False
    
    ' ワークシートの設定
    Set wsInput = Worksheets("input")
    Set wsLink_List = EnsureWorksheet("リンク一覧", "outcome_type")

    ' URLの基本形
    url = "https://pubmed.ncbi.nlm.nih.gov/"
    
    With wsLink_List
        ' ヘッダーのコピー
        wsInput.Range("B5:G5").Copy .Cells(5, 2)
        wsLink_List.Cells(5, 8).Value = "リンク"
    
        ' データとハイパーリンクの作成
        imax = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).Row
    
    
        For i = 6 To imax
            wsInput.Range(wsInput.Cells(i, 2), wsInput.Cells(i, 7)).Copy .Cells(i, 2)
            PMID = wsInput.Cells(i, 4).Value
            .Cells(i, 8).Value = url & PMID & "/"
            .Hyperlinks.Add Anchor:=.Cells(i, 8), Address:=.Cells(i, 8).Value
            .Hyperlinks.Add Anchor:=.Cells(i, 4), Address:=.Cells(i, 8).Value
        Next i
    End With
    
    Application.ScreenUpdating = True
    wsLink_List.Activate
End Sub

' ワークシートが存在しない場合は新しく作成する
Function EnsureWorksheet(sheetName As String, afterSheet As String) As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(afterSheet))
        ws.name = sheetName
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set EnsureWorksheet = ws
End Function

Sub PMID_OPEN()
    Dim i As Long, imax As Long
    Dim wsInput As Worksheet, wsLink_List As Worksheet
    
    Set wsInput = Worksheets("input")
    Set wsLink_List = Worksheets("リンク一覧")
    
    With wsInput
        ' 最大行数を取得
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' 各行のPubMedリンクを開く
    For i = 6 To imax
        Dim url As String
        url = wsLink_List.Cells(i, 8)
        
        ' Google Chromeを起動し指定URLに移動
        CreateObject("WScript.Shell").Run ("chrome.exe -url " & url)
    Next i
End Sub

Sub PMID_OR_Search_Expression()
    Call CheckPMIDDuplicatesOrMissing
    ufSet4.Show vbModeless
    ufSet4.Repaint
    Call ChromeDriverUpdate
    Call 検索式作成
End Sub
Sub CheckPMIDDuplicatesOrMissing()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 現在のシートを使用する

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row ' D列の最終行を取得

    Dim PMID As Variant
    Dim PMIDDict As Object
    Set PMIDDict = CreateObject("Scripting.Dictionary") ' 辞書オブジェクトの作成

    Dim i As Long
    Dim isProblemFound As Boolean
    isProblemFound = False

    ' D6セル以降のPMIDをチェック
    For i = 6 To lastRow
        PMID = ws.Cells(i, 4).Value
        If Not IsEmpty(PMID) Then
            If PMIDDict.Exists(PMID) Then
                ' 重複が見つかった
                isProblemFound = True
                Exit For
            Else
                PMIDDict.Add PMID, True
            End If
        End If
    Next i

    ' エラーメッセージの表示
    If isProblemFound Then
        MsgBox "PMIDに重複もしくは抜けが有ります。重複もしくは抜けがあると検索式の作成ができません", vbExclamation, "エラー発生"
    End If
End Sub
'// SeleniumBasic
Public Sub ChromeDriverUpdate()
    Dim Driver As New Selenium.ChromeDriver
    SafeOpen Driver, Chrome
    Driver.Get "https://www.google.co.jp/?q=selenium"
    Driver.Wait 3000
    Driver.Quit
    ' MsgBox ("Chrome Driverの更新完了")
End Sub

Sub 検索式作成()
    Dim i As Long, j As Long, k As Long, imax As Long
    Dim wsInput As Worksheet, wsLink_List As Worksheet
    Dim Driver As Selenium.WebDriver
    Dim Keys As Keys

    Set wsInput = Worksheets("input")
    Set wsLink_List = Worksheets("リンク一覧")
    Set Driver = New Selenium.WebDriver
    Set Keys = New Keys
    
    With wsInput
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With

    ' WebDriverの初期設定と起動
    Driver.Start "chrome"
    Driver.Get "https://pubmed.ncbi.nlm.nih.gov/advanced/"

    ' PMIDを検索して結果を取得
    For i = 6 To imax
        Dim PMID As String
        PMID = wsLink_List.Cells(i, 4)
        
        Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
        Driver.FindElementByCss("#query-box-input").SendKeys PMID
        Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    Next i
    
            
    Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
            
    ' 検索結果を組み合わせ
    k = imax - 5
    For j = imax To 6 Step -1
        Dim menuCss As String, orCss As String
        menuCss = "#search-history-table > tbody > tr:nth-child(" & k & ") > td.dt-center.history-actions.dropdown-block > div > button"
        orCss = "#search-history-table > tbody > tr:nth-child(" & k & ") > td.dt-center.history-actions.dropdown-block > div > div > button:nth-child(3)"
    

        Driver.FindElementByCss(menuCss).Click
        If j = imax Then
            Driver.FindElementByCss("#search-history-table > tbody > tr:nth-child(" & k & ") > td.dt-center.history-actions.dropdown-block > div > div > button.add-without-boolean.action-menu-item").Click
        Else
            Driver.FindElementByCss(orCss).ScrollIntoView True
            Driver.Wait 200
            Driver.FindElementByCss(orCss).Click
        End If

        k = k - 1
    Next j

    ' 最終的な検索結果を取得
    Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    wsLink_List.Cells(1, 1) = "全てのPMIDをORで結合した場合の検索式"
    wsLink_List.Cells(2, 1) = Driver.FindElementByCss("#id_term").Value
    
    ' WebDriverを閉じる
    Driver.Quit
    
    wsLink_List.Activate
    Unload ufSet4
    
    MsgBox "「" & wsLink_List.name & "」シートの" & wsLink_List.Cells(2, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セルに、全てのPMIDをORで結合した場合の検索式を出力しました。"
End Sub



