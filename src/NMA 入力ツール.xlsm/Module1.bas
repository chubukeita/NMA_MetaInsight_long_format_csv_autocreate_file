Attribute VB_Name = "Module1"
Option Explicit
Sub create()
    Application.ScreenUpdating = False
    
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("���̃}�N�������s����O�ɁA�K�v�ȃf�[�^�͑S�ĕۑ����Ă��������B�}�N�������s�����ꍇ�A���f�[�^�͂Ȃ��Ȃ�܂��B�f�[�^�͕ۑ����܂������H", vbCritical + vbOKCancel, "�x���A�C�R��")
    If rc = vbCancel Then
        Exit Sub
    End If

    Call outcome_Format
    
    Call leftsheet_delete
    
    Call outcome_sheet
    
    Call continuous_or_dichotomous
    
    Application.ScreenUpdating = True
    
    Call �V�[�g�폜
    
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("input")
    wsInput.Activate
    
    MsgBox "����"
End Sub

Sub outcome_Format()
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet
    Dim i As Long, k As Long, startCol As Long, lastCol As Long, Colwidth As Long
    Dim outcomeType As String
    
    ' ContinuousOutcome�̏ꍇ�A���������Z���̗񐔂�9
    ' DichotomousOutcome�̏ꍇ�A���������Z���̗񐔂�12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 12
    DichotomousWide = 9
    
    ' ���̓V�[�g��ݒ�
    Set wsInput = Worksheets("input")

    ' "Strategies"�������̈ʒu����肵�A�ŏI����擾
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' �o�̓V�[�g��T���A���݂���΍폜
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOutcome_Format = Worksheets("outcome_format")
    If Not wsOutcome_Format Is Nothing Then
        wsOutcome_Format.Delete
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' �V�����V�[�g���쐬���Ė��O��ݒ�
    Set wsOutcome_Format = Worksheets.Add(After:=Worksheets("�����N�ꗗ"))
    wsOutcome_Format.name = "outcome_format"

    ' �o�̓V�[�g�Ƀw�b�_�[��ݒ�
    With wsOutcome_Format
        .Cells(2, 1).Value = "No"
        .Cells(2, 2).Value = "type"
        .Cells(2, 3).Value = "outcome"
    End With
    
    i = 3 ' �f�[�^�s�̊J�n�ʒu

    ' �A�E�g�J���f�[�^������
    For k = startCol To lastCol
        Colwidth = wsInput.Cells(3, k).MergeArea.Columns.Count
        
        With wsOutcome_Format
            ' No��ɘA�Ԃ�ݒ�
            .Cells(i, 1) = i - 2
            
            ' type��ɃA�E�g�J���̃^�C�v��ݒ�
            Select Case Colwidth
                Case ContinuousWide
                    outcomeType = "Continuous"
                Case DichotomousWide
                    outcomeType = "Dichotomous"
                Case Else
                    outcomeType = ""
            End Select
            .Cells(i, 2) = outcomeType
            
            ' outcome��ɃA�E�g�J������ݒ�
            .Cells(i, 3) = wsInput.Cells(3, k).Value
        End With
        
        i = i + 1
        k = k + Colwidth - 1 ' �}�[�W���ꂽ�Z���̕������C���N�������g
    Next k
    Call Study_No_Assign
End Sub
Sub FindStrategiesAndSetColumns(wsInput As Worksheet, ByRef startCol As Long, ByRef lastCol As Long)
    Dim foundCell As Range

    ' "Strategies"�̌����Ɨ�̓���
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
    
    ' ���̓V�[�g�̃Z�b�g
    Set wsInput = Worksheets("input")
    
    ' �ŏI�s�̎擾
    imax = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).Row
    
    ' "Study No"�̗�̍X�V
    With wsInput
        For i = 6 To imax
            .Cells(i, 2).FormulaR1C1 = "=SUBTOTAL(3,R6C3:R" & i & "C3)"
        Next i
    End With
End Sub
Sub leftsheet_delete_alert()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("input�V�[�g���������̃V�[�g���폜���܂��B��낵���ł����H", vbCritical + vbOKCancel, "�x���A�C�R��")
    If rc = vbCancel Then
        Exit Sub
    End If
    
    Call leftsheet_delete
End Sub

Sub leftsheet_delete()
    
    Dim ws As Worksheet, Target As String

    '���̃V�[�g��荶�ɂ���V�[�g���폜�Ώ�
    Target = "input"
    
    Application.DisplayAlerts = False
    '�V�[�g�����[�v
    For Each ws In Worksheets
        '�uTarget�v�V�[�g����Ȃ���΃V�[�g�폜
        If ws.name <> Target Then
            ws.Delete
        Else
            '�uTarget�v�V�[�g���o�������烋�[�v�𔲂���
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True
    
End Sub

Sub outcome_sheet()
    
    ' �ϐ��̐錾
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet, ws As Worksheet
    Dim k As Long, i As Long, j As Long
    Dim Colwidth As Long, FirstCol As Long, startCol As Long, lastCol As Long
    Dim maxInputRow As Long, maxOutcome_Format_Row As Long
    
    ' �V�[�g�̐ݒ�
    Set wsInput = Worksheets("input")
    Set wsOutcome_Format = Worksheets("outcome_format")
    
    ' input�V�[�g�̍ŏI�s�ƍŏI����擾
    With wsInput
        maxInputRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' outcome_format�V�[�g�̍ŏI�s���擾
    With wsOutcome_Format
        maxOutcome_Format_Row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    ' Strategies�����݂����̎擾
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' �����̃A�E�g�J���V�[�g���폜
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
    
    ' �V�����A�E�g�J���V�[�g��ǉ�
    For i = 1 To maxOutcome_Format_Row - 2
        Set ws = Worksheets.Add(Before:=Worksheets("2�~4�\"))
        With wsInput
            ' ���ʂ̗���R�s�[
            .Range(.Cells(1, 1), .Cells(maxInputRow, FirstCol - 1)).Copy ws.Cells(1, 1)
            ' �e�A�E�g�J���̃f�[�^���R�s�[
            Colwidth = .Cells(3, startCol).MergeArea.Columns.Count
            .Range(.Cells(1, startCol), .Cells(maxInputRow, startCol + Colwidth - 1)).Copy ws.Cells(1, FirstCol)
        End With
        ws.name = wsOutcome_Format.Cells(i + 2, 3).Value ' �A�E�g�J�����ŃV�[�g�𖽖�
        startCol = startCol + Colwidth ' ���̃A�E�g�J�����
    Next i
    
End Sub

Sub continuous_or_dichotomous()

    ' �ϐ��̐錾
    Dim j As Long, wide As Long, startCol As Long, lastCol As Long
    Dim name As String
    Dim wsInput As Worksheet
    
    ' ContinuousOutcome�̏ꍇ�A���������Z���̗񐔂�9
    ' DichotomousOutcome�̏ꍇ�A���������Z���̗񐔂�12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 12
    DichotomousWide = 9

    ' "input" �V�[�g��ݒ�
    Set wsInput = Worksheets("input")

    ' "Strategies" �Z�����������āA�f�[�^�������J�n�������擾
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' startCol ���� lastCol �܂ł̗���J��Ԃ�����
    For j = startCol To lastCol
        ' ���݂̗�̕��i�}�[�W���ꂽ�Z���̐��j�Ɩ��O���擾
        With wsInput
            wide = .Cells(3, j).MergeArea.Columns.Count
            name = .Cells(3, j).Text
        End With
        
        ' �Ή����郏�[�N�V�[�g���A�N�e�B�u�ɂ��A�Q�Ƃ�ݒ�
        Worksheets(name).Activate
        
        ' ���ɉ����ēK�؂ȏ������Ăяo��
        If wide = ContinuousWide Then
            Call MetaInsightdataLONG_Continuous
        ElseIf wide = DichotomousWide Then
            Call MetaInsightdataLONG_Dichotomous
        End If
        
        ' ���̃}�[�W���ꂽ�Z���u���b�N�Ɉړ�
        j = j + wide - 1
    Next j
End Sub
Sub MetaInsightdataLONG_Continuous()
    ' �V�[�g�����p�̕ϐ�
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long
    
    ' �V�[�g�ւ̎Q�Ƃ�ݒ�
    Set wsInput = Worksheets("input")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' �V�����V�[�g�̐ݒ�
    With wsNew
        .Range("A1:E1").Value = Array("Study", "T", "N", "Mean", "SD")
        ix = 2
    End With
    
    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With

    ' �����f�[�^�p�̕ϐ� (4�s�ڂ�"Strategies"�Ō������ꂽ�Z�����܂�)
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' �f�[�^��������c�ɓ]�u
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 4
                ' ��{�I��studyData�𒊏o
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' ��������f�[�^���Ȃ��Ȃ�����I��
    
                ' ��{�I��studyData����R�����g��������쐬
                Dim commentStr As String
                commentStr = Join(studyData, " ")
    
                ' �f�[�^���������݁A�V�����V�[�g�ɃR�����g��ǉ�
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
    
     ' ��4��"NR"�܂��͋󔒂��܂ލs��AutoFilter���g���č폜
    Application.DisplayAlerts = False ' �A���[�g�_�C�A���O�𖳌��ɂ���
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=4, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    
    ' ��5��"NR"�܂��͋󔒂��܂ލs��AutoFilter���g���č폜
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=5, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' �A���[�g�_�C�A���O���ėL���ɂ���
    
    ' B1������B2����n�܂�2��ڂ��E�����ɐݒ�
    With wsNew.Range("B2:E" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).Row)
        .HorizontalAlignment = xlRight
    End With
End Sub
Sub MetaInsightdataLONG_Dichotomous()
    ' �V�[�g�����p�̕ϐ�
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long

    ' �V�[�g�ւ̎Q�Ƃ�ݒ�
    Set wsInput = Worksheets("input")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' �V�����V�[�g�̐ݒ�
    With wsNew
        .Range("A1:D1").Value = Array("Study", "T", "R", "N")
        ix = 2
    End With


    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' �����f�[�^�p�̕ϐ� (4�s�ڂ�"Strategies"�Ō������ꂽ�Z�����܂�)
    Call FindStrategiesAndSetColumns(wsActive, startCol, lastCol)

    ' �f�[�^��������c�ɓ]�u
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 3
                ' ��{�I��studyData�𒊏o
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' ��������f�[�^���Ȃ��Ȃ�����I��

                ' ��{�I��studyData����R�����g��������쐬
                Dim commentStr As String
                commentStr = Join(studyData, " ")

                ' �f�[�^���������݁A�V�����V�[�g�ɃR�����g��ǉ�
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
    
    ' ��3��"NR"�܂��͋󔒂��܂ލs��AutoFilter���g���č폜
    Application.DisplayAlerts = False ' �A���[�g�_�C�A���O�𖳌��ɂ���
    With wsNew
        .AutoFilterMode = False
        .Range("A1:D1").AutoFilter Field:=3, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' �A���[�g�_�C�A���O���ėL���ɂ���
    
    ' B1������B2����n�܂�2��ڂ��E�����ɐݒ�
    With wsNew.Range("B2:B" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).Row)
        .HorizontalAlignment = xlRight
    End With
End Sub

' �V�����V�[�g��ǉ����邩�A�����̂��̂��擾����w���p�[�֐�
Function AddOrGetSheet(name As String, Optional BeforeSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next ' �G���[���������Ă����̃R�[�h�s�ɐi��
    Set ws = Worksheets(name)
    On Error GoTo 0 ' �G���[�n���h�����O��W���ɖ߂�
    If ws Is Nothing Then ' �V�[�g�����݂��Ȃ��ꍇ�A�V�����V�[�g��ǉ�����
        Set ws = Worksheets.Add(Before:=BeforeSheet)
        ws.name = name
    End If
    Set AddOrGetSheet = ws ' �V�[�g��Ԃ�
End Function
Sub �V�[�g�폜()
    Const START_SHEET_NAME As String = "input"  ' �폜�J�n�V�[�g
    Const END_SHEET_NAME As String = "2�~4�\"    ' �폜�I���V�[�g

    Dim startIndex As Long, endIndex As Long, i As Long, temp As Long

    ' �V�[�g�C���f�b�N�X���擾
    startIndex = Sheets(START_SHEET_NAME).Index
    endIndex = Sheets(END_SHEET_NAME).Index

    ' �J�n�ƏI���̃C���f�b�N�X��K�؂ɐݒ�
    If startIndex > endIndex Then
        temp = startIndex
        startIndex = endIndex
        endIndex = temp
    End If

    ' �V�[�g�̍폜
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

    ans1 = GetSheetNumber("�J�n", outcnt)
    If ans1 = 0 Then Exit Sub
    
    ans2 = GetSheetNumber("�I��", outcnt)
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
    response = MsgBox("���̃}�N�������s����ƁAxlsm�t�@�C���̃f�[�^���Ȃ��Ȃ�܂��B" & _
                      "csv�t�@�C���쐬�O��xlsm�t�@�C����ۑ����Ă��������B" & _
                      "csv�t�@�C���o�͂�xlsm���R�s�[������ɍs���Ă��������B" & _
                      "�R�s�[���܂������H", vbCritical + vbOKCancel, "�x��")
    ConfirmDataSave = (response = vbOK)
End Function

Function GetSheetNumber(prompt As String, outcnt As Long) As Long
    Dim result As Variant
    result = InputBox("�����琔���āA" & prompt & "�V�[�g�͉��Ԗڂł����H" & _
                      "���݂̃A�E�g�J������" & outcnt & "�ł��B" & _
                      "����̃A�E�g�J���������w�肷��ꍇ�́A������������͂��Ă��������B", _
                      "csv�ɂ���͈�", "")
    If StrPtr(result) = 0 Then
        GetSheetNumber = 0
    Else
        GetSheetNumber = CLng(result)
    End If
End Function

Function GetSaveAsFileName(newFileName As String) As Variant
    GetSaveAsFileName = Application.GetSaveAsFileName( _
                        InitialFileName:=newFileName & ".csv", _
                        FileFilter:="CSV�t�@�C��(*.csv),*.csv", _
                        FilterIndex:=1, _
                        Title:="�ۑ��t�@�C���̎w��")
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

    ' �ŏI�s�ƍŏI����擾
    With ActiveSheet
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
        jmax = .Cells(5, .Columns.Count).End(xlToLeft).Column
        
        ' �Ώۗ�̏���
        For j = 2 To jmax
            ' �F�œh��Ԃ���Ă��Ȃ�����\���ɂ���
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
    
    ' ���[�N�V�[�g�̐ݒ�
    Set wsInput = Worksheets("input")
    Set wsLink_List = EnsureWorksheet("�����N�ꗗ", "outcome_type")

    ' URL�̊�{�`
    url = "https://pubmed.ncbi.nlm.nih.gov/"
    
    With wsLink_List
        ' �w�b�_�[�̃R�s�[
        wsInput.Range("B5:G5").Copy .Cells(5, 2)
        wsLink_List.Cells(5, 8).Value = "�����N"
    
        ' �f�[�^�ƃn�C�p�[�����N�̍쐬
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

' ���[�N�V�[�g�����݂��Ȃ��ꍇ�͐V�����쐬����
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
    Set wsLink_List = Worksheets("�����N�ꗗ")
    
    With wsInput
        ' �ő�s�����擾
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    ' �e�s��PubMed�����N���J��
    For i = 6 To imax
        Dim url As String
        url = wsLink_List.Cells(i, 8)
        
        ' Google Chrome���N�����w��URL�Ɉړ�
        CreateObject("WScript.Shell").Run ("chrome.exe -url " & url)
    Next i
End Sub

Sub PMID_OR_Search_Expression()
    Call CheckPMIDDuplicatesOrMissing
    ufSet4.Show vbModeless
    ufSet4.Repaint
    Call ChromeDriverUpdate
    Call �������쐬
End Sub
Sub CheckPMIDDuplicatesOrMissing()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' ���݂̃V�[�g���g�p����

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row ' D��̍ŏI�s���擾

    Dim PMID As Variant
    Dim PMIDDict As Object
    Set PMIDDict = CreateObject("Scripting.Dictionary") ' �����I�u�W�F�N�g�̍쐬

    Dim i As Long
    Dim isProblemFound As Boolean
    isProblemFound = False

    ' D6�Z���ȍ~��PMID���`�F�b�N
    For i = 6 To lastRow
        PMID = ws.Cells(i, 4).Value
        If Not IsEmpty(PMID) Then
            If PMIDDict.Exists(PMID) Then
                ' �d������������
                isProblemFound = True
                Exit For
            Else
                PMIDDict.Add PMID, True
            End If
        End If
    Next i

    ' �G���[���b�Z�[�W�̕\��
    If isProblemFound Then
        MsgBox "PMID�ɏd���������͔������L��܂��B�d���������͔���������ƌ������̍쐬���ł��܂���", vbExclamation, "�G���[����"
    End If
End Sub
'// SeleniumBasic
Public Sub ChromeDriverUpdate()
    Dim Driver As New Selenium.ChromeDriver
    SafeOpen Driver, Chrome
    Driver.Get "https://www.google.co.jp/?q=selenium"
    Driver.Wait 3000
    Driver.Quit
    ' MsgBox ("Chrome Driver�̍X�V����")
End Sub

Sub �������쐬()
    Dim i As Long, j As Long, k As Long, imax As Long
    Dim wsInput As Worksheet, wsLink_List As Worksheet
    Dim Driver As Selenium.WebDriver
    Dim Keys As Keys

    Set wsInput = Worksheets("input")
    Set wsLink_List = Worksheets("�����N�ꗗ")
    Set Driver = New Selenium.WebDriver
    Set Keys = New Keys
    
    With wsInput
        imax = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With

    ' WebDriver�̏����ݒ�ƋN��
    Driver.Start "chrome"
    Driver.Get "https://pubmed.ncbi.nlm.nih.gov/advanced/"

    ' PMID���������Č��ʂ��擾
    For i = 6 To imax
        Dim PMID As String
        PMID = wsLink_List.Cells(i, 4)
        
        Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
        Driver.FindElementByCss("#query-box-input").SendKeys PMID
        Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    Next i
    
            
    Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
            
    ' �������ʂ�g�ݍ��킹
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

    ' �ŏI�I�Ȍ������ʂ��擾
    Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    wsLink_List.Cells(1, 1) = "�S�Ă�PMID��OR�Ō��������ꍇ�̌�����"
    wsLink_List.Cells(2, 1) = Driver.FindElementByCss("#id_term").Value
    
    ' WebDriver�����
    Driver.Quit
    
    wsLink_List.Activate
    Unload ufSet4
    
    MsgBox "�u" & wsLink_List.name & "�v�V�[�g��" & wsLink_List.Cells(2, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "�Z���ɁA�S�Ă�PMID��OR�Ō��������ꍇ�̌��������o�͂��܂����B"
End Sub



