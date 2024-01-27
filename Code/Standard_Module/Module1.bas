Attribute VB_Name = "Module1"
' Module1
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

' This file was output using the VBAC tool; VBAC is bundled with the free VBA library Ariawase.
' Ariawase (and VBAC) are available in the following GitHub repository: https://github.com/vbaidiot/ariawase/tree/master
' The original author is Copyright (c) 2011 igeta.
' Ariawase and VBAC are used under the MIT License, and the full license can be found at the following link: https://github.com/vbaidiot/ariawase/blob/master/LICENSE.txt

Option Explicit
Sub create()
    Application.ScreenUpdating = False
    
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("Before executing this macro, please save all necessary data. When the macro is executed, the original data will be lost. Have you saved your data?", vbCritical + vbOKCancel, "Warning icon")
    If rc = vbCancel Then
        Exit Sub
    End If

    Call outcome_Format
    
    Call leftsheet_delete
    
    Call outcome_sheet
    
    Call continuous_or_dichotomous
    
    Application.ScreenUpdating = True
    
    Call StartEndSheetDelete
    
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    wsInput.Activate
    
    MsgBox "complete"
End Sub

Sub outcome_Format()
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet
    Dim i As Long, k As Long, startCol As Long, lastCol As Long, Colwidth As Long
    Dim outcomeType As String
    
    ' For ContinuousOutcome, the number of columns in the combined cell is 16
    ' For DichotomousOutcome, the number of columns in the combined cell is 12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 16
    DichotomousWide = 12
    
    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")

    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' Find the outcome_format sheet and delete it if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOutcome_Format = Worksheets("outcome_format")
    If Not wsOutcome_Format Is Nothing Then
        wsOutcome_Format.Delete
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create a new sheet and set a name
    Set wsOutcome_Format = Worksheets.Add(After:=Worksheets("Link_List"))
    wsOutcome_Format.name = "outcome_format"

    ' Set the header on the outcome_format sheet
    With wsOutcome_Format
        .Cells(2, 1).Value = "No"
        .Cells(2, 2).Value = "type"
        .Cells(2, 3).Value = "outcome"
    End With
    
    i = 3 ' Start position of data line

    ' Output the outcome entered in the InputSheet
    For k = startCol To lastCol
        Colwidth = wsInput.Cells(3, k).MergeArea.Columns.Count
        
        With wsOutcome_Format
            ' Set sequential numbers in column No.
            .Cells(i, 1) = i - 2
            
            ' Set outcome type in the type column
            Select Case Colwidth
                Case ContinuousWide
                    outcomeType = "Continuous"
                Case DichotomousWide
                    outcomeType = "Dichotomous"
                Case Else
                    outcomeType = ""
            End Select
            .Cells(i, 2) = outcomeType
            
            ' Set the outcome name in the outcome column
            .Cells(i, 3) = wsInput.Cells(3, k).Value
        End With
        
        i = i + 1
        k = k + Colwidth - 1 ' Increment by the amount of merged cells
    Next k
    Call Study_No_Assign
End Sub
Sub FindStrategiesAndSetColumns(wsInput As Worksheet, ByRef startCol As Long, ByRef lastCol As Long)
    Dim foundCell As Range

    ' Locate the column with "Strategies" and get the last column
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
    
    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")
    
    ' Get the last row
    imax = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).row
    
    ' Update "Study No." column
    With wsInput
        For i = 6 To imax
            .Cells(i, 2).FormulaR1C1 = "=SUBTOTAL(3,R6C3:R" & i & "C3)"
        Next i
    End With
End Sub
Sub leftsheet_delete_alert()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("Are you sure you want to delete the sheet to the left of InputSheet?", vbCritical + vbOKCancel, "Warning icon")
    If rc = vbCancel Then
        Exit Sub
    End If
    
    Call leftsheet_delete
End Sub

Sub leftsheet_delete()
    
    Dim ws As Worksheet, Target As String

    ' Sheets to the left of this sheet to be deleted
    Target = "InputSheet"
    
    ' Loop sheets
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        ' Delete a sheet if it is not a "Target" sheet
        If ws.name <> Target Then
            ws.Delete
        Else
            ' Exit the loop when the "Target" sheet appears.
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True
    
End Sub

Sub outcome_sheet()
    
    ' Variable Declaration
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet, ws As Worksheet
    Dim k As Long, i As Long, j As Long
    Dim Colwidth As Long, FirstCol As Long, startCol As Long, lastCol As Long
    Dim maxInputRow As Long, maxOutcome_Format_Row As Long
    
    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")
    Set wsOutcome_Format = Worksheets("outcome_format")
    
    ' Get the last row and last column of the InputSheet
    With wsInput
        maxInputRow = .Cells(.Rows.Count, 2).End(xlUp).row
    End With
    
    ' Get the last row of the outcome_format sheet
    With wsOutcome_Format
        maxOutcome_Format_Row = .Cells(.Rows.Count, 1).End(xlUp).row
    End With
    
    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' Delete existing outcome sheets
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
    
    ' Add new outcome sheets
    For i = 1 To maxOutcome_Format_Row - 2
        Set ws = Worksheets.Add(Before:=Worksheets("TableSheet"))
        With wsInput
            ' Copy Common Rows
            .Range(.Cells(1, 1), .Cells(maxInputRow, FirstCol - 1)).Copy ws.Cells(1, 1)
            ' Copy data for each outcome
            Colwidth = .Cells(3, startCol).MergeArea.Columns.Count
            .Range(.Cells(1, startCol), .Cells(maxInputRow, startCol + Colwidth - 1)).Copy ws.Cells(1, FirstCol)
        End With
        ws.name = wsOutcome_Format.Cells(i + 2, 3).Value ' Naming sheets with outcome names
        startCol = startCol + Colwidth ' Go to the next outcome column.
    Next i
    
End Sub

Sub continuous_or_dichotomous()

    ' Variable Declaration
    Dim j As Long, wide As Long, startCol As Long, lastCol As Long
    Dim name As String
    Dim wsInput As Worksheet
    
    ' For ContinuousOutcome, the number of columns in the combined cell is 16
    ' For DichotomousOutcome, the number of columns in the combined cell is 12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 16
    DichotomousWide = 12

    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")

    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' Iterate through the columns from startCol to lastCol
    For j = startCol To lastCol
        ' Get current column width (number of merged cells) and name
        With wsInput
            wide = .Cells(3, j).MergeArea.Columns.Count
            name = .Cells(3, j).Text
        End With
        
        ' Activate the corresponding worksheet and set references
        Worksheets(name).Activate
        
        ' Call the appropriate process depending on Outcome type (the width of the merged cell)
        If wide = ContinuousWide Then
            Call MetaInsightdataLONG_Continuous
        ElseIf wide = DichotomousWide Then
            Call MetaInsightdataLONG_Dichotomous
        End If
        
        ' Move to next merged cell block
        j = j + wide - 1
    Next j
End Sub
Sub MetaInsightdataLONG_Continuous()
    ' Variable Declaration
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long
    
    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' Set the header of the ContinuousOutcome table
    With wsNew
        .Range("A1:E1").Value = Array("Study", "T", "N", "Mean", "SD")
        ix = 2
    End With
    
    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).row
    End With

    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    ' Transposition of data from horizontal to vertical (converted to Long format)
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 4
                ' Extract basic studyData
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' Ends when there is no more data to process.
    
                ' Create comment string from basic studyData
                Dim commentStr As String
                commentStr = Join(studyData, " ")
    
                ' Write data and add comments to ContinuousOutcome table
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
    
    ' Delete rows containing "NR" or spaces in column 4 using AutoFilter
    Application.DisplayAlerts = False ' Disable Alert Dialog
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=4, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    
    ' Delete rows containing "NR" or spaces in column 5 using AutoFilter
    With wsNew
        .AutoFilterMode = False
        .Range("A1:E1").AutoFilter Field:=5, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' Re-enable the alert dialog
    
    ' Set the second column beginning with cell B2, excluding cell B1, to right-aligned
    With wsNew.Range("B2:E" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).row)
        .HorizontalAlignment = xlRight
    End With
End Sub
Sub MetaInsightdataLONG_Dichotomous()
    ' Variable Declaration
    Dim wsInput As Worksheet, wsActive As Worksheet, wsNew As Worksheet
    Dim i As Long, j As Long, ix As Long
    Dim studyData As Variant
    Dim startCol As Long, lastCol As Long, maxActiveRow As Long

    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")
    Set wsActive = ActiveSheet
    Set wsNew = AddOrGetSheet(wsActive.name & " table", wsInput)

    ' Set the header of the DichotomousOutcome table
    With wsNew
        .Range("A1:D1").Value = Array("Study", "T", "R", "N")
        ix = 2
    End With


    With wsActive
        maxActiveRow = .Cells(.Rows.Count, 2).End(xlUp).row
    End With
    
    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsActive, startCol, lastCol)

    ' Transposition of data from horizontal to vertical (converted to Long format)
    With wsActive
        For i = 6 To maxActiveRow
            For j = startCol To lastCol Step 3
                ' Extract basic studyData
                studyData = Array(.Cells(i, 2).Value, .Cells(i, 3).Value, .Cells(i, 4).Value, .Cells(i, 5).Value)
                If .Cells(i, j).Value = "" Then Exit For  ' Ends when there is no more data to process.

                ' Create comment string from basic studyData
                Dim commentStr As String
                commentStr = Join(studyData, " ")

                ' Write data and add comments to ContinuousOutcome table
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
    
    ' Delete rows containing "NR" or spaces in column 3 using AutoFilter
    Application.DisplayAlerts = False ' Disable Alert Dialog
    With wsNew
        .AutoFilterMode = False
        .Range("A1:D1").AutoFilter Field:=3, Criteria1:="NR", Operator:=xlOr, Criteria2:=""
        .AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Delete
        .AutoFilterMode = False
    End With
    Application.DisplayAlerts = True ' Re-enable the alert dialog
    
    ' Set the second column beginning with cell B2, excluding cell B1, to right-aligned
    With wsNew.Range("B2:B" & wsNew.Cells(wsNew.Rows.Count, "B").End(xlUp).row)
        .HorizontalAlignment = xlRight
    End With
End Sub

' Helper function to add a new sheet or retrieve an existing one
Function AddOrGetSheet(name As String, Optional BeforeSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next ' If an error occurs, proceed to the next line of code.
    Set ws = Worksheets(name)
    On Error GoTo 0 ' Return error handling to standard
    If ws Is Nothing Then ' If the sheet does not exist, add a new sheet
        Set ws = Worksheets.Add(Before:=BeforeSheet)
        ws.name = name
    End If
    Set AddOrGetSheet = ws ' return a sheet
End Function
Sub StartEndSheetDelete()
    Const START_SHEET_NAME As String = "InputSheet"  ' Delete start sheet
    Const END_SHEET_NAME As String = "TableSheet"    ' Delete end sheet

    Dim startIndex As Long, endIndex As Long, i As Long, temp As Long

    ' Get sheet index
    startIndex = Sheets(START_SHEET_NAME).Index
    endIndex = Sheets(END_SHEET_NAME).Index

    ' Set start and end indexes appropriately
    If startIndex > endIndex Then
        temp = startIndex
        startIndex = endIndex
        endIndex = temp
    End If

    ' Delete sheet
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
        outcnt = .Cells(.Rows.Count, 1).End(xlUp).row - 2
    End With

    ans1 = GetSheetNumber("start", outcnt)
    If ans1 = 0 Then Exit Sub
    
    ans2 = GetSheetNumber("end", outcnt)
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
    response = MsgBox("Executing this macro will result in the loss of data in the xlsm file." & _
                      " Please save the xlsm file and copy the file before creating the csv file." & _
                      " The csv file output should be done after copying the xlsm file." & _
                      " Have you copied it?", vbCritical + vbOKCancel, "Warning icon")
    ConfirmDataSave = (response = vbOK)
End Function

Function GetSheetNumber(prompt As String, outcnt As Long) As Long
    Dim result As Variant
    result = InputBox("Counting from the left, what is the number of sheet " & prompt & "? " & _
                      "The current number of outcomes is " & outcnt & _
                      ". To specifiy only a specific outcome, enter the same number.", _
                      "Scope to csv", "")
    If StrPtr(result) = 0 Then
        GetSheetNumber = 0
    Else
        GetSheetNumber = CLng(result)
    End If
End Function

Function GetSaveAsFileName(newFileName As String) As Variant
    GetSaveAsFileName = Application.GetSaveAsFileName( _
                        InitialFileName:=newFileName & ".csv", _
                        FileFilter:="CSV Files (*.csv),*.csv", _
                        FilterIndex:=1, _
                        Title:="Specify the file to save")
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

    ' Get the last row and column
    With ActiveSheet
        imax = .Cells(.Rows.Count, 2).End(xlUp).row
        jmax = .Cells(5, .Columns.Count).End(xlToLeft).Column
        
        ' Processing of target columns
        For j = 2 To jmax
            ' Hide columns that are not filled with a color
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
    
    ' Sheet Settings
    Set wsInput = Worksheets("InputSheet")
    Set wsLink_List = EnsureWorksheet("Link_List", "outcome_type")

    ' PubMed URL
    url = "https://pubmed.ncbi.nlm.nih.gov/"
    
    With wsLink_List
        ' Header Copy
        wsInput.Range("B5:G5").Copy .Cells(5, 2)
        wsLink_List.Cells(5, 8).Value = "Link"
    
        ' Create data and hyperlinks
        imax = wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).row
    
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

' If worksheet does not exist, create a new one
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
    
    Set wsInput = Worksheets("InputSheet")
    Set wsLink_List = Worksheets("Link_List")
    
    With wsInput
        ' Get maximum number of rows
        imax = .Cells(.Rows.Count, 2).End(xlUp).row
    End With
    
    ' Open PubMed links for each row
    For i = 6 To imax
        Dim url As String
        url = wsLink_List.Cells(i, 8)
        
        ' Run Google Chrome and navigate to the specified URL.
        CreateObject("WScript.Shell").Run ("chrome.exe -url " & url)
    Next i
End Sub

Sub PMID_OR_Search_Expression()
    Call CheckPMIDDuplicatesOrMissing
    ufSet4.Show vbModeless
    ufSet4.Repaint
    Call ChromeDriverUpdate
    Call create_search_expression
End Sub
Sub CheckPMIDDuplicatesOrMissing()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Use current sheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).row ' Get last row of column D

    Dim PMID As Variant
    Dim PMIDDict As Object
    Set PMIDDict = CreateObject("Scripting.Dictionary") ' Creating Dictionary Objects

    Dim i As Long
    Dim isProblemFound As Boolean
    isProblemFound = False

    ' Check PMID after cell D6
    For i = 6 To lastRow
        PMID = ws.Cells(i, 4).Value
        ' Determine if there are duplicates
        If Not IsEmpty(PMID) Then
            If PMIDDict.Exists(PMID) Then
                isProblemFound = True
                Exit For
            Else
                PMIDDict.Add PMID, True
            End If
        End If
    Next i

    ' Error Messages
    If isProblemFound Then
        MsgBox "There is a duplicate or missing PMID. If there are duplicates or omissions, you will not be able to create a search expression.", vbExclamation, "There are duplicate or missing PMIDs."
    End If
End Sub


Sub create_search_expression()
    Dim i As Long, j As Long, k As Long, imax As Long
    Dim wsInput As Worksheet, wsLink_List As Worksheet
    Dim Driver As Selenium.WebDriver
    Dim Keys As Keys

    Set wsInput = Worksheets("InputSheet")
    Set wsLink_List = Worksheets("Link_List")
    Set Driver = New Selenium.WebDriver
    Set Keys = New Keys
    
    With wsInput
        imax = .Cells(.Rows.Count, 2).End(xlUp).row
    End With

    ' WebDriver initialization and startup
    Driver.Start "chrome"
    Driver.Get "https://pubmed.ncbi.nlm.nih.gov/advanced/"

    ' Search for PMID and get results
    For i = 6 To imax
        Dim PMID As String
        PMID = wsLink_List.Cells(i, 4)
        
        Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
        Driver.FindElementByCss("#query-box-input").SendKeys PMID
        Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    Next i
    
            
    Driver.Get ("https://pubmed.ncbi.nlm.nih.gov/advanced/")
            
    ' Combine search queries with OR
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

    ' Get final search result
    Driver.FindElementByCss("#search-form > div > div > div.query-box-section-wrapper > div.button-wrapper > button > span").Click
    wsLink_List.Cells(1, 1) = "Search expression when all PMIDs joined by OR"
    wsLink_List.Cells(2, 1) = Driver.FindElementByCss("#id_term").Value
    
    ' Close WebDriver
    Driver.Quit
    
    wsLink_List.Activate
    Unload ufSet4
    
    MsgBox "Output of search formula for all PMIDs joined by OR in cell " & wsLink_List.Cells(2, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & " of " & wsLink_List.name & " Sheet"
End Sub



