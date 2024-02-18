VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSet 
   Caption         =   "Input Form"
   ClientHeight    =   7370
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11030
   OleObjectBlob   =   "ufSet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ufSet
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Private ButtonHandler As New Collection
' "Back" button action
Sub Back_method()
    Dim wsInput As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    wsInput.Activate
    
    If ActiveCell.row > 6 Then
        ActiveCell.Offset(-1, 0).Activate
        wsInput.Cells(ActiveCell.row, 2).Activate
        Call updateForm
    End If
    
End Sub
' "Next" button action
Sub Next_method()
    Dim wsInput As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    wsInput.Activate
    
    With wsInput
        Dim LastColoredRow As Long
        
        ' Get the last row with color after the last row with text
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check column 2
        
        If ActiveCell.row <= LastColoredRow - 1 Then
            ActiveCell.Offset(1, 0).Activate
            .Cells(ActiveCell.row, 2).Activate
        End If
    End With
    Call updateForm
End Sub
' "Add" button action
Sub Add_method()
    Dim wsInput As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    Application.ScreenUpdating = False
    
    
    Call updateForm
    
    With wsInput
        Dim maxInputRow As Long
        Dim LastColoredRow As Long
        maxInputRow = .Cells(.Rows.Count, 2).End(xlUp).row
        
        
        ' Get the last row with color after the last row with text
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check column 2
    
        ' Only add a new row if the last row with input and the last colored row are the same
        If maxInputRow = LastColoredRow Then
            Call InputRowInsert
            Call updateForm
        End If
    End With
    
    wsInput.Activate
    wsInput.Cells(wsInput.Rows.Count, 2).End(xlUp).Offset(1, 0).Activate
End Sub
' "Outcome_Add" button action
Sub Outcome_Add_method()
    Unload ufSet
    ufSet2.Show
End Sub
' "RowDelete" button action
Sub RowDelete_method()
    Dim wsInput As Worksheet
    Dim LastColoredRow As Long, DeletedLastColoredRow As Long
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    wsInput.Activate
    
    
    ' Get the last row with color after the last row with text
    LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check column 2
    
    
    If 6 <= ActiveCell.row Then
        ' Determine if the number of rows in the active cell is within the last row that is colored and delete rows only within the rows of the table
        If LastColoredRow >= ActiveCell.row Then
            
            ' Warning before deletion
            Dim rc As VbMsgBoxResult
            rc = MsgBox("Delete the currently selected row. Are you sure you want to delete?", vbCritical + vbOKCancel, "Warning icon")
            If rc = vbCancel Then
                Exit Sub
            End If
            
            ActiveCell.EntireRow.Delete
            ' Deleting a row in the table shifts the last row that is colored up by one.
            ' After deleting a row, get the last row that is colored again with DeletedLastColoredRow.
            DeletedLastColoredRow = LastColoredRow - 1
            
            ' If the active cell is now outside the table because you deleted a row (when the active cell was originally in the last row with a color)
            ' Correct by activating one cell above
            If DeletedLastColoredRow < ActiveCell.row Then
                ActiveCell.Offset(-1, 0).Activate
            End If
        Else
            MsgBox ("The Delete button cannot delete rows in cells outside the table. It can only delete rows inside the table.")
        End If
    End If
    Call updateForm
End Sub
' "Update" button action
Sub Update_method()
    ' Sheet Settings
    Dim wsInput As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    
    ' Warn before updating
    Dim rc As VbMsgBoxResult
    rc = MsgBox("Update the data of the currently selected row. Are you sure?", vbCritical + vbOKCancel, "Warning icon")
    If rc = vbCancel Then
        Exit Sub
    End If
    
    Dim outcomeIndex As Long
    Dim wide As Long
    Dim k As Long
    Dim startCol As Long
    Dim lastCol As Long
    
    ' For ContinuousOutcome, the number of columns in the combined cell is 16
    ' For DichotomousOutcome, the number of columns in the combined cell is 12
    Dim ContinuousWide As Long
    Dim DichotomousWide As Long
    ContinuousWide = 16
    DichotomousWide = 12

    ' Set value in active cell
    UpdateGeneralInfo wsInput
    UpdateStrategies wsInput

    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' Set data for each outcome
    outcomeIndex = 1
    For k = startCol To lastCol
        wide = wsInput.Cells(3, k).MergeArea.Columns.Count
        If wide = ContinuousWide Then ' If the outcome is Continuous
            UpdateContinuousOutcome wsInput, k, outcomeIndex
        ElseIf wide = DichotomousWide Then ' If the outcome is Dichotomous
            UpdateDichotomousOutcome wsInput, k, outcomeIndex
        End If
        outcomeIndex = outcomeIndex + 1
        k = k + wide - 1
    Next k
    Call Study_No_Assign
    Call updateForm
End Sub

' Subprocedure to update basic information
Private Sub UpdateGeneralInfo(ws As Worksheet)
    Dim j As Long
    Dim controlName As String
    With ws
        For j = 2 To 7
            controlName = "txt" & .Cells(4, j).Value
            .Cells(ActiveCell.row, j).Value = Me.Controls(controlName).Text
        Next j
    End With
End Sub
' Subprocedure to update Strategies
Private Sub UpdateStrategies(ws As Worksheet)
    Dim o As Long, r As Long
    r = 8 ' initial column position
    Dim TableSheet As Worksheet
    Set TableSheet = Worksheets("TableSheet")

    With ws
        Dim activeRow As Long
        activeRow = ActiveCell.row

        For o = 1 To 4
            .Cells(activeRow, r).Value = Me.Controls("txt" & TableSheet.Cells(3, 3).Value & "tr" & o).Text
            .Cells(activeRow, r + 1).Value = Me.Controls("txt" & TableSheet.Cells(5, 1).Value & "tr" & o).Text
            .Cells(activeRow, r + 2).Value = Me.Controls("txtntr" & o).Text
            
            ' Processing if checkbox is selected
            If Controls("chk" & .Cells(5, 27 + o).Value).Value = True Then
                Dim presentrow As Long
                 presentrow = Me.Controls("cmb" & .Cells(5, 27 + o).Value).ListIndex + 3
                .Cells(activeRow, 27 + o).FormulaR1C1 = _
                "=SortSheet!R" & presentrow & "C16"
                .Cells(activeRow, 27 + o).Interior.Color = rgbPaleTurquoise
            End If
            ' Deselect checkbox
            Controls("chk" & .Cells(5, 27 + o).Value).Value = False
            
            r = r + 3
        Next o
    End With
    
End Sub
' Subprocedure to update Continuous Outcome
Private Sub UpdateContinuousOutcome(ws As Worksheet, colIndex As Long, pageIndex As Long)
    Dim m As Long, p As Long
    p = colIndex
    For m = 1 To 4
        With ws
            .Cells(ActiveCell.row, p + 1) = Controls("txtout" & pageIndex & "meantr" & m).Text
            .Cells(ActiveCell.row, p + 2) = Controls("txtout" & pageIndex & "sdtr" & m).Text
            ' Processing if checkbox is selected
            If Controls("chkout" & pageIndex & "sub1tr" & m).Value = True Then
                .Cells(ActiveCell.row, p + 3) = Controls("txtout" & pageIndex & "sub1ntr" & m).Text
                .Cells(ActiveCell.row, p + 3).Interior.Color = rgbPaleTurquoise
            End If
            ' Deselect checkbox
            Controls("chkout" & pageIndex & "sub1tr" & m).Value = False
        End With
        p = p + 4
    Next m
End Sub

' Subprocedure to update Dichotmous Outcome
Private Sub UpdateDichotomousOutcome(ws As Worksheet, colIndex As Long, pageIndex As Long)
    Dim n As Long, q As Long
    q = colIndex
    For n = 1 To 4
        With ws
            .Cells(ActiveCell.row, q + 1) = Controls("txtout" & pageIndex & "eventtr" & n).Text
            ' Processing if checkbox is selected
            If Controls("chkout" & pageIndex & "sub2tr" & n).Value = True Then
                .Cells(ActiveCell.row, q + 2) = Controls("txtout" & pageIndex & "sub2ntr" & n).Text
                .Cells(ActiveCell.row, q + 2).Interior.Color = rgbPaleTurquoise
            End If
            ' Deselect checkbox
            Controls("chkout" & pageIndex & "sub2tr" & n).Value = False
        End With
        q = q + 3
    Next n
End Sub

' Works to switch between multi-page page tabs
Private Sub MultiPage1_Change()
    Dim wsInput As Worksheet, wsOutcome_Format As Worksheet
    Dim startCol As Long, lastCol As Long, wide As Long, k As Long
    Dim outcomeIndex As Long, outcomeCount As Long, outcomeName As String, captionName As String, extractName_pre As String, extractName As String
    Dim pageIndex As Long
    Dim imax As Long

    Set wsInput = Worksheets("InputSheet")
    Set wsOutcome_Format = Worksheets("outcome_format")


    ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)

    With wsOutcome_Format
        imax = .Cells(.Rows.Count, 1).End(xlUp).row
        outcomeCount = imax - 2
        outcomeIndex = 1
    End With

    ' Program to scroll in sync with MultiPage1 page tab selection
    pageIndex = ufSet.MultiPage1.Value
    captionName = MultiPage1.Pages.Item(pageIndex).caption
    
    ' MultiPage1 page tab name ("PF ratio" of "outcome1(PF ratio)" is extracted)
    extractName_pre = Mid(captionName, InStr(captionName, "(") + 1)
    extractName = left(extractName_pre, Len(extractName_pre) - 1)
    
    
    With wsInput
        ' Scroll in sync with outcome tab selection
        If pageIndex >= 2 Then
            For k = startCol To lastCol
                wide = .Cells(3, k).MergeArea.Columns.Count
                outcomeName = wsOutcome_Format.Cells(outcomeIndex + 2, 3)
    
                If extractName = outcomeName Then
                    ActiveWindow.ScrollColumn = .Cells(3, k).Column
                        Exit For
                End If
    
                outcomeIndex = outcomeIndex + 1
                k = k + wide - 1
            Next k
        ' Scroll in sync with the selection of the Information and Strategies tabs
        ElseIf pageIndex = 0 Or pageIndex = 1 Then
            ActiveWindow.ScrollColumn = .Cells(3, 1).Column
        End If
    End With
End Sub

' initialization
Private Sub UserForm_Initialize()
    
    ' Start of outcome_format
    Call outcome_Format
    
    ' Sheet Settings
    Dim wsInput As Worksheet, wsOutcome_Type As Worksheet, wsOutcome_Format As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    Set wsOutcome_Type = ThisWorkbook.Worksheets("outcome_type")
    Set wsOutcome_Format = ThisWorkbook.Worksheets("outcome_format")

    ' Various controls are placed on the first page of the multipage (Information)
    MultiPage1.Pages(0).caption = wsInput.Cells(3, 2).Value
    Call PageInformation_AddLabelsToPage(0)
    Call PageInformation_AddTextBoxesToPage(0)
    Call PageInformation_AddButtonsToPage(0)
    
    ' Set properties of each control on the first page of multi-page (Information)
    Call PageInformation_SetControlsProperties(0)
    
    ' Add Strategies page and place various controls
    MultiPage1.Pages.Add ("Strategies")
    Call Strategies_AddLabelsToPage(0)
    Call Strategies_AddTextBoxesToPage(0)
    Call Strategies_AddComboBoxesToPage(0)
    Call Strategies_AddCheckBoxesToPage(0)
    
    ' Set various control properties on the Strategies page
    Call Strategies_SetControlsProperties(0)

    ' Add pages to the multipage for each outcome and place various controls
    Dim i As Long, imax As Long
    With wsOutcome_Format
        imax = .Cells(.Rows.Count, 1).End(xlUp).row - 2

        For i = 1 To imax
            ' Add page
            MultiPage1.Pages.Add ("outcome" & i & " (" & .Cells(i + 2, 3) & ")")
            
            ' Placement of various controls on various outcome pages
            Call Outcomes_AddLabelsToPage(i)
            Call Outcomes_AddTextBoxesToPage(i)
            Call Outcomes_AddComboBoxesToPage(i)
            Call Outcomes_AddCheckBoxesToPage(i)
    
            ' Set properties of various controls on various outcome pages
            Call Outcomes_SetControlsProperties(i)
    
            ' Show/hide controls based on specific criteria
            Call Outcomes_ToggleControlsVisibility(i, .Cells(i + 2, 2).Value)
        Next i
    End With

    ' Value setting
    Call updateForm
End Sub
Sub PageInformation_AddLabelsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newLabel As Control
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        controlNames = Array("title", _
                             .Cells(4, 2).Value, .Cells(4, 3).Value, .Cells(4, 4).Value, _
                             .Cells(4, 5).Value, .Cells(4, 6).Value, .Cells(4, 7).Value) ' Create list of control names
        
    End With

    ' Add labels beginning with lbl
    For i = LBound(controlNames) To UBound(controlNames)
        Set newLabel = MultiPage1.Pages(pageIndex).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
    Next i
End Sub
Sub PageInformation_AddTextBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newText As Control
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        controlNames = Array(.Cells(4, 2).Value, .Cells(4, 3).Value, .Cells(4, 4).Value, _
                             .Cells(4, 5).Value, .Cells(4, 6).Value, .Cells(4, 7).Value) ' Create list of control names
        
    End With
    
    ' Add textboxes beginning with txt
    For i = LBound(controlNames) To UBound(controlNames)
        Set newText = MultiPage1.Pages(pageIndex).Controls.Add("Forms.TextBox.1", "txt" & controlNames(i))
        ' You can set other properties of newText here
    Next i
End Sub

Sub PageInformation_AddButtonsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newButton As Control
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        controlNames = Array("Back", "Next", "Add", "Outcome_Add", "RowDelete", _
                             "Update") ' Create list of control names
        
    End With
    
   ' Add CommandButton beginning with btn
    For i = LBound(controlNames) To UBound(controlNames)
        Set newButton = MultiPage1.Pages(pageIndex).Controls.Add("Forms.CommandButton.1", "btn" & controlNames(i))
        
        Dim buttonEvent As New Class1
        
        ' Event Handler Assignment
        With buttonEvent
            If i = 0 Then
                Set .btnBack = newButton
            ElseIf i = 1 Then
                Set .btnNext = newButton
            ElseIf i = 2 Then
                Set .btnAdd = newButton
            ElseIf i = 3 Then
                Set .btnOutcome_Add = newButton
            ElseIf i = 4 Then
                Set .btnRowDelete = newButton
            ElseIf i = 5 Then
                Set .btnUpdate = newButton
            End If
        End With
        
        ButtonHandler.Add buttonEvent
    Next i
End Sub
' Subroutine to set properties of each control
Private Sub PageInformation_SetControlsProperties(pageIndex As Long)

    Dim i As Long
    Dim wsInput As Worksheet
    Dim maincaption As String, controlName As String
    Dim btncontrolNames As Variant, btncaptionNames As Variant
    
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        ' Label property setting (lbl)
        ' title label
        
        maincaption = MultiPage1.Pages(0).caption
        controlName = "lbl" & "title"
        
        With Me.Controls(controlName)
              .caption = maincaption ' caption
              .Font.Size = 9 ' Size 9
              .Font.bold = True ' Bold
              .left = 10 ' Distance from the left
              .top = 10 ' Distance from the top
              .width = 90
              .height = 10
        End With
            
        For i = 2 To 7
            ' StudyNo label, Authors label, PMID label, Year label, Country label, ResearchPeriod label
            controlName = "lbl" & .Cells(4, i).Value
            With Me.Controls(controlName)
                .caption = wsInput.Cells(4, i) ' caption
                .Font.Size = 9 ' Size 9
                .left = 10  ' Distance from the left
                .top = 25 + (i - 2) * 30 ' Distance from the top
                .width = 90
                .height = 10
            End With
        Next i
        ' Set other labels in the same way.
        
    
        ' Textbox property setting (txt)
        For i = 2 To 7
            ' No textbox, Authors textbox, PMID textbox, Year textbox, Country textbox, ResearchPeriod textbox
            controlName = "txt" & .Cells(4, i).Value
            With Me.Controls(controlName)
                .left = 10
                .top = 35 + (i - 2) * 30
                .width = 90
                .height = 15
            End With
        Next i
        ' Set other textboxes in the same way.
        
        ' commandbutton property setting (btn)
        btncontrolNames = Array("Back", "Next", "Add", "Outcome_Add", "RowDelete", _
                                "Update") ' Create list of control names
                             
        btncaptionNames = Array("Back", "Next", "Add", "Outcome_Add", "RowDelete", _
                                "Update") ' Create list of caption names
        
        For i = 1 To 5
            ' Backcommandbutton, Nextcommandbutton, Addcommandbutton, Outcome_Addcommandbutton, RowDeletecommandbutton
            controlName = "btn" & btncontrolNames(i - 1)
            With Me.Controls(controlName)
                .caption = btncaptionNames(i - 1)
                .left = 110 + (i - 1) * 65
                .top = 10
                .width = 60
                .height = 20
            End With
        Next i
        
        i = 6
        ' Updatecommandbutton
        controlName = "btn" & btncontrolNames(i - 1)
        With Me.Controls(controlName)
            .caption = btncaptionNames(i - 1)
            .left = 10
            .top = 205
            .width = 90
            .height = 20
        End With
        
        ' Set other commandbuttons in the same way.
    End With

End Sub
Sub Strategies_AddLabelsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newLabel As Control
    Dim TableSheet As Worksheet, wsInput As Worksheet
    Set TableSheet = Worksheets("TableSheet")
    Set wsInput = Worksheets("InputSheet")
    
    With TableSheet
        ' Create list of control names
        controlNames = Array("maintitle", "tr1", "tr2", "tr3", "tr4", _
                             .Cells(3, 3).Value & "tr1", .Cells(3, 3).Value & "tr2", .Cells(3, 3).Value & "tr3", .Cells(3, 3).Value & "tr4", _
                             .Cells(5, 1).Value & "tr1", .Cells(5, 1).Value & "tr2", .Cells(5, 1).Value & "tr3", .Cells(5, 1).Value & "tr4", _
                             wsInput.Cells(5, 28).Value, wsInput.Cells(5, 29).Value, wsInput.Cells(5, 30).Value, wsInput.Cells(5, 31).Value, _
                             "ntr1", "ntr2", "ntr3", "ntr4")
    End With

    ' Add labels beginning with lbl
    For i = LBound(controlNames) To UBound(controlNames)
        Set newLabel = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
    Next i
End Sub
Sub Strategies_AddTextBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newText As Control
    Dim TableSheet As Worksheet, wsInput As Worksheet
    Set TableSheet = Worksheets("TableSheet")
    Set wsInput = Worksheets("InputSheet")
    
    With TableSheet
        ' Create list of control names
        controlNames = Array(.Cells(3, 3).Value & "tr1", .Cells(3, 3).Value & "tr2", .Cells(3, 3).Value & "tr3", .Cells(3, 3).Value & "tr4", _
                             .Cells(5, 1).Value & "tr1", .Cells(5, 1).Value & "tr2", .Cells(5, 1).Value & "tr3", .Cells(5, 1).Value & "tr4", _
                             "ntr1", "ntr2", "ntr3", "ntr4")
    End With
    
    ' Add labels beginning with txt
    For i = LBound(controlNames) To UBound(controlNames)
        Set newText = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.TextBox.1", "txt" & controlNames(i))
        ' You can set other properties of newText here
    Next i
End Sub
Sub Strategies_AddComboBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newCombo As Control
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        ' Create list of control names
        controlNames = Array(.Cells(5, 28).Value, .Cells(5, 29).Value, .Cells(5, 30).Value, .Cells(5, 31).Value)
    End With
    
    ' Add ComboBoxes beginning with Cmb
    For i = LBound(controlNames) To UBound(controlNames)
        Set newCombo = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.ComboBox.1", "cmb" & controlNames(i))
        ' You can set other properties of newCombo here
    Next i
End Sub
Sub Strategies_AddCheckBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newChk As Control
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        ' Create list of control names
        controlNames = Array(.Cells(5, 28).Value, .Cells(5, 29).Value, .Cells(5, 30).Value, .Cells(5, 31).Value)
    End With
    
    ' Add checkboxes beginning with chk
    For i = LBound(controlNames) To UBound(controlNames)
        Set newChk = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.CheckBox.1", "chk" & controlNames(i))
        ' You can set other properties of newChk here
    Next i
End Sub

' Subroutine to set properties of each control
Private Sub Strategies_SetControlsProperties(pageIndex As Long)
    Dim i As Long
    Dim TableSheet As Worksheet, wsInput As Worksheet, SortSheet As Worksheet
    Dim controlName As String
    
    Set TableSheet = Worksheets("TableSheet")
    Set wsInput = Worksheets("InputSheet")
    Set SortSheet = Worksheets("SortSheet")
    
    With TableSheet
        For i = 1 To 4
            ' Label property setting (lbl)
            ' maintitle label
            controlName = "lbl" & "maintitle"
            With Me.Controls(controlName)
               .caption = "Strategies" ' caption
               .Font.Size = 9 ' Size 9
               .Font.bold = True ' Bold
               .left = 10 ' Distance from the left
               .top = 10 ' Distance from the top
               .width = 90
               .height = 10
            End With
            
            ' treatment label
            controlName = "lbl" & "tr" & i
            With Me.Controls(controlName)
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 25 ' Distance from the top
                .width = 90
                .height = 10
            End With
            
            ' IndexA label
            controlName = "lbl" & .Cells(3, 3).Value & "tr" & i
            With Me.Controls(controlName)
                    .caption = TableSheet.Cells(3, 3).Value & i ' caption
                    .Font.Size = 9 ' Size 9
                    .left = 10 + (i - 1) * 100 ' Distance from the left
                    .top = 40 ' Distance from the top
                    .width = 90
                    .height = 10
            End With
    
            ' IndexB label
            controlName = "lbl" & .Cells(5, 1).Value & "tr" & i
            With Me.Controls(controlName)
                    .caption = TableSheet.Cells(5, 1).Value & i ' caption
                    .Font.Size = 9 ' Size 9
                    .left = 10 + (i - 1) * 100 ' Distance from the left
                    .top = 70 ' Distance from the top
                    .width = 90
                    .height = 10
            End With
    
            ' n label
            controlName = "lbl" & "ntr" & i
            With Me.Controls(controlName)
                    .caption = "Patients (n" & i & ")" ' caption
                    .Font.Size = 9 ' Size 9
                    .left = 10 + (i - 1) * 100 ' Distance from the left
                    .top = 100 ' Distance from the top
                    .width = 90
                    .height = 10
            End With
            
             ' arm label
            controlName = "lbl" & left(wsInput.Cells(5, 27 + i).Value, Len(wsInput.Cells(5, 27 + i).Value) - 1) & i
            With Me.Controls(controlName)
                    .caption = wsInput.Cells(5, 27 + i).Value ' caption
                    .Font.Size = 9 ' Size 9
                    .left = 10 + (i - 1) * 100 ' Distance from the left
                    .top = 130 ' Distance from the top
                    .width = 90
                    .height = 10
            End With
            
        Next i
        ' Set other labels in the same way.
        
    
        ' Textboxes property setting (txt)
        For i = 1 To 4
            'IndexA textbox
            controlName = "txt" & .Cells(3, 3).Value & "tr" & i
            With Me.Controls(controlName)
                .left = 10 + (i - 1) * 100
                .top = 50
                .width = 90
                .height = 15
                .IMEMode = 3 ' Set IME mode
                .tabIndex = 1 + (i - 1) * 4
            End With
            
            ' IndexB textbox
            controlName = "txt" & .Cells(5, 1).Value & "tr" & i
            With Me.Controls(controlName)
                .left = 10 + (i - 1) * 100
                .top = 80
                .width = 90
                .height = 15
                .IMEMode = 3
                .tabIndex = 2 + (i - 1) * 4
            End With
            
            ' n textbox
            controlName = "txt" & "ntr" & i
            With Me.Controls(controlName)
                .left = 10 + (i - 1) * 100
                .top = 110
                .width = 90
                .height = 15
                .IMEMode = 3
                .tabIndex = 3 + (i - 1) * 4
            End With
        Next i
        ' Set other textboxes in the same way.
        
        ' ComboBoxes property setting (cmb)
        Dim imax As Long
        Dim ary_arm As Variant
        With SortSheet
            imax = .Cells(.Rows.Count, 16).End(xlUp).row
            ary_arm = .Range(.Cells(3, 16), .Cells(imax, 16)).Value
        End With
        
        For i = 1 To 4
            'cmb checkbox
            controlName = "cmb" & left(wsInput.Cells(5, 27 + i).Value, Len(wsInput.Cells(5, 27 + i).Value) - 1) & i
            With Me.Controls(controlName)
                .left = 10 + (i - 1) * 100
                .top = 140
                .width = 90
                .height = 15
                .tabIndex = 4 + (i - 1) * 4
                .List = Application.Transpose(ary_arm)
            End With
        Next i
        ' Set other checkboxes in the same way.
        
        ' CheckBoxes property setting (chk)
        For i = 1 To 4
            'chk checkbox
            controlName = "chk" & left(wsInput.Cells(5, 27 + i).Value, Len(wsInput.Cells(5, 27 + i).Value) - 1) & i
            With Me.Controls(controlName)
                .caption = "Enter the name of arm" & i & " directly." ' caption
                .left = 410
                .top = 50 + (i - 1) * 30
                .width = 140
                .height = 15
            End With
        Next i
    ' Set other checkboxes in the same way.
    End With
End Sub

Sub Outcomes_AddLabelsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newLabel As Control

    ' Create list of control names
    controlNames = Array("maintitle", "subtitle1", "sub1tr1", "sub1tr2", "sub1tr3", "sub1tr4", _
                         "meantr1", "meantr2", "meantr3", "meantr4", "sdtr1", "sdtr2", "sdtr3", "sdtr4", _
                         "sub1ntr1", "sub1ntr2", "sub1ntr3", "sub1ntr4", _
                         "subtitle2", "sub2tr1", "sub2tr2", "sub2tr3", "sub2tr4", _
                         "eventtr1", "eventtr2", "eventtr3", "eventtr4", _
                         "sub2ntr1", "sub2ntr2", "sub2ntr3", "sub2ntr4")

    ' Add labels beginning with lblout
    For i = LBound(controlNames) To UBound(controlNames)
        Set newLabel = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.Label.1", "lblout" & pageIndex & controlNames(i))
        ' You can set other properties of newLabel here
    Next i

    ' lblcmb and lbltype have different name patterns, so add them separately
    Set newLabel = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.Label.1", "lblcmb" & pageIndex)
    Set newLabel = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.Label.1", "lbltype" & pageIndex)
End Sub
Sub Outcomes_AddTextBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newText As Control

    ' Create list of control names
    controlNames = Array("meantr1", "meantr2", "meantr3", "meantr4", _
                         "sdtr1", "sdtr2", "sdtr3", "sdtr4", _
                         "sub1ntr1", "sub1ntr2", "sub1ntr3", "sub1ntr4", _
                         "eventtr1", "eventtr2", "eventtr3", "eventtr4", _
                         "sub2ntr1", "sub2ntr2", "sub2ntr3", "sub2ntr4")

    ' Add textboxes beginning with txtout
    For i = LBound(controlNames) To UBound(controlNames)
        Set newText = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.TextBox.1", "txtout" & pageIndex & controlNames(i))
        ' You can set other properties of newText here
    Next i
End Sub
Sub Outcomes_AddComboBoxesToPage(pageIndex As Long)
    Dim controlPrefixes As Variant, prefix As Variant
    Dim newCombo As Control

    ' Create list of control names
    controlPrefixes = Array("", "type")

    ' Add comboboxes beginning with cmb
    For Each prefix In controlPrefixes
        Set newCombo = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.ComboBox.1", "cmb" & prefix & pageIndex)
        ' You can set other properties of newCombo here
    Next prefix
End Sub
Sub Outcomes_AddCheckBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant, name As Variant
    Dim newChk As Control

    ' Create list of control names
    controlNames = Array("sub1tr1", "sub1tr2", "sub1tr3", "sub1tr4", "sub2tr1", "sub2tr2", "sub2tr3", "sub2tr4")

    ' Add checkboxes beginning with chk
    For Each name In controlNames
        Set newChk = MultiPage1.Pages(pageIndex + 1).Controls.Add("Forms.CheckBox.1", "chkout" & pageIndex & name)
       ' You can set other properties of newChk here
    Next name
End Sub
' Subroutine to set properties of each control
Private Sub Outcomes_SetControlsProperties(pageIndex As Long)
    Dim i As Long
    Dim wsOutcome_Format As Worksheet
    Dim controlName As String
    
    Set wsOutcome_Format = Worksheets("outcome_format")
    
    For i = 1 To 4
        ' Label property setting (lbl)
        ' maintitle label
        controlName = "lblout" & pageIndex & "maintitle"
        With Me.Controls(controlName)
           .caption = "Outcome" & (pageIndex) & " (" & wsOutcome_Format.Cells(pageIndex + 2, 3).Value & ")" ' caption
           .Font.Size = 9 ' Size 9
           .Font.bold = True ' Bold
           .left = 10 ' Distance from the left
           .top = 10 ' Distance from the top
           .width = 200
           .height = 10
        End With
        
        ' ContinuousOutcome
        ' subtitle lbel
        controlName = "lblout" & pageIndex & "subtitle1"
         With Me.Controls(controlName)
            .caption = "Continuous" ' caption
            .Font.Size = 9 ' Size 9
            .Font.bold = True ' Bold
            .left = 10 ' Distance from the left
            .top = 25 ' Distance from the top
            .width = 90
            .height = 10
        End With
        
        ' treatment label
        controlName = "lblout" & pageIndex & "sub1tr" & i
        With Me.Controls(controlName)
            .Font.Size = 9 ' Size 9
            .left = 10 + (i - 1) * 100 ' Distance from the left
            .top = 40 ' Distance from the top
            .width = 90
            .height = 10
        End With
        
        ' mean label
        controlName = "lblout" & pageIndex & "meantr" & i
        With Me.Controls(controlName)
                .caption = "μ" & i & " (mean" & i & ")" ' caption
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 55 ' Distance from the top
                .width = 90
                .height = 10
        End With

        ' SD label
        controlName = "lblout" & pageIndex & "sdtr" & i
        With Me.Controls(controlName)
                .caption = "±SD" & i ' caption
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 85 ' Distance from the top
                .width = 90
                .height = 10
        End With

        ' n label
        controlName = "lblout" & pageIndex & "sub1ntr" & i
        With Me.Controls(controlName)
                .caption = "n" & i ' caption
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 115 ' Distance from the top
                .width = 90
                .height = 10
        End With
        
        ' DichotomousOutcome
        ' subtitle label
        controlName = "lblout" & pageIndex & "subtitle2"
        With Me.Controls(controlName)
            .caption = "Dichotomous" ' caption
            .Font.Size = 9 ' Size 9
            .Font.bold = True ' Bold
            .left = 10 ' Distance from the left
            .top = 25 ' Distance from the top
            .width = 90
            .height = 10
        End With
        
        ' treatment label
        controlName = "lblout" & pageIndex & "sub2tr" & i
        With Me.Controls(controlName)
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 40 ' Distance from the top
                .width = 90
                .height = 10
        End With
        
        ' evnet label
        controlName = "lblout" & pageIndex & "eventtr" & i
        With Me.Controls(controlName)
             .caption = "event" & i  ' caption
            .Font.Size = 9 ' Size 9
            .left = 10 + (i - 1) * 100 ' Distance from the left
            .top = 55 ' Distance from the top
            .width = 90
            .height = 10
        End With
        
        ' n label
        controlName = "lblout" & pageIndex & "sub2ntr" & i
        With Me.Controls(controlName)
                .caption = "n" & i  ' caption
                .Font.Size = 9 ' Size 9
                .left = 10 + (i - 1) * 100 ' Distance from the left
                .top = 85 ' Distance from the top
                .width = 90
                .height = 10
        End With
        
        ' cmb label
        controlName = "lblcmb" & pageIndex
        With Me.Controls(controlName)
            .caption = "Outcome"  ' caption
            .Font.Size = 9 ' Size 9
            .Font.bold = True ' Bold
            .left = 10 ' Distance from the left
            .top = 150 ' Distance from the top
            .width = 90
            .height = 10
        End With
        
        ' type label
        controlName = "lbltype" & pageIndex
        With Me.Controls(controlName)
            .caption = "Outcome type"  ' caption
            .Font.Size = 9 ' Size 9
            .Font.bold = True ' Bold
            .left = 160 ' Distance from the left
            .top = 150 ' Distance from the top
            .width = 90
            .height = 10
        End With
    Next i
    ' Set other labels in the same way.

    ' Textbox property setting (txt)
    For i = 1 To 4
        ' ContinuousOutcome
        ' mean textbox
        controlName = "txtout" & pageIndex & "meantr" & i
        With Me.Controls(controlName)
            .left = 10 + (i - 1) * 100
            .top = 65
            .width = 90
            .height = 15
            .IMEMode = 3 ' Set IME mode
            .tabIndex = 1 + (i - 1) * 5
        End With
        
        ' SD textbox
        controlName = "txtout" & pageIndex & "sdtr" & i
        With Me.Controls(controlName)
            .left = 10 + (i - 1) * 100
            .top = 95
            .width = 90
            .height = 15
            .IMEMode = 3
            .tabIndex = 2 + (i - 1) * 5
        End With
        
        ' n textbox
        controlName = "txtout" & pageIndex & "sub1ntr" & i
        With Me.Controls(controlName)
            .left = 10 + (i - 1) * 100
            .top = 125
            .width = 90
            .height = 15
            .IMEMode = 3
            .tabIndex = 3 + (i - 1) * 5
        End With
        
        ' DichotomousOutcome
        ' event textbox
        controlName = "txtout" & pageIndex & "eventtr" & i
        With Me.Controls(controlName)
            .left = 10 + (i - 1) * 100
            .top = 65
            .width = 90
            .height = 15
            .IMEMode = 3
            .tabIndex = 4 + (i - 1) * 5
        End With
        
        ' n textbox
        controlName = "txtout" & pageIndex & "sub2ntr" & i
        With Me.Controls(controlName)
            .left = 10 + (i - 1) * 100
            .top = 95
            .width = 90
            .height = 15
            .IMEMode = 3
            .tabIndex = 5 + (i - 1) * 5
        End With
    Next i
    ' Set other textboxes in the same way.

    ' ComboBoxes property setting (cmb)
    ' List array set
    Dim imax As Long
    Dim ary_outcome As Variant
    Dim ary_type As Variant
    With wsOutcome_Format
        imax = .Cells(.Rows.Count, 1).End(xlUp).row
        ary_outcome = .Range(.Cells(3, 3), .Cells(imax, 3)).Value
        ary_type = .Range(.Cells(3, 2), .Cells(imax, 2)).Value
    End With
    
    ' cmb combobox
    controlName = "cmb" & pageIndex
    With Me.Controls(controlName)
        .left = 10
        .top = 160
        .width = 140
        .height = 15
        .List = Application.Transpose(ary_outcome) ' Set array as list
    End With
    
    'cmbtype combobox
    controlName = "cmbtype" & pageIndex
    With Me.Controls(controlName)
        .left = 160
        .top = 160
        .width = 140
        .height = 15
        .List = Application.Transpose(ary_type) ' Set array as list
    End With
    ' Set other comboboxes in the same way.

    '  CheckBoxes property setting (chk)
    For i = 1 To 4
        'ContinuousOutcome
        'chkout checkbox
        With Me.Controls("chkout" & pageIndex & "sub1tr" & i)
            .caption = "Change n" & i & " of treatment" & i ' caption
            .left = 410
            .top = 65 + (i - 1) * 30
            .width = 140
            .height = 15
        End With
        
        'DichotomousOutcome
        'chkout checkbox
        With Me.Controls("chkout" & pageIndex & "sub2tr" & i)
            .caption = "Change n" & i & " of treatment" & i ' caption
            .left = 410
            .top = 65 + (i - 1) * 30
            .width = 140
            .height = 15
        End With
        
    Next i
    ' Set other checkboxes in the same way.
End Sub

' Show/hide controls based on specific criteria
Private Sub Outcomes_ToggleControlsVisibility(pageIndex As Long, outcomeType As String)
    Dim i As Long
    
    ' Case DichotomousOutcome
    If outcomeType = "Dichotomous" Then
        Me.Controls("lblout" & pageIndex & "subtitle1").Visible = False
        For i = 1 To 4
            Me.Controls("lblout" & pageIndex & "sub1tr" & i).Visible = False
            Me.Controls("lblout" & pageIndex & "meantr" & i).Visible = False
            Me.Controls("lblout" & pageIndex & "sdtr" & i).Visible = False
            Me.Controls("lblout" & pageIndex & "sub1ntr" & i).Visible = False
            Me.Controls("txtout" & pageIndex & "meantr" & i).Visible = False
            Me.Controls("txtout" & pageIndex & "sdtr" & i).Visible = False
            Me.Controls("txtout" & pageIndex & "sub1ntr" & i).Visible = False
            Me.Controls("chkout" & pageIndex & "sub1tr" & i).Visible = False
        Next i
    ' Case ContinuousOutcome
    ElseIf outcomeType = "Continuous" Then
        Me.Controls("lblout" & pageIndex & "subtitle2").Visible = False
        For i = 1 To 4
            Me.Controls("lblout" & pageIndex & "sub2tr" & i).Visible = False
            Me.Controls("lblout" & pageIndex & "eventtr" & i).Visible = False
            Me.Controls("lblout" & pageIndex & "sub2ntr" & i).Visible = False
            Me.Controls("txtout" & pageIndex & "eventtr" & i).Visible = False
            Me.Controls("txtout" & pageIndex & "sub2ntr" & i).Visible = False
            Me.Controls("chkout" & pageIndex & "sub2tr" & i).Visible = False
        Next i
    End If
    ' You can add a process to show/hide controls based on other criteria.
End Sub
Sub updateForm()
    ' Sheet Settings
    Dim wsInput As Worksheet
    Set wsInput = ThisWorkbook.Worksheets("InputSheet")
    wsInput.Activate

    ' Setting values from the row of the active cell to the form control
    UpdateGeneralControlsFromSheet wsInput
    UpdateStrategiesControlsFromSheet wsInput
    UpdateOutcomeControlsFromSheet wsInput
End Sub
' Set common information to form controls
Private Sub UpdateGeneralControlsFromSheet(ws As Worksheet)
    With ws
        Dim activeRow As Long, j As Long
        Dim controlName As String
        
        activeRow = ActiveCell.row
        
        For j = 2 To 7
            controlName = "txt" & .Cells(4, j)
            Me.Controls(controlName) = .Cells(activeRow, j).Value
        Next j
    End With
End Sub

' Setting Strategies to Form Controls
Private Sub UpdateStrategiesControlsFromSheet(ws As Worksheet)
    Dim o As Long, r As Long, arm As Range, arms As Range
    r = 8 ' initial column position
    Dim TableSheet As Worksheet
    Set TableSheet = Worksheets("TableSheet")

    With ws
        Dim activeRow As Long
        activeRow = ActiveCell.row

        For o = 1 To 4
            Me.Controls("txt" & TableSheet.Cells(3, 3).Value & "tr" & o).Text = .Cells(activeRow, r).Value
            Me.Controls("txt" & TableSheet.Cells(5, 1).Value & "tr" & o).Text = .Cells(activeRow, r + 1).Value
            Me.Controls("txtntr" & o).Text = .Cells(activeRow, r + 2).Value
            Me.Controls("cmb" & .Cells(5, 27 + o).Value).Text = .Cells(activeRow, 27 + o)
            Me.Controls("chk" & .Cells(5, 27 + o).Value).Value = False
            r = r + 3
        Next o
        
        ' Define arms range
        Set arms = .Range(.Cells(activeRow, 28), .Cells(activeRow, 31))
        
        o = 1
        For Each arm In arms
            If arm <> "" Then
                Me.Controls("lbl" & "tr" & o).caption = arm & " arm"
            ElseIf arm = "" Then
                Me.Controls("lbl" & "tr" & o).caption = "treatment" & o
            End If
            o = o + 1
        Next arm
    End With
    
End Sub

' Set outcome settings to form controls
Private Sub UpdateOutcomeControlsFromSheet(ws As Worksheet)
    ' Sheet Settings
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    Dim outcomeIndex As Long, wide As Long, k As Long, activeRow As Long, startCol As Long, lastCol As Long
    
    ' For ContinuousOutcome, the number of columns in the combined cell is 16
    ' For DichotomousOutcome, the number of columns in the combined cell is 12
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 16
    DichotomousWide = 12

    outcomeIndex = 1 ' outcomeIndex
   ' Locate the column with "Strategies" and get the last column
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
   
    With wsInput
        activeRow = ActiveCell.row

        For k = startCol To lastCol
            wide = .Cells(3, k).MergeArea.Columns.Count
            
            If wide = ContinuousWide Then ' If the outcome is Continuous
                UpdateContinuousControlsFromSheet .Cells(activeRow, k), outcomeIndex
            ElseIf wide = DichotomousWide Then ' If the outcome is Dichotomous
                UpdateDichotomousControlsFromSheet .Cells(activeRow, k), outcomeIndex
            End If

            outcomeIndex = outcomeIndex + 1
            k = k + wide - 1 ' Jump to next outcome
        Next k
    End With
End Sub

' Update ContinuousOutcome Controls
Private Sub UpdateContinuousControlsFromSheet(startCell As Range, pageIndex As Long)
    Dim m As Long, p As Long, activeRow As Long, arm As Range, arms As Range
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    p = 1 ' Column offset
    For m = 1 To 4
        Me.Controls("txtout" & pageIndex & "meantr" & m).Text = startCell.Offset(0, p).Value
        Me.Controls("txtout" & pageIndex & "sdtr" & m).Text = startCell.Offset(0, p + 1).Value
        Me.Controls("txtout" & pageIndex & "sub1ntr" & m).Text = startCell.Offset(0, p + 2).Value
        Me.Controls("chkout" & pageIndex & "sub1tr" & m).Value = False
        p = p + 4
    Next
    
    ' Define arms range
    With wsInput
        activeRow = ActiveCell.row
        Set arms = .Range(.Cells(activeRow, 28), .Cells(activeRow, 31))
    End With
    
    m = 1
    For Each arm In arms
        If arm <> "" Then
            Me.Controls("lblout" & pageIndex & "sub1tr" & m).caption = arm & " arm"
        ElseIf arm = "" Then
            Me.Controls("lblout" & pageIndex & "sub1tr" & m).caption = "treatment" & m
        End If
        m = m + 1
    Next arm
End Sub

' Update DichotomousOutcome Controls
Private Sub UpdateDichotomousControlsFromSheet(startCell As Range, pageIndex As Long)
    Dim n As Long, q As Long, activeRow As Long, arm As Range, arms As Range
    Dim wsInput As Worksheet
    Set wsInput = Worksheets("InputSheet")
    
    q = 1 ' Column offset
    For n = 1 To 4
        Me.Controls("txtout" & pageIndex & "eventtr" & n).Text = startCell.Offset(0, q).Value
        Me.Controls("txtout" & pageIndex & "sub2ntr" & n).Text = startCell.Offset(0, q + 1).Value
        Me.Controls("chkout" & pageIndex & "sub2tr" & n).Value = False
        q = q + 3
    Next n
    
    ' Define arms range
    With wsInput
        activeRow = ActiveCell.row
        Set arms = .Range(.Cells(activeRow, 28), .Cells(activeRow, 31))
    End With
    
    n = 1
    For Each arm In arms
        If arm <> "" Then
            Me.Controls("lblout" & pageIndex & "sub2tr" & n).caption = arm & " arm"
        ElseIf arm = "" Then
            Me.Controls("lblout" & pageIndex & "sub2tr" & n).caption = "treatment" & n
        End If
        n = n + 1
    Next arm
End Sub
