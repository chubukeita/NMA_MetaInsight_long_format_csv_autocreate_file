Attribute VB_Name = "Module6"
' Module6
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

' This file was output using the VBAC tool; VBAC is bundled with the free VBA library Ariawase.
' Ariawase (and VBAC) are available in the following GitHub repository: https://github.com/vbaidiot/ariawase/tree/master
' The original author is Copyright (c) 2011 igeta.
' Ariawase and VBAC are used under the MIT License, and the full license can be found at the following link: https://github.com/vbaidiot/ariawase/blob/master/LICENSE.txt

Option Explicit
Sub InputRowInsert()
    ' Initialize worksheet variables
    Dim wsInput As Worksheet
    Dim maxInputRow As Long, startRow As Long, startCol As Long, lastCol As Long
    
    Dim name As String
    Dim j As Long, wide As Long, k As Long, LastColoredRow As Long
    
    ' Define the width of merged cells for Continuous and Dichotomous outcomes
    Dim ContinuousWide As Long, DichotomousWide As Long
    ContinuousWide = 16 ' Width for Continuous outcome merged cells
    DichotomousWide = 12  ' Width for Dichotomous outcome merged cells
    
    ' Set reference to the InputSheet
    Set wsInput = Worksheets("InputSheet")
    wsInput.Activate
    
    ' Get columns where strategies exist
    Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
    
    ' Start processing the InputSheet
    With wsInput
        ' Find the last row with data
        maxInputRow = .Cells(.Rows.Count, 2).End(xlUp).row
        
        ' Get the last colored row after the last row with text
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check in column 2
    
        ' Add Information and Treatment
        ' Copy the last row with data and insert it below
        .Range(.Cells(maxInputRow, 2), .Cells(maxInputRow, lastCol)).Select
        Selection.Copy
        Selection.Insert xlDown
        Application.CutCopyMode = False
        
        ' Clear the contents of the newly added row from startCol to lastCol_tr4
        Dim foundCell As Range
        Dim lastCol_tr4 As Long
        Set foundCell = wsInput.Rows(4).Find("treatment4")
        If Not foundCell Is Nothing Then
            lastCol_tr4 = foundCell.Column + foundCell.MergeArea.Columns.Count - 1
            .Range(.Cells(maxInputRow + 1, 2), .Cells(maxInputRow + 1, lastCol_tr4)).Clearcontents
        Else
            MsgBox "treatment4 not found."
            Exit Sub
        End If

        ' Iterate through each column from startCol to lastCol
        For j = startCol To lastCol
            ' Determine the width of the current column (number of merged cells) and get its name
            
            wide = .Cells(3, j).MergeArea.Columns.Count
            name = .Cells(3, j).Value
            
            ' Depending on the width, call the appropriate function to insert Continuous or Dichotomous outcomes
            If wide = ContinuousWide Then
                Call Insert_Continuous(name, j)
            ElseIf wide = DichotomousWide Then
                Call Insert_Dichotomous(name, j)
            End If
            
            ' Move to the next merged cell block
            j = j + wide - 1
        Next j

        ' Apply formatting for outcomes
        Call outcome_Format
        
        ' Reactivate the InputSheet and move the cursor to the newly added row
        .Activate
        .Cells(maxInputRow + 1, 2).Activate
    End With
End Sub

Sub Insert_Continuous(name, j)
    ' Initialize worksheet variables
    Dim wsInput As Worksheet, wsContinuous As Worksheet
    Dim LastColoredRow As Long, startCol As Long, lastCol As Long, k As Long, wide As Long, outcomeIndex As Long
    outcomeIndex = 1
    
    ' Define ranges for arms and patients
    Dim arm As Range, arms As Range, armnum As Long
    Dim patient As Range, patients As Range, patientnum As Long
    
    ' Set worksheet references
    Set wsContinuous = Worksheets("ContinuousSheet")
    Set wsInput = Worksheets("InputSheet")
    
    ' Find the last colored row after the last row with text in wsInput
    With wsInput
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check in column 2
    End With
    
    ' Combine data from ContinuousSheet to wsInput
    With wsContinuous
        ' Clear existing values
        Call Clearcontents(wsContinuous)
        
        ' Change red-filled cells to no color
        Call Clear_Red_Interior(wsContinuous)
    
        ' Define ranges for arms and patients using Union
        Set arms = Union(.Cells(6, 2), .Cells(6, 6), .Cells(6, 10), .Cells(6, 14))
        Set patients = Union(.Cells(6, 5), .Cells(6, 9), .Cells(6, 13), .Cells(6, 17))
        
        ' Get columns with strategies in wsInput
        Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
        
        ' Insert formulas for arms
        armnum = startCol - 4 ' Column number for arm1 in Strategies
        For Each arm In arms
            arm.FormulaR1C1 = "=InputSheet!RC" & armnum
            armnum = armnum + 1
        Next arm
        
        ' Insert formulas for patients
        patientnum = 10
        For Each patient In patients
            patient.FormulaR1C1 = "=IF(InputSheet!RC" & patientnum & "="""","""",InputSheet!RC" & patientnum & ")"
            patientnum = patientnum + 3
        Next patient
        
        ' Autofill formulas to the last colored row
        .Range("B6:Q6").AutoFill .Range("B6:Q" & LastColoredRow), Type:=xlFillDefault
        
        ' Copy the results to wsInput
        .Range(.Cells(6, 2), .Cells(6, 2 + 15)).Copy wsInput.Cells(LastColoredRow, j)
    End With
End Sub
Sub Insert_Dichotomous(name, j)
    ' Initialize worksheet variables
    Dim wsInput As Worksheet, wsDichotomous As Worksheet
    Dim LastColoredRow As Long, startCol As Long, lastCol As Long, k As Long, wide As Long, outcomeIndex As Long
    
    Dim arm As Range, arms As Range, armnum As Long
    Dim patient As Range, patients As Range, patientnum As Long
    
    outcomeIndex = 1
    
    ' Set worksheet references
    Set wsDichotomous = Worksheets("DichotomousSheet")
    Set wsInput = Worksheets("InputSheet")
    
    ' Find the last colored row after the last row with text in wsInput
    With wsInput
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check in column 2
    End With
    
    ' Combine data from DichotomousSheet to wsInput
    With wsDichotomous
        ' Clear existing values
        Call Clearcontents(wsDichotomous)
        
        ' Change red-filled cells to no color
        Call Clear_Red_Interior(wsDichotomous)
    
        ' Define ranges for arms and patients using Union
        Set arms = Union(.Cells(6, 2), .Cells(6, 5), .Cells(6, 8), .Cells(6, 11))
        Set patients = Union(.Cells(6, 4), .Cells(6, 7), .Cells(6, 10), .Cells(6, 13))
        
        ' Get columns with strategies in wsInput
        Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
        
        ' Insert formulas for arms
        armnum = startCol - 4
        For Each arm In arms
            arm.FormulaR1C1 = "=InputSheet!RC" & armnum
            armnum = armnum + 1
        Next arm
        
        ' Insert formulas for patients
        patientnum = 10
        For Each patient In patients
            patient.FormulaR1C1 = "=IF(InputSheet!RC" & patientnum & "="""","""",InputSheet!RC" & patientnum & ")"
            patientnum = patientnum + 3
        Next patient
        
        ' Autofill formulas to the last colored row
        .Range("B6:M6").AutoFill .Range("B6:M" & LastColoredRow), Type:=xlFillDefault
        
        ' Copy the results to wsInput
        .Range(.Cells(6, 2), .Cells(6, 2 + 11)).Copy wsInput.Cells(LastColoredRow, j)
    End With
End Sub




