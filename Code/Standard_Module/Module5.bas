Attribute VB_Name = "Module5"
' Module5
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Sub binding_Continuous(name As String)
    Dim wsInput As Worksheet, wsContinuous As Worksheet
    Dim LastColoredRow As Long, startCol As Long, lastCol As Long
    
    Dim arm As Range, arms As Range, armnum As Long
    Dim patient As Range, patients As Range, patientnum As Long
    
    ' Initialize worksheets
    Set wsContinuous = Worksheets("ContinuousSheet")
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        ' Get the last row with color after the last row with text
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check 2nd column
    End With
    
    ' Combine data from ContinuousSheet to wsInput
    With wsContinuous
        ' Sheet initialization
        ' Delete values
        Call Clearcontents(wsContinuous)
        
        ' Change cells filled with paleturquoise color to no color
        Call Clear_PaleTurquoise_Interior(wsContinuous)
    
        ' Use Union to define ranges for arms and patients
        Set arms = Union(.Cells(6, 2), .Cells(6, 6), .Cells(6, 10), .Cells(6, 14))
        Set patients = Union(.Cells(6, 5), .Cells(6, 9), .Cells(6, 13), .Cells(6, 17))
        
        ' Get columns where Strategies are present
        Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
        
        ' Enter formulas for arms
        armnum = startCol - 4 ' Get column number for arm1 in Strategies
        For Each arm In arms
            arm.FormulaR1C1 = "=InputSheet!RC" & armnum
            armnum = armnum + 1
        Next arm
        
        ' Enter formulas for patients
        patientnum = 10
        For Each patient In patients
            patient.FormulaR1C1 = "=IF(InputSheet!RC" & patientnum & "="""","""",InputSheet!RC" & patientnum & ")"
            patientnum = patientnum + 3
        Next patient
        
        ' Autofill formulas to the last row
        .Range("B6:Q6").AutoFill .Range("B6:Q" & LastColoredRow), Type:=xlFillDefault
        
        ' Copy results to wsInput
        .Range(.Cells(3, 2), .Cells(LastColoredRow, 2 + 15)).Copy wsInput.Cells(3, lastCol + 1)
        
        ' Add outcome name
        wsInput.Cells(3, lastCol + 1).Activate
        wsInput.Cells(3, lastCol + 1).Value = name
        
    End With
End Sub
Sub binding_Dichotomous(name As String)
    Dim wsInput As Worksheet, wsDichotomous As Worksheet
    Dim LastColoredRow As Long, startCol As Long, lastCol As Long
    
    Dim arm As Range, arms As Range, armnum As Long
    Dim patient As Range, patients As Range, patientnum As Long
    
    ' Initialize worksheets
    Set wsDichotomous = Worksheets("DichotomousSheet")
    Set wsInput = Worksheets("InputSheet")
    
    ' Find the last row and column with data in wsInput
    With wsInput
        ' Get the last row with color after the last row with text
        LastColoredRow = GetLastColoredRowAfterText(wsInput, 2) ' Check 2nd column
    End With
    
    ' Combine data from DichotomousSheet to wsInput
    With wsDichotomous
        ' Sheet initialization
        ' Delete values
        Call Clearcontents(wsDichotomous)
        
        ' Change cells filled with paleturquoise color to no color
        Call Clear_PaleTurquoise_Interior(wsDichotomous)
    
        ' Use Union to define ranges for arms and patients
        Set arms = Union(.Cells(6, 2), .Cells(6, 5), .Cells(6, 8))
        Set patients = Union(.Cells(6, 4), .Cells(6, 7), .Cells(6, 10))
        
        ' Get columns where Strategies are present
        Call FindStrategiesAndSetColumns(wsInput, startCol, lastCol)
        
        ' Enter formulas for arms
        armnum = startCol - 4
        For Each arm In arms
            arm.FormulaR1C1 = "=InputSheet!RC" & armnum
            armnum = armnum + 1
        Next arm
        
        ' Enter formulas for patients
        patientnum = 10
        For Each patient In patients
            patient.FormulaR1C1 = "=IF(InputSheet!RC" & patientnum & "="""","""",InputSheet!RC" & patientnum & ")"
            patientnum = patientnum + 3
        Next patient
        
        ' Autofill formulas to the last row
        .Range("B6:M6").AutoFill .Range("B6:M" & LastColoredRow), Type:=xlFillDefault
        
        ' Copy results to wsInput
        .Range(.Cells(3, 2), .Cells(LastColoredRow, 2 + 11)).Copy wsInput.Cells(3, lastCol + 1)
        
        ' Add outcome name
        wsInput.Cells(3, lastCol + 1).Value = name
        
    End With
End Sub
Sub Clearcontents(Sheet As Worksheet)
    With Sheet
        ' Delete values
        .Range(.Cells(6, 2), .Cells(350, 17)).Clearcontents
    End With
End Sub
Sub Clear_PaleTurquoise_Interior(Sheet As Worksheet)
    Dim r As Range
    With Sheet
        ' Change cells filled with paleturquoise color to no color
        For Each r In .Range(.Cells(6, 2), .Cells(350, 17))
            If r.Interior.Color = rgbPaleTurquoise Then
                r.Interior.Color = xlNone
            End If
        Next r
    End With
End Sub
Function GetLastColoredRowAfterText(ws As Worksheet, CheckColumn As Long) As Long
    Dim LastRowWithText As Long, LastRowWithColor As Long, i As Long

    Application.ScreenUpdating = False
    ' Get the last row with text
    LastRowWithText = ws.Cells(ws.Rows.Count, CheckColumn).End(xlUp).row

    ' Initialize the last row with color
    LastRowWithColor = LastRowWithText

    ' Scan the specified column from bottom to top and search for colored cells
    For i = 350 To LastRowWithText + 1 Step -1
        If ws.Cells(i, CheckColumn).Interior.ColorIndex <> xlNone Then
            LastRowWithColor = i
            Exit For
        End If
    Next i
    Application.ScreenUpdating = True

    ' Return the last row where color was found
    GetLastColoredRowAfterText = LastRowWithColor
End Function

