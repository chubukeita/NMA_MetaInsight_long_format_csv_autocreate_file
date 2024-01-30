Attribute VB_Name = "Module3"
' Module3
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Function IndexB_sort(IndexB As Long, sign1 As String, _
                    High_label As String, Low_label As String, _
                    Low_max_value As Long)
                    ' Argument IndexB is the value of target IndexB, Arguments sign1 are inequality signs., Argument *_label is a classification label,
                    ' Arguments Low_max_value are threshold values (boundary IndexB values)
    
    ' Judges "<" and "<=" inequality in sign1 (judge)
    If InStr(sign1, "=") = 0 Then
        sign1 = "<"
    Else
        sign1 = Mid(sign1, WorksheetFunction.Find("=", sign1), 1)
    End If
    
    Dim TableSheet As Worksheet
    Set TableSheet = Worksheets("TableSheet")
    
    With TableSheet
        ' Conditional branching depending on the "<=" and "<" patterns of each inequality sign in sign1
        
        ' When the inequality sign of sign1 is "<"
        If sign1 = "<" Then
            Select Case IndexB
                Case Is >= Low_max_value
                    IndexB_sort = High_label & "er " & .Cells(5, 1).Value
                Case Is > 0
                    IndexB_sort = Low_label & "er " & .Cells(5, 1).Value
            End Select
        ' When the inequality sign of sign1 is "<="
        ElseIf sign1 = "=" Then
            Select Case IndexB
                Case Is > Low_max_value
                    IndexB_sort = High_label & "er " & .Cells(5, 1).Value
                Case Is > 0
                    IndexB_sort = Low_label & "er " & .Cells(5, 1).Value
            End Select
        End If
    End With
End Function
Sub RegisterIndexB_sort()
    ' Macro to display help for a function
    
    Application.MacroOptions Macro:="IndexB_sort2", Description:= _
    "The IndexB_sort function classifies the target IndexB values into two categories, High and Low, according to a table of threshold values.", _
    Category:="Lookup/Array", ArgumentDescriptions:=Array _
    ("Argument 1:Target IndexB value", "Argument 2:Low max inequality", _
    "Argument 3:High label", "Argument 4:Low label", _
    "Argument 5:Low's max threshold"), _
    HelpFile:="http://www.microsoft.com/help/helpPage.html"
End Sub


