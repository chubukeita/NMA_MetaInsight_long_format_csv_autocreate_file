Attribute VB_Name = "Module2"
' Module2
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

' This file was output using the VBAC tool; VBAC is bundled with the free VBA library Ariawase.
' Ariawase (and VBAC) are available in the following GitHub repository: https://github.com/vbaidiot/ariawase/tree/master
' The original author is igeta.
' Ariawase and VBAC are used under the MIT License, and the full license can be found at the following link: https://github.com/vbaidiot/ariawase/blob/master/LICENSE.txt

Option Explicit
Function IndexA_sort(IndexA As Long, _
                    sign1 As String, sign2 As String, _
                    High_label As String, Intermediate_label As String, Low_label As String, Zero_label As String, _
                    Intermediate_max_value As Long, Low_max_value As Long, Low_min_value As Long, Zero_value As Long)

                    ' Argument IndexA is the value of target IndexA, Arguments sign1 and sign2 are inequality signs., Argument *_label is a classification label,
                    ' Arguments Intermediate_max_value, Low_max_value,  Low_min_value, and Zero_value are threshold values (boundary IndexA values)
    

    
    
    ' Judges "<" and "<=" inequality in sign1 (judge)
    If InStr(sign1, "=") = 0 Then
        sign1 = "<"
    Else
        sign1 = Mid(sign1, WorksheetFunction.Find("=", sign1), 1)
    End If

    ' Judges "<" and "<=" inequality in sign2 (judge)
    If InStr(sign2, "=") = 0 Then
        sign2 = "<"
    Else
        sign2 = Mid(sign2, WorksheetFunction.Find("=", sign2), 1)
    End If
        
    Dim TableSheet As Worksheet
    Set TableSheet = Worksheets("TableSheet")
    
    With TableSheet
        ' Conditional branching depending on the "<=" and "<" patterns of each inequality sign in sign1 and sign2
        
        ' When both inequality signs of sign1 and sign2 are "<".
        If sign1 = "<" And sign2 = "<" Then
            Select Case IndexA
                Case Is >= Intermediate_max_value
                    IndexA_sort = High_label & "er " & .Cells(3, 3).Value
                Case Is >= Low_max_value
                    IndexA_sort = Intermediate_label & " " & .Cells(3, 3).Value
                Case Is > Low_min_value
                    IndexA_sort = Low_label & "er " & .Cells(3, 3).Value
                Case Is = Zero_value
                    IndexA_sort = Zero_label
            End Select
        ' When the inequality sign of sign1 is "<=" and the inequality sign of sign2 is "<".
        ElseIf sign1 = "=" And sign2 = "<" Then
            Select Case IndexA
                Case Is >= Intermediate_max_value
                    IndexA_sort = High_label & "er " & .Cells(3, 3).Value
                Case Is > Low_max_value
                    IndexA_sort = Intermediate_label & " " & .Cells(3, 3).Value
                Case Is > Low_min_value
                    IndexA_sort = Low_label & "er " & .Cells(3, 3).Value
                Case Is = Zero_value
                    IndexA_sort = Zero_label
            End Select
        ' When the inequality sign of sign1 is "<=" and the inequality sign of sign2 is "<=".
        ElseIf sign1 = "<" And sign2 = "=" Then
            Select Case IndexA
                Case Is > Intermediate_max_value
                    IndexA_sort = High_label & "er " & .Cells(3, 3).Value
                Case Is >= Low_max_value
                    IndexA_sort = Intermediate_label & " " & .Cells(3, 3).Value
                Case Is > Low_min_value
                    IndexA_sort = Low_label & "er " & .Cells(3, 3).Value
                Case Is = Zero_value
                    IndexA_sort = Zero_label
            End Select
        ' When both inequality signs of sign1 and sign2 are "<=".
        ElseIf sign1 = "=" And sign2 = "=" Then
            Select Case IndexA
                Case Is > Intermediate_max_value
                    IndexA_sort = High_label & "er " & .Cells(3, 3).Value
                Case Is > Low_max_value
                    IndexA_sort = Intermediate_label & " " & .Cells(3, 3).Value
                Case Is > Low_min_value
                    IndexA_sort = Low_label & "er " & .Cells(3, 3).Value
                Case Is = Zero_value
                    IndexA_sort = Zero_label
            End Select
        End If
    End With
End Function
Sub RegisterIndexA_sort()
    ' Macro to display help for a function
    
    Application.MacroOptions Macro:="IndexA_sort", Description:= _
    "The IndexA_sort function classifies the target IndexA value into one of four categories: High, Intermediate, Low, or Zero, according to a table of threshold values.", _
    Category:="Lookup/Array", ArgumentDescriptions:=Array _
    ("Argument 1:Target IndexA value", "Argument 2:Low max inequality", "Argument 3:Intermediate max inequality", _
    "Argument 4:High label", "Argument 5:Intermediate label", "Argument 6:Low label", "Argument 7:Zero label", _
    "Argument 8:Intermediate's max threshold", "Argument 9:Low's max threshold", "Argument 10:Low's min threshold", "Argument 11:Zero's threshold 0"), _
    HelpFile:="http://www.microsoft.com/help/helpPage.html"
End Sub
