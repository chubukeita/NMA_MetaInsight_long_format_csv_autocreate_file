Attribute VB_Name = "Module7"
' Module7
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Sub InsertCirclesBasedOnIndexAandIndexB()
    ' Initialize worksheet variables
    Dim SortSheet As Worksheet, InputSheet As Worksheet
    Set SortSheet = Worksheets("SortSheet")
    Set InputSheet = Worksheets("InputSheet")

    ' Variables for loop and cell values
    Dim lastRow As Long, IndexA As Long, IndexB As Long, maru As String
    maru = "ÅZ" ' Circle mark
    lastRow = SortSheet.Cells(SortSheet.Rows.Count, "J").End(xlUp).row

    ' Call related subroutines
    Call combinationIndexAIndexB
    Call SortValueIndexA
    Call SortValueIndexB
    
    ' Loop to insert circles based on IndexA and IndexB values
    Dim i As Long
    For i = 3 To lastRow
        With SortSheet
            IndexA = .Cells(i, 10).Value
            IndexB = .Cells(i, 11).Value
            .Range("D" & i & ":F" & i).Value = "" ' Initialize all columns

            ' Place circles based on IndexA and IndexB values
            Select Case True
                Case IndexA = 1 And IndexB = 1: .Cells(i, "D").Value = maru
                Case IndexA = 2 And IndexB = 1: .Cells(i, "E").Value = maru
                Case IndexA = 3 And IndexB = 1: .Cells(i, "F").Value = maru
                Case IndexA = 0 And IndexB = 2: .Cells(i, "D").Value = maru: .Cells(i, "E").Value = maru: .Cells(i, "F").Value = maru
                Case IndexA = 1 And IndexB = 2: .Cells(i, "D").Value = maru: .Cells(i, "E").Value = maru
                Case IndexA = 2 And IndexB = 2: .Cells(i, "E").Value = maru: .Cells(i, "F").Value = maru
                Case IndexA = 3 And IndexB = 2: .Cells(i, "D").Value = maru: .Cells(i, "F").Value = maru
            End Select

            ' Set up Flash Fill
            .Cells(3, 2).FormulaR1C1 = "=IF(TEXTJOIN(" & Chr(34) & "+" & Chr(34) & ",TRUE,R[0]C[5],R[0]C[6],R[0]C[7])="""",""Nothing"",TEXTJOIN(" & Chr(34) & "+" & Chr(34) & ",TRUE,R[0]C[5],R[0]C[6],R[0]C[7]))"
            .Cells(3, 3).FormulaR1C1 = "=IF(TEXTJOIN("""",TRUE,R[0]C[4],R[0]C[5],R[0]C[6])="""",""Nothing"",TEXTJOIN("""",TRUE,R[0]C[4],R[0]C[5],R[0]C[6]))"
            .Range("B3:C3").AutoFill Destination:=.Range("B3:C" & lastRow), Type:=xlFillDefault
        End With
    Next i

    ' Label A, B, C settings
    Dim arr As Variant
    arr = Array("A", "B", "C")
    For i = 4 To 6
        SortSheet.Cells(2, i) = arr(i - 4)
    Next i

    ' Provide guidance to the user
    MsgBox "You can change the name of the group in cells D2 to F2."
    Dim targetrow As Long
    For targetrow = 3 To lastRow
        With SortSheet
            MsgBox "Think of the arm as being assigned by the combination of the two unique numbers (IndexA and IndexB)." & _
                    " For example, arm" & .Cells(targetrow, 3).Value & " has IndexA=" & .Cells(targetrow, 10).Value & " and IndexB=" & .Cells(targetrow, 11).Value & _
                    ", so enter " & .Cells(targetrow, 10).Value & " in the IndexA column and " & .Cells(targetrow, 11).Value & " in the IndexB column of the treatment on the " & _
                    InputSheet.name & "."
        End With
    Next targetrow
End Sub

Sub InsertABCBasedOnIndexAandIndexB()
    ' Initialize worksheet variables
    Dim SortSheet As Worksheet, InputSheet As Worksheet
    Set SortSheet = Worksheets("SortSheet")
    Set InputSheet = Worksheets("InputSheet")

    ' Variables for last row and IndexA, IndexB values
    Dim lastRow As Long, IndexA As Long, IndexB As Long

    ' Call related subroutines for combination and sorting
    Call combinationIndexAIndexB
    Call SortValueIndexA
    Call SortValueIndexB

    ' Main logic for inserting ABC based on IndexA and IndexB
    With SortSheet
        ' Get the last row with IndexA and IndexB values
        lastRow = .Cells(.Rows.Count, "J").End(xlUp).row

        ' Loop to insert ABC based on IndexA and IndexB values
        Dim i As Long
        For i = 3 To lastRow
            IndexA = .Cells(i, 10).Value
            IndexB = .Cells(i, 11).Value

            ' Initialize all columns by clearing ABC values
            .Range("B" & i & ":F" & i).Value = ""

            ' Assign ABC based on IndexA and IndexB values
            Select Case True
                Case IndexA = 0 And IndexB = 1: .Cells(i, "C").Value = "A"
                Case IndexA = 1 And IndexB = 1: .Cells(i, "C").Value = "B"
                Case IndexA = 2 And IndexB = 1: .Cells(i, "C").Value = "C"
                Case IndexA = 3 And IndexB = 1: .Cells(i, "C").Value = "D"
                Case IndexA = 0 And IndexB = 2: .Cells(i, "C").Value = "E"
                Case IndexA = 1 And IndexB = 2: .Cells(i, "C").Value = "F"
                Case IndexA = 2 And IndexB = 2: .Cells(i, "C").Value = "G"
                Case IndexA = 3 And IndexB = 2: .Cells(i, "C").Value = "H"
            End Select
        Next i

        ' Set ABC labels
        Dim arr As Variant
        arr = Array("A", "B", "C")
        Dim j As Long
        For j = 4 To 6
            .Cells(2, j) = arr(j - 4)
        Next j

        ' Provide guidance to the user
        MsgBox "You can change the name of the group in cells C3 to C10."
        Dim targetrow As Long
        For targetrow = 3 To lastRow
            MsgBox "Think of the arm as being assigned by the combination of the two unique numbers (IndexA and IndexB)." & _
                    " For example, arm" & .Cells(targetrow, 3).Value & " has IndexA=" & .Cells(targetrow, 10).Value & " and IndexB=" & .Cells(targetrow, 11).Value & _
                    ", so enter " & .Cells(targetrow, 10).Value & " in the IndexA column and " & .Cells(targetrow, 11).Value & " in the IndexB column of the treatment on the " & _
                    InputSheet.name & "."
        Next targetrow
    End With
End Sub

Sub combinationIndexAIndexB()
    ' Prepare SortSheet for combination of IndexA and IndexB
    Dim SortSheet As Worksheet
    Set SortSheet = ThisWorkbook.Worksheets("SortSheet")
    Dim i As Long
    
    ' Assign values to IndexA and IndexB
    With SortSheet
        For i = 0 To 3
            .Cells(i + 3, 10).Value = i
            .Cells(i + 3, 11).Value = 1
        Next i
        
        For i = 4 To 7
            .Cells(i + 3, 10).Value = i - 4
            .Cells(i + 3, 11).Value = 2
        Next i
    End With
End Sub

Sub SortValueIndexA()
    ' Prepare IndexA worksheet for sorting values
    Dim wsIndexA As Worksheet
    Set wsIndexA = ThisWorkbook.Worksheets("IndexA")
    Dim i As Long
    
    ' Assign sorting values to IndexA
    With wsIndexA
        For i = 7 To 8
            .Cells(i, 4).Value = "<="
            .Cells(i, 5).Value = i - 6
        Next i
    End With
End Sub

Sub SortValueIndexB()
    ' Prepare IndexB worksheet for sorting values
    Dim wsIndexB As Worksheet
    Set wsIndexB = ThisWorkbook.Worksheets("IndexB")
    
    ' Assign sorting value to IndexB
    With wsIndexB
        .Cells(6, 4).Value = "<="
        .Cells(6, 5).Value = 1
    End With
End Sub

Sub InsertThresholdBasedOnIndexAandIndexB()
    ' Initialize worksheet variables
    Dim SortSheet As Worksheet, InputSheet As Worksheet, StrategiesSheet As Worksheet
    Set SortSheet = Worksheets("SortSheet")
    Set InputSheet = Worksheets("InputSheet")
    Set StrategiesSheet = Worksheets("StrategiesSheet")

    ' Variables for last row and IndexA, IndexB values
    Dim lastRow As Long, IndexA As Long, IndexB As Long
    Dim i As Long, targetrow As Long
    Dim arr() As Variant
    ufSet5.Show

    With SortSheet
        ' Get the last row with IndexA and IndexB values
        lastRow = .Cells(.Rows.Count, "J").End(xlUp).row

        ' Initialize the range
        For i = 3 To lastRow
            .Range("B" & i & ":F" & i).Value = ""
        Next i

        arr = Array("1", "2", "3", "4", "A", "B", "C", "D")

        ' Insert group names based on IndexA and IndexB values
        For i = 3 To lastRow
            IndexA = .Cells(i, 10).Value
            IndexB = .Cells(i, 11).Value

            ' Set group names
            If i <= UBound(arr) + 3 Then
                .Cells(i, 3) = arr(i - 3)
            End If
        Next i

        ' Clear contents of IndexA and IndexB
        Call ClearContentsIndexAIndexB
        
        ' Input values for IndexA and IndexB
        Call InputValueIndexAIndexB

        ' Provide guidance to the user
        SortSheet.Activate
        SortSheet.Cells(3, 3).Activate
        MsgBox "You can change the name of the group in cells C3 to C10."
        For targetrow = 3 To lastRow
            MsgBox "Think of the arm as being assigned by the combination of the sizes of the two variables (IndexA and IndexB)." & _
                    " For example, a group with the combination IndexA=" & .Cells(targetrow, 10).Value & " and IndexB=" & .Cells(targetrow, 11).Value & " " & _
                    "would be classified as " & .Cells(targetrow, 3).Value & " arm from the Classification Table on the " & _
                    StrategiesSheet.name & "."
        Next targetrow
    End With
End Sub

Sub ClearContentsIndexAIndexB()
    ' Clear contents of IndexA and IndexB in SortSheet
    Dim SortSheet As Worksheet
    Set SortSheet = ThisWorkbook.Worksheets("SortSheet")
    
    With SortSheet
        .Range(.Cells(3, 10), .Cells(10, 11)).Clearcontents
    End With
End Sub

Sub InputValueIndexAIndexB()
    ' Initialize worksheet variables
    Dim SortSheet As Worksheet, wsIndexA As Worksheet, wsIndexB As Worksheet
    
    Set SortSheet = ThisWorkbook.Worksheets("SortSheet")
    Set wsIndexA = ThisWorkbook.Worksheets("IndexA")
    Set wsIndexB = ThisWorkbook.Worksheets("IndexB")

    ' Initialize default values for IndexA and IndexB
    With SortSheet
        .Cells(3, 10).Value = 0
        .Cells(7, 10).Value = 0
    End With
    
    ' Set values for IndexA
    With wsIndexA
        ' Generate random values based on the inequality sign
        If .Cells(7, 4).Value = "<" And .Cells(8, 4).Value = "<" Then
            AssignRandomValuesToSortSheetIndexA SortSheet, .Cells(7, 2).Value, .Cells(7, 5).Value, .Cells(8, 2).Value, .Cells(8, 5).Value, .Cells(9, 2).Value, "<", "<"
        ElseIf .Cells(7, 4).Value = "<=" And .Cells(8, 4).Value = "<" Then
            AssignRandomValuesToSortSheetIndexA SortSheet, .Cells(7, 2).Value, .Cells(7, 5).Value, .Cells(8, 2).Value, .Cells(8, 5).Value, .Cells(9, 2).Value, "<=", "<"
        ElseIf .Cells(7, 4).Value = "<" And .Cells(8, 4).Value = "<=" Then
            AssignRandomValuesToSortSheetIndexA SortSheet, .Cells(7, 2).Value, .Cells(7, 5).Value, .Cells(8, 2).Value, .Cells(8, 5).Value, .Cells(9, 2).Value, "<", "<="
        ElseIf .Cells(7, 4).Value = "<=" And .Cells(8, 4).Value = "<=" Then
            AssignRandomValuesToSortSheetIndexA SortSheet, .Cells(7, 2).Value, .Cells(7, 5).Value, .Cells(8, 2).Value, .Cells(8, 5).Value, .Cells(9, 2).Value, "<=", "<="
        Else
            MsgBox "The inequality sign for IndexA is not entered"
        End If
    End With
    
    ' Set values for IndexB
    With wsIndexB
        ' Generate random values based on the inequality sign
        If .Cells(6, 4).Value = "<" Then
            AssignRandomValuesToSortSheetIndexB SortSheet, .Cells(6, 2).Value, .Cells(6, 5).Value, .Cells(7, 2).Value, "<"
        ElseIf .Cells(6, 4).Value = "<=" Then
            AssignRandomValuesToSortSheetIndexB SortSheet, .Cells(6, 2).Value, .Cells(6, 5).Value, .Cells(7, 2).Value, "<="
        Else
            MsgBox "The inequality sign for IndexB is not entered"
        End If
    End With
End Sub

' Auxiliary function to assign a random value to IndexA
Private Sub AssignRandomValuesToSortSheetIndexA(ByRef SortSheet As Worksheet, lowerBound1 As Integer, upperBound1 As Integer, lowerBound2 As Integer, upperBound2 As Integer, lowerBound3 As Integer, sign1 As String, sign2 As String)
    With SortSheet
        If sign1 = "<" And sign2 = "<" Then
            .Cells(4, 10).Value = RandomExclusiveExclusive(lowerBound1, upperBound1)
            .Cells(8, 10).Value = RandomExclusiveExclusive(lowerBound1, upperBound1)
            .Cells(5, 10).Value = RandomInclusiveExclusive(lowerBound2, upperBound2)
            .Cells(9, 10).Value = RandomInclusiveExclusive(lowerBound2, upperBound2)
            .Cells(6, 10).Value = RandomInclusiveInclusive(lowerBound3, 10000)
            .Cells(10, 10).Value = RandomInclusiveInclusive(lowerBound3, 10000)
        ElseIf sign1 = "<=" And sign2 = "<" Then
            .Cells(4, 10).Value = RandomExclusiveInclusive(lowerBound1, upperBound1)
            .Cells(8, 10).Value = RandomExclusiveInclusive(lowerBound1, upperBound1)
            .Cells(5, 10).Value = RandomExclusiveExclusive(lowerBound2, upperBound2)
            .Cells(9, 10).Value = RandomExclusiveExclusive(lowerBound2, upperBound2)
            .Cells(6, 10).Value = RandomInclusiveInclusive(lowerBound3, 10000)
            .Cells(10, 10).Value = RandomInclusiveInclusive(lowerBound3, 10000)
        ElseIf sign1 = "<" And sign2 = "<=" Then
            .Cells(4, 10).Value = RandomExclusiveExclusive(lowerBound1, upperBound1)
            .Cells(8, 10).Value = RandomExclusiveExclusive(lowerBound1, upperBound1)
            .Cells(5, 10).Value = RandomInclusiveInclusive(lowerBound2, upperBound2)
            .Cells(9, 10).Value = RandomInclusiveInclusive(lowerBound2, upperBound2)
            .Cells(6, 10).Value = RandomExclusiveInclusive(lowerBound3, 10000)
            .Cells(10, 10).Value = RandomExclusiveInclusive(lowerBound3, 10000)
        ElseIf sign1 = "<=" And sign2 = "<=" Then
            .Cells(4, 10).Value = RandomExclusiveInclusive(lowerBound1, upperBound1)
            .Cells(8, 10).Value = RandomExclusiveInclusive(lowerBound1, upperBound1)
            .Cells(5, 10).Value = RandomExclusiveInclusive(lowerBound2, upperBound2)
            .Cells(9, 10).Value = RandomExclusiveInclusive(lowerBound2, upperBound2)
            .Cells(6, 10).Value = RandomExclusiveInclusive(lowerBound3, 10000)
            .Cells(10, 10).Value = RandomExclusiveInclusive(lowerBound3, 10000)
        End If
    End With
End Sub

' Auxiliary functions for assigning random values to IndexB
Private Sub AssignRandomValuesToSortSheetIndexB(ByRef SortSheet As Worksheet, lowerBound As Integer, upperBound As Integer, threshold As Integer, sign As String)
     With SortSheet
         If sign = "<" Then
            .Cells(3, 11).Value = RandomExclusiveExclusive(lowerBound, upperBound)
            .Cells(4, 11).Value = RandomExclusiveExclusive(lowerBound, upperBound)
            .Cells(5, 11).Value = RandomExclusiveExclusive(lowerBound, upperBound)
            .Cells(6, 11).Value = RandomExclusiveExclusive(lowerBound, upperBound)
            .Cells(7, 11).Value = RandomInclusiveInclusive(threshold, 10000)
            .Cells(8, 11).Value = RandomInclusiveInclusive(threshold, 10000)
            .Cells(9, 11).Value = RandomInclusiveInclusive(threshold, 10000)
            .Cells(10, 11).Value = RandomInclusiveInclusive(threshold, 10000)
        ElseIf sign = "<=" Then
            .Cells(3, 11).Value = RandomExclusiveInclusive(lowerBound, upperBound)
            .Cells(4, 11).Value = RandomExclusiveInclusive(lowerBound, upperBound)
            .Cells(5, 11).Value = RandomExclusiveInclusive(lowerBound, upperBound)
            .Cells(6, 11).Value = RandomExclusiveInclusive(lowerBound, upperBound)
            .Cells(7, 11).Value = RandomExclusiveInclusive(threshold, 10000)
            .Cells(8, 11).Value = RandomExclusiveInclusive(threshold, 10000)
            .Cells(9, 11).Value = RandomExclusiveInclusive(threshold, 10000)
            .Cells(10, 11).Value = RandomExclusiveInclusive(threshold, 10000)
        End If
    End With
End Sub

Function RandomInclusiveInclusive(lowerBound As Integer, upperBound As Integer)
    ' Selects a random integer greater than or equal to lowerBound and less than or equal to upperBound
    Randomize
    RandomInclusiveInclusive = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
End Function

Function RandomExclusiveInclusive(lowerBound As Integer, upperBound As Integer)
    ' Selects a random integer greater than lowerBound and less than or equal to upperBound
    Randomize
    RandomExclusiveInclusive = Int((upperBound - (lowerBound + 1) + 1) * Rnd + (lowerBound + 1))
End Function

Function RandomInclusiveExclusive(lowerBound As Integer, upperBound As Integer)
    ' Selects a random integer greater than or equal to lowerBound and less than upperBound
    Randomize
    RandomInclusiveExclusive = Int(((upperBound - 1) - lowerBound + 1) * Rnd + lowerBound)
End Function

Function RandomExclusiveExclusive(lowerBound As Integer, upperBound As Integer)
    ' Selects a random integer greater than lowerBound and less than upperBound
    ' However, there must be at least a difference of 1 between the lower and upper bounds
    If (upperBound - lowerBound) < 2 Then
        MsgBox "An error has occurred. The message box that follows is incorrect; please try the process again."
        RandomExclusiveExclusive = "error"
        Exit Function
    End If

    Randomize ' Initialize the random number generator
    Dim result As Integer
    RandomExclusiveExclusive = Int((upperBound - 1 - (lowerBound + 1) + 1) * Rnd + (lowerBound + 1))
End Function
