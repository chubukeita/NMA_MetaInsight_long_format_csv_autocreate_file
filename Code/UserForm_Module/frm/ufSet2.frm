VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSet2 
   Caption         =   "Outcome Setting"
   ClientHeight    =   2290
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2480
   OleObjectBlob   =   "ufSet2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufSet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ufSet2
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) 2024 chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Private Sub btnOutcome_Insort_Click()
    Dim wsInput As Worksheet
    Dim i As Long, imax As Long, name As String
    Dim Colmax As Long
    
    ' Warning before adding outcome
    Dim rc As VbMsgBoxResult
    rc = MsgBox("Add the outcome to the right side of the table. Are you sure?", vbCritical + vbOKCancel, "Warning icon")
    If rc = vbCancel Then
        Exit Sub
    End If
    
    ufSet4.Show vbModeless
    ufSet4.Repaint
    
    name = Controls("txtoutcome_name").Text
    
    ' If Outcome_type is ContinuousOutcome
    If Controls("cmboutcome_type").ListIndex = 0 Then
        Call binding_Continuous(name)
    ' If Outcome_type is DichotomousOutcome
    ElseIf Controls("cmboutcome_type").ListIndex = 1 Then
        Call binding_Dichotomous(name)
    End If
    
    Call outcome_Format
    
    Set wsInput = Worksheets("InputSheet")
    
    With wsInput
        Colmax = .Cells(5, .Columns.Count).End(xlToLeft).Column
    End With
    
    wsInput.Activate
    wsInput.Cells(3, Colmax).Activate
    Unload ufSet2
    Unload ufSet4
    MsgBox ("outcome addition completed")
End Sub
Private Sub UserForm_Initialize()
    Dim wsOutcome_Type As Worksheet
    Dim i As Long
    Dim imax As Long
    Dim ary_t
    
    Set wsOutcome_Type = Worksheets("outcome_type")
  
    With wsOutcome_Type
        imax = .Cells(.Rows.Count, 2).End(xlUp).row
        ary_t = .Range(.Cells(3, 2), .Cells(imax, 2)).Value
        Controls("cmboutcome_type").List = ary_t
    End With
End Sub

