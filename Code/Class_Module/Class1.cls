VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class1
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit

Public WithEvents btnBack As MSForms.CommandButton
Attribute btnBack.VB_VarHelpID = -1
Public WithEvents btnNext As MSForms.CommandButton
Attribute btnNext.VB_VarHelpID = -1
Public WithEvents btnAdd As MSForms.CommandButton
Attribute btnAdd.VB_VarHelpID = -1
Public WithEvents btnOutcome_Add As MSForms.CommandButton
Attribute btnOutcome_Add.VB_VarHelpID = -1
Public WithEvents btnRowDelete As MSForms.CommandButton
Attribute btnRowDelete.VB_VarHelpID = -1
Public WithEvents btnUpdate As MSForms.CommandButton
Attribute btnUpdate.VB_VarHelpID = -1
' Back buttuon
Private Sub btnBack_Click()
    Call ufSet.Back_method
End Sub
' Next buttuon
Private Sub btnNext_Click()
   Call ufSet.Next_method
End Sub
' Add buttuon
Private Sub btnAdd_Click()
    Call ufSet.Add_method
End Sub
' OutcomeAdd buttuon
Private Sub btnOutcome_Add_Click()
    Call ufSet.Outcome_Add_method
End Sub
' RowDelete buttuon
Private Sub btnRowDelete_Click()
    Call ufSet.RowDelete_method
End Sub
' Update buttuon
Private Sub btnUpdate_Click()
    Call ufSet.Update_method
End Sub
