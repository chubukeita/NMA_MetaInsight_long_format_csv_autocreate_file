VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class2
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit

Public WithEvents btnIndexAinput As MSForms.CommandButton
Attribute btnIndexAinput.VB_VarHelpID = -1
Public WithEvents btnIndexBinput As MSForms.CommandButton
Attribute btnIndexBinput.VB_VarHelpID = -1

Private Sub btnIndexAinput_Click()
    Call ufSet5.IndexAinput_method
End Sub
Private Sub btnIndexBinput_Click()
    Call ufSet5.IndexBinput_method
End Sub
