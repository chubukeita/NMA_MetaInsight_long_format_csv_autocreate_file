VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSet5 
   Caption         =   "Threshold settings"
   ClientHeight    =   6570
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7760
   OleObjectBlob   =   "ufSet5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufSet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ufSet5
' This code was created by chubukeita and refactored by ChatGPT 4.0.
' Copyright (c) chubukeita, subject to MIT license.
' More information about the new license can be found at the following link: https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

Option Explicit
Private ButtonHandler As New Collection
Sub IndexAinput_method()
    Dim wsIndexA As Worksheet
    Set wsIndexA = Worksheets("IndexA")

    With wsIndexA
        ' Retrieve values using VLOOKUP function based on the selected index from the combo box
        .Cells(7, 4).Value = WorksheetFunction.VLookup(Me.Controls("cmbIndexAsign1").ListIndex, .Range(.Cells(5, 10), .Cells(6, 12)), 2, False)
        .Cells(8, 4).Value = WorksheetFunction.VLookup(Me.Controls("cmbIndexAsign2").ListIndex, .Range(.Cells(5, 10), .Cells(6, 12)), 2, False)

        ' Set values from the text boxes
        .Cells(7, 5).Value = Me.Controls("txtIndexALow_max_value").Value
        .Cells(8, 5).Value = Me.Controls("txtIndexAIntermediate_max_value").Value
    End With

    MultiPage1.Value = 1 ' Switch the tab of MultiPage
End Sub

Sub IndexBinput_method()
    Dim wsIndexB As Worksheet
    Set wsIndexB = Worksheets("IndexB")
    
    With wsIndexB
        ' Retrieve values using VLOOKUP function based on the selected index from the combo box
        .Cells(6, 4).Value = WorksheetFunction.VLookup(Controls("cmbIndexBsign1").ListIndex, .Range(.Cells(5, 10), .Cells(6, 12)), 2, False)
        
        ' Set values from the text boxes
        .Cells(6, 5).Value = Controls("txtIndexBLow_max_value").Value
    End With
    
    Unload ufSet5 ' Close ufSet5 form
End Sub

Private Sub MultiPage1_Change()
    Dim wsIndexA As Worksheet, wsIndexB As Worksheet
    
    Set wsIndexA = Worksheets("IndexA")
    Set wsIndexB = Worksheets("IndexB")
     ' Perform actions based on the selected page
    Select Case MultiPage1.Value
        Case 0 ' First page (e.g., Page1)
            wsIndexA.Activate
        Case 1 ' Second page (e.g., Page2)
            wsIndexB.Activate
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim wsIndexA As Worksheet, wsIndexB As Worksheet
    Dim i As Long, imax As Long
    Dim ary_t
    
    Set wsIndexA = Worksheets("IndexA")
    Set wsIndexB = Worksheets("IndexB")
    
    wsIndexA.Activate
  
    MultiPage1.Pages(0).caption = wsIndexA.Cells(4, 2).Value
    MultiPage1.Pages.Add wsIndexB.Cells(4, 2).Value
    
    ' Place various controls on each outcome page
    Call IndexA_AddLabelsToPage(0)
    Call IndexA_AddTextBoxesToPage(0)
    Call IndexA_AddComboBoxesToPage(0)
    Call IndexA_AddButtonsToPage(0)

    ' Set properties for various controls on each outcome page
    Call IndexA_SetControlsProperties(i)
    With wsIndexA
        ary_t = .Range(.Cells(5, 12), .Cells(6, 12))
    
        For i = 1 To 2
            Controls("cmbIndexAsign" & i).List = ary_t
        Next i
    End With
    
    ' Place various controls on each outcome page
    Call IndexB_AddLabelsToPage(1)
    Call IndexB_AddTextBoxesToPage(1)
    Call IndexB_AddComboBoxesToPage(1)
    Call IndexB_AddButtonsToPage(1)

    ' Set properties for various controls on each outcome page
    Call IndexB_SetControlsProperties(1)
    With wsIndexB
        ary_t = .Range(.Cells(5, 12), .Cells(6, 12))
        Controls("cmbIndexBsign1").List = ary_t
    End With
End Sub
Sub IndexA_AddLabelsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newLabel As Control

    ' Create list of control names
    controlNames = Array("IndexALow_max_value", "IndexAIntermediate_max_value")

    ' Add labels beginning with lbl
    For i = LBound(controlNames) To UBound(controlNames)
        Set newLabel = MultiPage1.Pages(pageIndex).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
    Next i
    
    controlNames = Array("IndexAsign1", "IndexAsign2")

    
    For i = LBound(controlNames) To UBound(controlNames)
        ' lblcmb and lbltype have different name patterns, so add them separately
        Set newLabel = MultiPage1.Pages(pageIndex).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
        Next i
End Sub
Sub IndexA_AddTextBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newText As Control

    ' Create list of control names
    controlNames = Array("IndexALow_max_value", "IndexAIntermediate_max_value")

    ' Add textboxes beginning with txt
    For i = LBound(controlNames) To UBound(controlNames)
        Set newText = MultiPage1.Pages(pageIndex).Controls.Add("Forms.TextBox.1", "txt" & controlNames(i))
        ' You can set other properties of newText here
    Next i
End Sub
Sub IndexA_AddComboBoxesToPage(pageIndex As Long)
    Dim controlPrefixes As Variant, prefix As Variant
    Dim newCombo As Control

    ' Create list of control names
    controlPrefixes = Array("IndexAsign1", "IndexAsign2")

    ' Add comboboxes beginning with cmb
    For Each prefix In controlPrefixes
        Set newCombo = MultiPage1.Pages(pageIndex).Controls.Add("Forms.ComboBox.1", "cmb" & prefix)
        ' You can set other properties of newCombo here
    Next prefix
End Sub
Sub IndexA_AddButtonsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newButton As Control, newText As Control

    ' Create list of control names
    controlNames = Array("IndexAinput")

    ' Add commmandbuttons beginning with btn
    For i = LBound(controlNames) To UBound(controlNames)
        Set newButton = MultiPage1.Pages(pageIndex).Controls.Add("Forms.CommandButton.1", "btn" & controlNames(i))
        ' You can set other properties of newButton here
        Dim buttonEvent As New Class2
        
        ' Event Handler Assignment
        With buttonEvent
            Set .btnIndexAinput = newButton
        End With
        
        ButtonHandler.Add buttonEvent
        
    Next i
End Sub
Private Sub IndexA_SetControlsProperties(pageIndex As Long)
    Dim i As Long
    Dim wsIndexA As Worksheet
    Set wsIndexA = Worksheets("IndexA")
    
    With wsIndexA
        ' Label property setting (lbl)
        For i = 1 To 2
            IndexA_SetLabelProperties "lblIndexAsign" & i, 60, 10 + (i - 1) * 30, 180, 20, "Please select less than or or below from the list"
        Next i
        
        IndexA_SetLabelProperties "lblIndexALow_max_value", 10, 10, 50, 20, .Cells(7, 5).Address(False, False) & " Cell Value"
        IndexA_SetLabelProperties "lblIndexAIntermediate_max_value", 10, 40, 50, 20, .Cells(8, 5).Address(False, False) & " Cell Value"
    End With
    
    ' Textbox property setting (txt)
    IndexA_SetTextBoxProperties "txtIndexALow_max_value", 10, 20, 50, 15, 1
    IndexA_SetTextBoxProperties "txtIndexAIntermediate_max_value", 10, 50, 50, 15, 3

    ' ComboBoxes property setting (cmb)
    For i = 1 To 2
        IndexA_SetComboBoxProperties "cmbIndexAsign" & i, 60, 20 + (i - 1) * 30, 50, 15, 2 + (i - 1) * 3
    Next i

    ' commandbutton property setting (btn)
    IndexA_SetCommandButtonProperties "btnIndexAinput", "IndexInput", 10, 75, 100, 20
End Sub

' Label property setting functions
Private Sub IndexA_SetLabelProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, caption As String)
    With Me.Controls(controlName)
        .caption = caption
        .Font.Size = 9
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End Sub

' Textbox property setting functions
Private Sub IndexA_SetTextBoxProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, tabIndex As Integer)
    With Me.Controls(controlName)
        .left = left
        .top = top
        .width = width
        .height = height
        .tabIndex = tabIndex
        .IMEMode = 3
        .TextAlign = 3
    End With
End Sub

' ComboBox property setting functions
Private Sub IndexA_SetComboBoxProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, tabIndex As Integer)
    With Me.Controls(controlName)
        .left = left
        .top = top
        .width = width
        .height = height
        .tabIndex = tabIndex
        .TextAlign = 1
    End With
End Sub

' Command button property setting functions
Private Sub IndexA_SetCommandButtonProperties(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer)
    With Me.Controls(controlName)
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End Sub


Sub IndexB_AddLabelsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newLabel As Control

    ' Create list of control names
    controlNames = Array("IndexBLow_max_value")

    ' Add labels beginning with lbl
    For i = LBound(controlNames) To UBound(controlNames)
        Set newLabel = MultiPage1.Pages(pageIndex).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
    Next i
    
    controlNames = Array("IndexBsign1")

    
    For i = LBound(controlNames) To UBound(controlNames)
        ' lblcmb and lbltype have different name patterns, so add them separately
        Set newLabel = MultiPage1.Pages(pageIndex).Controls.Add("Forms.Label.1", "lbl" & controlNames(i))
        ' You can set other properties of newLabel here
        Next i
End Sub
Sub IndexB_AddTextBoxesToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newText As Control

    ' Create list of control names
    controlNames = Array("IndexBLow_max_value")

    ' Add textboxes beginning with txt
    For i = LBound(controlNames) To UBound(controlNames)
        Set newText = MultiPage1.Pages(pageIndex).Controls.Add("Forms.TextBox.1", "txt" & controlNames(i))
        ' You can set other properties of newText here
    Next i
End Sub
Sub IndexB_AddComboBoxesToPage(pageIndex As Long)
    Dim controlPrefixes As Variant, prefix As Variant
    Dim newCombo As Control

    ' Create list of control names
    controlPrefixes = Array("IndexBsign1")

    ' Add comboboxes beginning with cmb
    For Each prefix In controlPrefixes
        Set newCombo = MultiPage1.Pages(pageIndex).Controls.Add("Forms.ComboBox.1", "cmb" & prefix)
        ' You can set other properties of newCombo here
    Next prefix
End Sub
Sub IndexB_AddButtonsToPage(pageIndex As Long)
    Dim controlNames As Variant
    Dim i As Long
    Dim newButton As Control, newText As Control

    ' Create list of control names
    controlNames = Array("IndexBinput")

    ' Add CommandButton beginning with btn
    For i = LBound(controlNames) To UBound(controlNames)
        Set newButton = MultiPage1.Pages(pageIndex).Controls.Add("Forms.CommandButton.1", "btn" & controlNames(i))
        ' You can set other properties of newText here
        Dim buttonEvent As New Class2
        
        ' Event Handler Assignment
        With buttonEvent
            Set .btnIndexBinput = newButton
        End With
        
        ButtonHandler.Add buttonEvent
        
    Next i
End Sub
Private Sub IndexB_SetControlsProperties(pageIndex As Long)
    Dim wsIndexB As Worksheet
    Set wsIndexB = Worksheets("IndexB")
    
    With wsIndexB
        ' Label property setting (lbl)
        IndexB_SetLabelProperties "lblIndexBsign1", 60, 10, 180, 20, "Please select less than or or below from the list"
        IndexB_SetLabelProperties "lblIndexBLow_max_value", 10, 10, 50, 20, .Cells(6, 5).Address(False, False) & " Cell Value"
    End With

    ' Textbox property setting (txt)
    IndexB_SetTextBoxProperties "txtIndexBLow_max_value", 10, 20, 50, 15, 1

    ' ComboBoxes property setting (cmb)
    IndexB_SetComboBoxProperties "cmbIndexBsign1", 60, 20, 50, 15, 2

    ' commandbutton property setting (btn)
    IndexB_SetCommandButtonProperties "btnIndexBinput", "IndexInput", 10, 75, 100, 20
End Sub

' Helper functions for setting various properties
Private Sub IndexB_SetLabelProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, caption As String)
    With Me.Controls(controlName)
        .caption = caption
        .Font.Size = 9
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End Sub

Private Sub IndexB_SetTextBoxProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, tabIndex As Integer)
    With Me.Controls(controlName)
        .left = left
        .top = top
        .width = width
        .height = height
        .tabIndex = tabIndex
        .IMEMode = 3
        .TextAlign = 3
    End With
End Sub

Private Sub IndexB_SetComboBoxProperties(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer, tabIndex As Integer)
    With Me.Controls(controlName)
        .left = left
        .top = top
        .width = width
        .height = height
        .tabIndex = tabIndex
        .TextAlign = 1
    End With
End Sub

Private Sub IndexB_SetCommandButtonProperties(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer)
    With Me.Controls(controlName)
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End Sub


