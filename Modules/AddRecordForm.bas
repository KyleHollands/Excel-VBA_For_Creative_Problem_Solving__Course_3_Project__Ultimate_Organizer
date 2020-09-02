VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddRecordForm 
   Caption         =   "Add Record"
   ClientHeight    =   5475
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   9300.001
   OleObjectBlob   =   "AddRecordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddButton_Click()

Dim nr As Integer
Dim mainPage As Worksheet

PrepareForm

Set mainPage = Sheet1

nr = WorksheetFunction.CountA(Columns("A:A")) - 1

mainPage.Range("A1").Select

ActiveCell.Offset(nr + 1, 0).Select

ActiveCell.Offset(0, 0) = AddRecordForm.Input1
ActiveCell.Offset(0, 1) = AddRecordForm.Input2
ActiveCell.Offset(0, 2) = AddRecordForm.Input3
ActiveCell.Offset(0, 3) = AddRecordForm.Input4
ActiveCell.Offset(0, 4) = AddRecordForm.Input5
ActiveCell.Offset(0, 5) = AddRecordForm.Input6
ActiveCell.Offset(0, 6) = AddRecordForm.Input7
ActiveCell.Offset(0, 7) = AddRecordForm.Input8
ActiveCell.Offset(0, 8) = AddRecordForm.Input9
ActiveCell.Offset(0, 9) = AddRecordForm.Input10
ActiveCell.Offset(0, 10) = AddRecordForm.Input11
ActiveCell.Offset(0, 11) = AddRecordForm.Input12

AddRecordForm.Input1 = ""
AddRecordForm.Input2 = ""
AddRecordForm.Input3 = ""
AddRecordForm.Input4 = ""
AddRecordForm.Input5 = ""
AddRecordForm.Input6 = ""
AddRecordForm.Input7 = ""
AddRecordForm.Input8 = ""
AddRecordForm.Input9 = ""
AddRecordForm.Input10 = ""
AddRecordForm.Input11 = ""
AddRecordForm.Input12 = ""

Call Reformat

PrepareForm

End Sub

Private Sub QuitButton_Click()

Unload AddRecordForm

PrepareForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
