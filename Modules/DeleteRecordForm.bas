VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteRecordForm 
   Caption         =   "Delete Record"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "DeleteRecordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim i As Integer
Dim mainPage As Worksheet
Dim delValue As Variant

Set mainPage = Sheet1

mainPage.Range("A1").Select

delValue = DeleteRecordForm.DeleteRecordComboBox

Ans = MsgBox("Are you sure you want to delete this record?", 1, "Confirmation")

If Ans = 1 Then
    Do
        i = i + 1
        If ActiveCell.Offset(i - 1, 0) = delValue Then
            ActiveCell.Offset(i - 1, 0).EntireRow.ClearContents: ActiveCell.Offset(i - 1, 0).EntireRow.Delete Shift:=xlUp
            Exit Do
        End If
    Loop
End If

Call Reformat

Unload DeleteRecordForm

PrepareForm

End Sub

Private Sub CommandButton2_Click()

Unload DeleteRecordForm

PrepareForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
