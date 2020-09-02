VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteCategoryForm 
   Caption         =   "Delete Category"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "DeleteCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteCategoryForm"
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

delValue = DeleteCategoryForm.DeleteCategoryComboBox

Ans = MsgBox("Are you sure you want to delete this category?", 1, "Confirmation")

If Ans = 1 Then
    Do
        i = i + 1
        If ActiveCell.Offset(0, i - 1) = delValue Then
            ActiveCell.Offset(0, i - 1).EntireColumn.ClearContents: ActiveCell.Offset(0, i - 1).EntireColumn.Delete Shift:=xlLeft
            Exit Do
        End If
    Loop
End If

Call Reformat

Unload DeleteCategoryForm

PrepareForm

End Sub

Private Sub CommandButton2_Click()

Unload DeleteCategoryForm

PrepareForm

End Sub

Private Sub DeleteCategoryComboBox_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
