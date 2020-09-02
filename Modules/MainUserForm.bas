VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainUserForm 
   Caption         =   "Ultimate Organizer"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4620
   OleObjectBlob   =   "MainUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

AddCategoryForm.Show

End Sub

Private Sub CommandButton2_Click()

AddRecordForm.Show

End Sub

Private Sub CommandButton3_Click()

For i = 1 To WorksheetFunction.CountA(Rows("1:1"))
        DeleteCategoryForm.DeleteCategoryComboBox.AddItem ActiveCell.Offset(0, i - 1)
Next i

DeleteCategoryForm.DeleteCategoryComboBox.Text = Range("A1")

PrepareForm

DeleteCategoryForm.Show

End Sub

Private Sub CommandButton4_Click()

For i = 1 To WorksheetFunction.CountA(Columns("A:A")) - 1
        DeleteRecordForm.DeleteRecordComboBox.AddItem ActiveCell.Offset(i, 0)
Next i

DeleteRecordForm.DeleteRecordComboBox.Text = Range("A2")

PrepareForm

DeleteRecordForm.Show

End Sub

Private Sub CommandButton5_Click()

PrepareForm

For i = 1 To WorksheetFunction.CountA(Columns("A:A")) - 1
    SearchForm.ChooseRecordComboBox.AddItem ActiveCell.Offset(i, 0)
Next i

For i = 1 To WorksheetFunction.CountA(Rows("1:1"))
    SearchForm.DisplayCategoryComboBox.AddItem ActiveCell.Offset(0, i - 1)
Next i

SearchForm.ChooseRecordComboBox.Text = Range("A2")
SearchForm.DisplayCategoryComboBox.Text = Range("A1")

PrepareForm

SearchForm.Show

End Sub

Private Sub CommandButton6_Click()

Unload MainUserForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
