VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCategoryForm 
   Caption         =   "Add Category"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "AddCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Dim i As Integer
Dim mainPage As Worksheet

Set mainPage = Sheet1
mainPage.Range("A1").Select

Do
    i = i + 1
    If IsEmpty(ActiveCell.Offset(0, i - 1)) Then
        Exit Do
    ElseIf ActiveCell.Offset(0, i - 1) = NewCategoryName Then
        MsgBox ("Duplicate entry, try again.")
        GoTo Reset:
    ElseIf i = 12 Then
        MsgBox ("Maximum of 12 categories reached.")
        GoTo Reset:
    ElseIf AddCategoryForm.NewCategoryName = "" Then
        MsgBox ("Category cannot be empty.")
        GoTo Reset:
    End If
        
Loop

ActiveCell.Offset(0, i - 1) = AddCategoryForm.NewCategoryName
ActiveCell.Offset(0, i - 1).Font.Bold = True

Reset:

AddCategoryForm.NewCategoryName.Text = ""

Call Reformat

PrepareForm

End Sub

Private Sub CommandButton2_Click()

Unload AddCategoryForm

PrepareForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
