VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DefaultCategoryForm 
   Caption         =   "Enter Default Category"
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DefaultCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DefaultCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim mainPage As Worksheet

Set mainPage = Sheet1

If DefaultCategoryForm.DefaultCategoryName.Text = "" Then
    MsgBox ("Cannot be blank.")
    Unload DefaultCategoryForm
    DefaultCategoryForm.Show
Else:
    mainPage.Range("A1") = DefaultCategoryForm.DefaultCategoryName.Text
    mainPage.Range("A1").Font.Bold = True
End If

Call Reformat

Unload DefaultCategoryForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
