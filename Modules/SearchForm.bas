VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "Search and Replace Tool"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "SearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChooseRecordComboBox_Change()

End Sub

Private Sub CommandButton1_Click()

Dim i As Variant, j As Variant
Dim Ans As Variant, Inp As Variant

Set mainPage = Sheet1

For i = 1 To WorksheetFunction.CountA(Rows("1:1"))
    If ActiveCell.Offset(0, i - 1) = SearchForm.DisplayCategoryComboBox Then
        SearchForm.Label4 = SearchForm.DisplayCategoryComboBox
        For j = 1 To WorksheetFunction.CountA(Columns("A:A")) - 1
            If ActiveCell.Offset(j, 0) = SearchForm.ChooseRecordComboBox Then
                If ActiveCell.Offset(j, i - 1) = "" Then
                    Ans = MsgBox("No record found, would you like to add one?", vbYesNo)
                        If Ans = vbYes Then
                            Inp = InputBox("Please enter new record: ")
                            ActiveCell.Offset(j, i - 1) = Inp
                            SearchForm.SearchResult = ""
                            Call Reformat
                            PrepareForm
                        Else:
                            SearchForm.SearchResult = ""
                        End If
                Else:
                    SearchForm.SearchResult = ActiveCell.Offset(j, i - 1)
                End If
            End If
        Next j
    End If
Next i

End Sub

Private Sub CommandButton2_Click()

Dim i As Integer, Ans As Variant, j As Integer

If SearchForm.SearchResult = "" Then
    MsgBox ("Cannot be blank.")
Else:
    Ans = MsgBox("Are you sure you want to replace this data?", 1, "Confirmation")
    If Ans = 1 Then
        For i = 1 To WorksheetFunction.CountA(Rows("1:1"))
            If ActiveCell.Offset(0, i - 1) = SearchForm.DisplayCategoryComboBox Then
                For j = 1 To WorksheetFunction.CountA(Columns("A:A")) - 1
                    If ActiveCell.Offset(j, 0) = SearchForm.ChooseRecordComboBox Then
                        ActiveCell.Offset(j, i - 1) = SearchResult
                    End If
                Next j
            End If
        Next i
    End If
End If

SearchForm.SearchResult = ""

Call PopulateSearchReplace
Call Reformat

PrepareForm

End Sub

Private Sub CommandButton3_Click()

Unload SearchForm

PrepareForm

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
