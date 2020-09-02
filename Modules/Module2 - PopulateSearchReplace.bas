Attribute VB_Name = "Module2"
Public Sub PopulateSearchReplace()

Dim name As Variant

PrepareForm

SearchForm.ChooseRecordComboBox.Clear
SearchForm.DisplayCategoryComboBox.Clear

For i = 1 To WorksheetFunction.CountA(Columns("A:A")) - 1
        SearchForm.ChooseRecordComboBox.AddItem ActiveCell.Offset(i, 0)
Next i

For i = 1 To WorksheetFunction.CountA(Rows("1:1"))
        SearchForm.DisplayCategoryComboBox.AddItem ActiveCell.Offset(0, i - 1)
Next i

SearchForm.ChooseRecordComboBox.Text = Range("A2")
SearchForm.DisplayCategoryComboBox.Text = Range("A1")

PrepareForm

End Sub
