Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub RunForm()

PrepareForm

MainUserForm.Show

End Sub

Sub PrepareForm()

Dim nr As Integer, nc As Integer, i As Integer
Dim categories() As Variant
Dim mainPage As Worksheet

Set mainPage = Sheet1

If mainPage.Range("A1") = "" Then
    DefaultCategoryForm.Show
End If

nr = WorksheetFunction.CountA(Columns("A:A")) - 1
nc = WorksheetFunction.CountA(Rows("1:1"))

mainPage.Range("A1").Select

ReDim categories(nc) As Variant

For i = 1 To nc
    categories(i) = ActiveCell.Offset(0, i - 1)
Next i

With AddRecordForm
    .Input1.Visible = False
    .Input2.Visible = False
    .Input3.Visible = False
    .Input4.Visible = False
    .Input5.Visible = False
    .Input6.Visible = False
    .Input7.Visible = False
    .Input8.Visible = False
    .Input9.Visible = False
    .Input10.Visible = False
    .Input11.Visible = False
    .Input12.Visible = False
End With

If nc >= 1 Then
    AddRecordForm.Label1 = categories(1)
    AddRecordForm.Input1.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label1 = ""
End If
If nc >= 2 Then
    AddRecordForm.Label2 = categories(2)
    AddRecordForm.Input2.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label2 = ""
End If
If nc >= 3 Then
    AddRecordForm.Label3 = categories(3)
    AddRecordForm.Input3.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label3 = ""
End If
If nc >= 4 Then
    AddRecordForm.Label4 = categories(4)
    AddRecordForm.Input4.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label4 = ""
End If
If nc >= 5 Then
    AddRecordForm.Label5 = categories(5)
    AddRecordForm.Input5.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label5 = ""
End If
If nc >= 6 Then
    AddRecordForm.Label6 = categories(6)
    AddRecordForm.Input6.Visible = True
    AddRecordForm.Width = 250
Else:
    AddRecordForm.Label6 = ""
End If
If nc >= 7 Then
    AddRecordForm.Label7 = categories(7)
    AddRecordForm.Input7.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label7 = ""
End If
If nc >= 8 Then
    AddRecordForm.Label8 = categories(8)
    AddRecordForm.Input8.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label8 = ""
End If
If nc >= 9 Then
    AddRecordForm.Label9 = categories(9)
    AddRecordForm.Input9.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label9 = ""
End If
If nc >= 10 Then
    AddRecordForm.Label10 = categories(10)
    AddRecordForm.Input10.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label10 = ""
End If
If nc >= 11 Then
    AddRecordForm.Label11 = categories(11)
    AddRecordForm.Input11.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label11 = ""
End If
If nc >= 12 Then
    AddRecordForm.Label12 = categories(12)
    AddRecordForm.Input12.Visible = True
    AddRecordForm.Width = 500
Else:
    AddRecordForm.Label12 = ""
End If

End Sub
