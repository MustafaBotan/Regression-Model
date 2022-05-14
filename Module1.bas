Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub RunForm()
UserForm1.Show
End Sub

Sub test()
Dim tWB As Workbook
Dim UserXRange As Range, UserYRange As Range
Dim rng As Range
Dim x, i As Integer
Dim y, j As Integer

'PLACE YOUR ADDITIONAL DIM STATEMENTS IN THIS REGION
Dim Xt, X1, XtX, XtXinv, XtY, Beta As Variant

Set tWB = ThisWorkbook
tWB.Activate
'THE FOLLOWING TWO LINES JUST SETS A DEFAULT RANGE IN THE INPUT BOXES, THAT'S ALL
Set UserXRange = Application.InputBox("X Input Range", "X Input", "Sheet2!$A$1:$A$10", Type:=8)
Set UserYRange = Application.InputBox("Y Input Range", "Y Input", "Sheet2!$B$1:$B$10", Type:=8)

'PLACE THE MAIN BULK OF YOUR CODE IN THIS REGION!
ReDim x(UserXRange.Rows.Count, UserXRange.Columns.Count) As Double
ReDim y(UserYRange.Rows.Count, UserYRange.Columns.Count) As Double

x = UserXRange
y = UserYRange

ReDim X1(UBound(x, 1), UBound(x, 2) + 1) As Double

For i = 1 To UBound(x, 1)
    X1(i, 1) = 1
Next

For i = 1 To UBound(x, 1)
For j = 1 To UBound(x, 2)
    X1(i, j + 1) = x(i, j)
Next
Next


Xt = WorksheetFunction.Transpose(X1)
XtX = WorksheetFunction.MMult(Xt, X1)
XtY = WorksheetFunction.MMult(Xt, y)
XtXinv = WorksheetFunction.MInverse(XtX)
Beta = WorksheetFunction.MMult(XtXinv, XtY)
Sheets("sheet2").Select

'Set Rng = Range("e41")

'Rng.Cells(1, 1).Resize(UBound(Beta, 1), UBound(Beta, 2)) = Beta

For i = 1 To UBound(Beta, 1)
    Beta(i, 1) = WorksheetFunction.Round(Beta(i, 1), 3)
Next

MsgBox ("Model is: y = " & Beta(1, 1) & " + " & Beta(2, 1) & "*x^2 + " & Beta(3, 1) & "*sqrt(x) + " & Beta(4, 1) & "*1/x + " & Beta(5, 1) & "*ln(x)")


'adjusted r squired




End Sub

Sub mat()

Dim A(2, 2) As Variant, B(2, 3) As Variant
Dim c As Variant
Dim fx As String
Dim y, d As Integer

y = Cells(4, 1)







End Sub
