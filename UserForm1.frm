VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Regression Toolbox"
   ClientHeight    =   5172
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   7236
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub GoButton_Click()
Dim tWB As Workbook
Dim UserXRange As Range, UserYRange As Range
Dim rng As Range
Dim x, i As Integer
Dim y, j As Integer
Dim Parameters As Integer
Dim v1, v2, v3, v4 As Double

'PLACE YOUR ADDITIONAL DIM STATEMENTS IN THIS REGION
Dim Xt, X1, XtX, XtXinv, XtY, Beta As Variant
Dim sqRes, yp, fxn As Variant
Dim SST  As Double
Dim SSE As Double
Dim adjR As Double
Dim yAvg As Double
Dim ans As Integer
Dim messageB As String

Set tWB = ThisWorkbook
tWB.Activate

'PLACE THE MAIN BULK OF YOUR CODE IN THIS REGION!
If Not fxn1 = "" Then
    Parameters = 1 + Parameters
End If
If Not fxn2 = "" Then
    Parameters = 1 + Parameters
End If
If Not fxn3 = "" Then
    Parameters = 1 + Parameters
End If
If Not fxn4 = "" Then
    Parameters = 1 + Parameters
End If

If Parameters = 0 Then
    MsgBox ("Please input variables")
    Exit Sub
End If

Cells(14, 20) = "=" & Replace(fxn1, "x", 1)
If IsNumeric(Cells(14, 20)) = False & Not fxn1 = "" Or IsError(Cells(14, 20)) = True Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If

Cells(14, 20) = "=" & Replace(fxn2, "x", 1)
If IsNumeric(Cells(14, 1)) = False & Not fxn2 = "" Or IsError(Cells(14, 20)) = True Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If
Cells(14, 20) = "=" & Replace(fxn3, "x", 1)
If IsNumeric(Cells(14, 20)) = False & Not fxn3 = "" Or IsError(Cells(14, 20)) = True Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If
Cells(14, 20) = "=" & Replace(fxn4, "x", 1)
If IsNumeric(Cells(14, 20)) = False & Not fxn4 = "" Or IsError(Cells(14, 20)) = True Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If

If Not InStr(1, fxn1, "x", vbTextCompare) <> 0 And Not fxn1 = "" Then
   MsgBox ("Please use excel syntax")
   Exit Sub
End If
If Not InStr(1, fxn2, "x", vbTextCompare) <> 0 And Not fxn2 = "" Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If
If Not InStr(1, fxn3, "x", vbTextCompare) <> 0 And Not fxn3 = "" Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If
If Not InStr(1, fxn4, "x", vbTextCompare) <> 0 And Not fxn4 = "" Then
    MsgBox ("Please use excel syntax")
    Exit Sub
End If


Set UserXRange = Application.InputBox("X Input Range", "X Input", "Sheet1!$a$1:$a$10", Type:=8)
Set UserYRange = Application.InputBox("Y Input Range", "Y Input", "Sheet1!$b$1:$b$10", Type:=8)

ReDim x(UserXRange.Rows.Count, 1) As Double
ReDim y(UserYRange.Rows.Count, 1) As Double

x = UserXRange
y = UserYRange

ReDim X1(UBound(x, 1), Parameters + 1) As Double

For i = 1 To UBound(x, 1)
    X1(i, 1) = 1
Next

ReDim fxn(Parameters) As Variant
i = 1
    If Not fxn1 = "" Then
        fxn(i) = fxn1
        i = 1 + i
    End If
    If Not fxn2 = "" Then
        fxn(i) = fxn2
        i = 1 + i
    End If
    If Not fxn3 = "" Then
        fxn(i) = fxn3
        i = 1 + i
    End If
    If Not fxn4 = "" Then
        fxn(i) = fxn4
        i = 1 + i
    End If


For i = 1 To UBound(x, 1)
    If Parameters = 1 Then
        Cells(i, 10).Formula = "=" & Replace(fxn(1), "x", x(i, 1))
        X1(i, 2) = Cells(i, 10)
    End If
    If Parameters = 2 Then
        Cells(i, 10).Formula = "=" & Replace(fxn(1), "x", x(i, 1))
        X1(i, 2) = Cells(i, 10)
        Cells(i, 11).Formula = "=" & Replace(fxn(2), "x", x(i, 1))
        X1(i, 3) = Cells(i, 11)
    End If
    If Parameters = 3 Then
        Cells(i, 10).Formula = "=" & Replace(fxn(1), "x", x(i, 1))
        X1(i, 2) = Cells(i, 10)
        Cells(i, 11).Formula = "=" & Replace(fxn(2), "x", x(i, 1))
        X1(i, 3) = Cells(i, 11)
        Cells(i, 12).Formula = "=" & Replace(fxn(3), "x", x(i, 1))
        X1(i, 4) = Cells(i, 12)
    End If
    If Parameters = 4 Then
        Cells(i, 10).Formula = "=" & Replace(fxn(1), "x", x(i, 1))
        X1(i, 2) = Cells(i, 10)
        Cells(i, 11).Formula = "=" & Replace(fxn(2), "x", x(i, 1))
        X1(i, 3) = Cells(i, 11)
        Cells(i, 12).Formula = "=" & Replace(fxn(3), "x", x(i, 1))
        X1(i, 4) = Cells(i, 12)
        Cells(i, 13).Formula = "=" & Replace(fxn(4), "x", x(i, 1))
        X1(i, 5) = Cells(i, 13)
    End If
Next



Xt = WorksheetFunction.Transpose(X1)
XtX = WorksheetFunction.MMult(Xt, X1)
XtY = WorksheetFunction.MMult(Xt, y)
XtXinv = WorksheetFunction.MInverse(XtX)
Beta = WorksheetFunction.MMult(XtXinv, XtY)


For i = 1 To UBound(Beta, 1)
    Beta(i, 1) = WorksheetFunction.Round(Beta(i, 1), 3)
    Beta(i, 1) = Replace(Beta(i, 1), ",", ".")
Next

messageB = "Model is: y = " & Beta(1, 1)
'MsgBox ("Model is: y = " & Beta(1, 1) & " + " & Beta(2, 1) & "*" & fxn(1) & " + " & Beta(3, 1) & "*" & fxn(2) & " + " & Beta(4, 1) & "*" & fxn(3) & " + " & Beta(5, 1) & "*" & fxn(4))
For i = 1 To Parameters
    messageB = messageB & " + " & Beta(1 + i, 1) & "*" & fxn(i)
Next

MsgBox messageB

'adjusted r squired

For i = 1 To UBound(x, 1)
    If Parameters = 1 Then
        v1 = Cells(i, 10)
        Cells(i, 15).Formula = "=" & Beta(1, 1) & "+" & Beta(2, 1) & "*" & Replace(v1, ",", ".")
    End If
    If Parameters = 2 Then
       v1 = Cells(i, 10)
       v2 = Cells(i, 11)
        Cells(i, 15).Formula = "=" & Beta(1, 1) & "+" & Beta(2, 1) & "*" & Replace(v1, ",", ".") & " + " & Beta(3, 1) & " * " & Replace(v2, ",", ".")
    End If
    If Parameters = 3 Then
        v1 = Cells(i, 10)
        v2 = Cells(i, 11)
        v3 = Cells(i, 12)
        Cells(i, 15).Formula = "=" & Beta(1, 1) & "+" & Beta(2, 1) & "*" & Replace(v1, ",", ".") & " + " & Beta(3, 1) & " * " & Replace(v2, ",", ".") & " + " & Beta(4, 1) & " * " & Replace(v3, ",", ".")
    End If
    If Parameters = 4 Then
        v1 = Cells(i, 10)
        v2 = Cells(i, 11)
        v3 = Cells(i, 12)
        v4 = Cells(i, 13)
        Cells(i, 15).Formula = "=" & Beta(1, 1) & "+" & Beta(2, 1) & "*" & Replace(v1, ",", ".") & " + " & Beta(3, 1) & " * " & Replace(v2, ",", ".") & " + " & Beta(4, 1) & " * " & Replace(v3, ",", ".") & " + " & Beta(5, 1) & " * " & Replace(v4, ",", ".")
    End If
Next

'residual
ReDim sqRes(UBound(y, 1)) As Double
ReDim yp(UBound(y, 1)) As Double

For i = 1 To UBound(y, 1)
    yp(i) = Cells(i, 15)
Next

For i = 1 To UBound(y, 1)
    sqRes(i) = (y(i, 1) - yp(i)) ^ 2
Next

For i = 1 To UBound(sqRes, 1)
SSE = sqRes(i) + SSE
Next

For i = 1 To UBound(y, 1)
    yAvg = y(i, 1) + yAvg
Next
yAvg = yAvg / UBound(y, 1)

For i = 1 To UBound(y, 1)
    SST = SST + (y(i, 1) - yAvg) ^ 2
Next

adjR = 1 - (SSE / (UBound(y, 1) - Parameters)) / (SST / (UBound(y, 1) - 1))
MsgBox ("Adusted R^2: " & adjR)


ans = MsgBox("Would you like to plot the data?", vbYesNo)
If ans = 6 Then
    Call Plotting(x, y, yp)
End If

End Sub

Private Sub QuitButton_Click()
Unload Me
End Sub

Sub Plotting(x, y, yp As Variant)
Dim rng As Range
Dim i As Integer
Dim cht As Object
For i = 1 To UBound(yp, 1)
    Cells(i, 3) = yp(i)
Next
    
    Range("a1:b" & UBound(y, 1)).Select
    Set rng = Selection
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.ChartTitle.Delete
    
    
    
    'adds new series to chart
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet1!$A$1:$A$" & UBound(yp, 1)
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet1!$C$1:$C$" & UBound(yp, 1)
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).Smooth = True
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    Selection.Format.Line.Visible = msoFalse
    Selection.MarkerStyle = -4142
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.Legend.Select
    ActiveChart.FullSeriesCollection(1).Name = "=""Experimental data"""
    ActiveChart.FullSeriesCollection(2).Name = "=""model predictions"""
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    
    
    

    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleHorizontal)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "y"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "x"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "x"
    
    
    
   
End Sub
