Attribute VB_Name = "Module2"
Option Base 1
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

Dim yp(11, 1) As Variant
yp(1, 1) = 1
For i = 1 To 10
yp(i + 1, 1) = 1 + yp(i, 1)
Next
Dim rng As Range
'Set rng = Range("c1")

'rng.Cells(1, 1).Resize(11, 1) = yp


Set rng = Range("a1")

'rng.Resize(UBound(yp, 1)).Value = yp
rng.Resize(10).Value = yp


End Sub

Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
Dim cht As Object
' Macro6 Macro
'
    Range("b1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Select
'
    'ActiveSheet.ChartObjects("Chart 11").Activate
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!$A$1:$A$10"
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    Application.Left = 506.8
    Application.Top = 22.6
    Range("A1:B10").Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$B$10")
    Application.Left = 440.2
    Application.Top = 43
    ActiveChart.ChartTitle.Select
    Selection.Delete
    Application.Left = 401.2
    Application.Top = 23.8
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    'ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Application.Left = 569.2
    Application.Top = 15.4
    ActiveChart.Axes(xlValue).AxisTitle.Select
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Application.Left = 359.8
    Application.Top = 11.8
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "x"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "x"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""predictions"""
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet1!$A$1:$A$10"
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet1!$C$1:$C$10"
    Application.Left = 743.8
    Application.Top = 34.6
    ActiveChart.FullSeriesCollection(2).Select
    Application.Left = 226.6
    Application.Top = 25.6
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    Selection.MarkerStyle = -4142
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    Application.Left = 370
    Application.Top = 47.2
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Name = "=""series data"""
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    Range("A1:B10").Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$B$10")
    ActiveChart.ChartTitle.Select
    Selection.Delete
    Application.Left = 460
    Application.Top = 46.6
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleHorizontal)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "y"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "x"
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet1!$A$1:$A$10"
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet1!$C$1:$C$10"
    ActiveChart.FullSeriesCollection(2).Select
    Application.Left = 143.2
    Application.Top = 1
    ActiveChart.FullSeriesCollection(2).Smooth = True
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    Selection.Format.Line.Visible = msoFalse
    Selection.MarkerStyle = -4142
    Application.CommandBars("Format Object").Visible = False
    ActiveChart.SetElement (msoElementLegendRight)
    Application.Left = -4.4
    Application.Top = -4.4
    Application.Width = 1162.8
    Application.Height = 628.8
    ActiveChart.Legend.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Name = "=""Experimental data"""
    ActiveChart.FullSeriesCollection(2).Name = "=""model predictions"""
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

    ActiveSheet.ChartObjects("Chart 42").Activate
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)

End Sub
Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro
'

'
    ActiveSheet.ChartObjects("Chart 42").Activate
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    ActiveSheet.ChartObjects("Chart 42").Activate
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleHorizontal)
End Sub
Sub Macro12()
Attribute Macro12.VB_ProcData.VB_Invoke_Func = " \n14"
Dim u As Integer
u = 87
If IsNumeric(87) = False Then
    MsgBox ("Please use excel syntax")
End If

End Sub
