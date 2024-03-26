Attribute VB_Name = "InfogaDiagramTitel"
'First of all, you need to select your chart and the run this code.
'You will get an input box to enter chart title.
Sub AddDiagramTitle()
Dim i As Variant
i = InputBox("Please enter your chart title", "Chart Title")
On Error GoTo Last
ActiveChart.SetElement (msoElementChartTitleAboveChart)
ActiveChart.ChartTitle.Text = i
Last:
Exit Sub
End Sub
