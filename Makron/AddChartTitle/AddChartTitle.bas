Attribute VB_Name = "AddChartTitle"
'Lägger till en rubrik till markerat diagram.
'markera diagram, kör makro och skriv texten i rutan som visas.
Sub AddChartTitle()
Dim i As Variant
i = InputBox("Please enter your chart title", "Chart Title")
On Error GoTo Last
ActiveChart.SetElement (msoElementChartTitleAboveChart)
ActiveChart.ChartTitle.Text = i
Last:
Exit Sub
End Sub
