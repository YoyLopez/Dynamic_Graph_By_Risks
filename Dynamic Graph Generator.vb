Sub CrearGraficoDinamicoConLineaLimite()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim series As series
    Dim i As Long

    Set ws = Worksheets("Hoja1")
    Set tbl = ws.ListObjects("Nombre_tabla")

    With tbl.ListColumns("VENCIDO").DataBodyRange
        .Formula = "=IF([@[TIEMPO]]>[@LÍMITE],[@[TIEMPO]],"""")"
    End With

    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=500, Top:=50, Height:=300)
    Set chart = chartObj.chart
    chart.ChartType = xlColumnClustered

    With chart
        .SetSourceData Source:=tbl.ListColumns("TIEMPO").DataBodyRange
        Set series = .SeriesCollection(1)
        series.XValues = tbl.ListColumns("EVENTO").DataBodyRange
        series.Name = "TIEMPO"
        series.Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
    End With

    With chart.SeriesCollection.NewSeries
        .Name = "VENCIDO"
        .XValues = tbl.ListColumns("EVENTO").DataBodyRange
        .Values = tbl.ListColumns("VENCIDO").DataBodyRange
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End With

    With chart.SeriesCollection.NewSeries
        .Name = "LÍMITE"
        .XValues = tbl.ListColumns("EVENTO").DataBodyRange
        .Values = tbl.ListColumns("LÍMITE").DataBodyRange
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = 2
    End With

    With chart
        .HasTitle = True
        .ChartTitle.Text = "Visualización de Riesgo"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Evento"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tiempo(Días)"
    End With
End Sub

