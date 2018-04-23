Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web


Namespace T449201
    Public Class DiagramCreationHelper
        Public Shared Sub CreatePieChart(ByVal wbook As IWorkbook)

            Dim worksheet As DevExpress.Spreadsheet.Worksheet = SetActiveWorksheet(wbook, "Range1")

            ' Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.PieExploded, worksheet("B2:C7"))

            ' Display the chart title
            chart.Title.Visible = True
            chart.Title.SetReference(worksheet("B1"))

            chart.TopLeftCell = worksheet.Cells("E2")
            chart.BottomRightCell = worksheet.Cells("K15")

            ' Set the chart style
            chart.Style = ChartStyle.ColorGradient

            ' Hide the legend
            chart.Legend.Visible = False

            ' Rotate the pie chart view
            chart.Views(0).FirstSliceAngle = 100

            ' Display data labels
            Dim dataLabels As DataLabelOptions = chart.Views(0).DataLabels
            dataLabels.ShowCategoryName = True
            dataLabels.ShowPercent = True
            dataLabels.Separator = ControlChars.Lf
        End Sub
        Public Shared Sub CreateBarChart(ByVal wbook As IWorkbook)

            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range1")

            Dim chart As Chart = worksheet.Charts.Add(ChartType.BarFullStacked)
            chart.TopLeftCell = worksheet.Cells("E2")
            chart.BottomRightCell = worksheet.Cells("K15")

            ' Select chart data
            chart.SelectData(worksheet("B2:C7"), ChartDataDirection.Row)

            ' Display the chart title
            chart.Title.Visible = True
            chart.Title.SetReference(worksheet("B1"))

            ' Change legend position
            chart.Legend.Position = LegendPosition.Bottom

            ' Hide the category axis
            chart.PrimaryAxes(0).Visible = False

            ' Set major unit of the value axis
            chart.PrimaryAxes(1).MajorUnit = 0.2

        End Sub
        Public Shared Sub CreateColumnChart(ByVal wbook As IWorkbook)
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range2")
            'Create data range for a chart

            Dim chart As Chart = worksheet.Charts.Add(ChartType.ColumnClustered)
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N14")

            ' Add series
            chart.Series.Add(worksheet("D2"), worksheet("B3:B6"), worksheet("D3:D6"))
            chart.Series.Add(worksheet("F2"), worksheet("B3:B6"), worksheet("F3:F6"))

            ' Display the chart title
            chart.Title.Visible = True
            chart.Title.SetValue("Mobile OS market share")

            ' Customize the appearance and scale of the axes
            Dim axis As Axis = chart.PrimaryAxes(0)
            axis.MajorTickMarks = AxisTickMarks.None
            axis = chart.PrimaryAxes(1)
            axis.Outline.SetNoFill()
            axis.MajorTickMarks = AxisTickMarks.None
            axis.NumberFormat.FormatCode = "0%"
            axis.NumberFormat.IsSourceLinked = False
            axis.Scaling.AutoMax = False
            axis.Scaling.Max = 1
            axis.Scaling.AutoMin = False
            axis.Scaling.Min = 0

            ' Set the gap width between data series
            Dim view As ChartView = chart.Views(0)
            view.GapWidth = 75

            ' Display data labels
            view.DataLabels.ShowValue = True
            view.DataLabels.NumberFormat.FormatCode = "0%"
            view.DataLabels.NumberFormat.IsSourceLinked = False

            ' Set the chart style
            chart.Style = ChartStyle.ColorGradient
        End Sub
        Public Shared Sub CreateComplexChart(ByVal wbook As IWorkbook)

            'Create data range for a chart
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range3")

            Dim chart As Chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet("B2:D8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L15")

            ' Change the chart type of the second series
            chart.Series(1).ChangeType(ChartType.Line)
            chart.Series(1).Smooth = True

            ' Use secondary axes
            chart.Series(1).AxisGroup = AxisGroup.Secondary

            ' Specify the chart style
            chart.Style = ChartStyle.ColorGradient

            ' Set the position of the legend
            chart.Legend.Position = LegendPosition.Bottom
        End Sub
        Public Shared Sub CreateDoughnutChart(ByVal wbook As IWorkbook)

            'Create data range for a chart
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range2")

            ' Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Doughnut)
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N14")

            ' Add the data series
            chart.Series.Add(worksheet("E2"), worksheet("B3:B6"), worksheet("E3:E6"))

            ' Display the chart title
            chart.Title.Visible = True
            chart.Title.SetValue("Mobile OS market share Q4'13")

            ' Change the hole size
            chart.Views(0).HoleSize = 60

            ' Display the data labels
            chart.Views(0).DataLabels.ShowPercent = True
        End Sub
        Public Shared Sub CreatePie3dChart(ByVal wbook As IWorkbook)

            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range2")

            Dim chart As Chart = worksheet.Charts.Add(ChartType.Pie3D)
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N14")

            ' Add the data series
            chart.Series.Add(worksheet("E2"), worksheet("B3:B6"), worksheet("E3:E6"))

            ' Set the explosion value for the slice
            chart.Series(0).CustomDataPoints.Add(2).Explosion = 25

            ' Set the rotation of the  3-D chart view
            chart.View3D.YRotation = 255

            ' Set the chart style
            chart.Style = ChartStyle.ColorGradient

        End Sub
        Public Shared Sub CreateScatterChart(ByVal wbook As IWorkbook)
            'Create data range for a chart
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range4")

            Dim chart As Chart = worksheet.Charts.Add(ChartType.ScatterLineMarkers, worksheet("C2:D52"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L15")

            ' Set the marker symbol
            chart.Series(0).Marker.Symbol = MarkerStyle.Circle

            ' Set appearance and scale of the X axis
            Dim axis As Axis = chart.PrimaryAxes(0)
            axis.Scaling.AutoMax = False
            axis.Scaling.AutoMin = False
            axis.Scaling.Max = 60.0
            axis.Scaling.Min = -60.0
            axis.MajorGridlines.Visible = True

            ' Set appearance and scale of the Y axis
            axis = chart.PrimaryAxes(1)
            axis.Scaling.AutoMax = False
            axis.Scaling.AutoMin = False
            axis.Scaling.Max = 50.0
            axis.Scaling.Min = -50.0
            axis.MajorUnit = 10.0



        End Sub
        Public Shared Sub CreateStockChart(ByVal wbook As IWorkbook)
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range5")
            Dim chart As Chart = worksheet.Charts.Add(ChartType.StockOpenHighLowClose, worksheet("B2:F7"))
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N15")

            ' Display the chart title
            chart.Title.Visible = True
            chart.Title.SetValue("NASDAQ:MSFT")

            ' Hide the legend
            chart.Legend.Visible = False

            ' Set appearance and scale of the value axis
            Dim axis As Axis = chart.PrimaryAxes(1)
            axis.Scaling.AutoMax = False
            axis.Scaling.Max = 40.5
            axis.Scaling.AutoMin = False
            axis.Scaling.Min = 38.5
            axis.MajorUnit = 0.25

            ' Format the axis labels
            axis.NumberFormat.FormatCode = "#0.00"
            axis.NumberFormat.IsSourceLinked = False

            ' Display the axis title
            axis.Title.Visible = True
            axis.Title.SetValue("Price in USD")


        End Sub
        Public Shared Sub CreateBubbleChart(ByVal wbook As IWorkbook)
            Dim worksheet As Worksheet = SetActiveWorksheet(wbook, "Range6")

            ' Create a chart and specify its location
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Bubble3D)
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L15")

            Dim s1 As Series = chart.Series.Add(worksheet("A3"), worksheet("C3:C7"), worksheet("D3:D7"))
            s1.BubbleSize = ChartData.FromRange(worksheet("E3:E7"))
            Dim s2 As Series = chart.Series.Add(worksheet("A9"), worksheet("C9:C13"), worksheet("D9:D13"))
            s2.BubbleSize = ChartData.FromRange(worksheet("E9:E13"))

            ' Set the chart style
            chart.Style = ChartStyle.ColorGradient
            ' Set the bubble size 1.5x relative to the default setting.
            chart.Views(0).BubbleScale = 150

            ' Hide the legend
            chart.Legend.Visible = False

            ' Display data labels
            Dim dataLabels As DataLabelOptions = chart.Views(0).DataLabels
            dataLabels.ShowBubbleSize = True

            ' Set the minimum and maximum values for the chart value axis.
            Dim axis As Axis = chart.PrimaryAxes(1)
            axis.Scaling.AutoMax = False
            axis.Scaling.Max = 82
            axis.Scaling.AutoMin = False
            axis.Scaling.Min = 64

        End Sub
        Public Shared Function SetActiveWorksheet(ByVal workbook As IWorkbook, ByVal sheetName As String) As Worksheet
            If workbook.Worksheets.ActiveWorksheet IsNot workbook.Worksheets(sheetName) Then
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets(sheetName)
            End If
            Dim worksheet As Worksheet = workbook.Worksheets.ActiveWorksheet
            worksheet.Charts.Clear()
            Return worksheet
        End Function


    End Class
End Namespace
