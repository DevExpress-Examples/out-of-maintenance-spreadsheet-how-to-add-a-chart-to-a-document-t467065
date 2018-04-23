﻿using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace T449201 {
    public class DiagramCreationHelper {
        public static void CreatePieChart(IWorkbook wbook) {

            DevExpress.Spreadsheet.Worksheet worksheet = SetActiveWorksheet(wbook, "Range1");

            // Create a chart and specify its location
            Chart chart = worksheet.Charts.Add(ChartType.PieExploded, worksheet["B2:C7"]);

            // Display the chart title
            chart.Title.Visible = true;
            chart.Title.SetReference(worksheet["B1"]);

            chart.TopLeftCell = worksheet.Cells["E2"];
            chart.BottomRightCell = worksheet.Cells["K15"];

            // Set the chart style
            chart.Style = ChartStyle.ColorGradient;

            // Hide the legend
            chart.Legend.Visible = false;

            // Rotate the pie chart view
            chart.Views[0].FirstSliceAngle = 100;

            // Display data labels
            DataLabelOptions dataLabels = chart.Views[0].DataLabels;
            dataLabels.ShowCategoryName = true;
            dataLabels.ShowPercent = true;
            dataLabels.Separator = "\n";
        }
        public static void CreateBarChart(IWorkbook wbook) {

            Worksheet worksheet = SetActiveWorksheet(wbook, "Range1");

            Chart chart = worksheet.Charts.Add(ChartType.BarFullStacked);
            chart.TopLeftCell = worksheet.Cells["E2"];
            chart.BottomRightCell = worksheet.Cells["K15"];

            // Select chart data
            chart.SelectData(worksheet["B2:C7"], ChartDataDirection.Row);

            // Display the chart title
            chart.Title.Visible = true;
            chart.Title.SetReference(worksheet["B1"]);

            // Change legend position
            chart.Legend.Position = LegendPosition.Bottom;

            // Hide the category axis
            chart.PrimaryAxes[0].Visible = false;

            // Set major unit of the value axis
            chart.PrimaryAxes[1].MajorUnit = 0.2;

        }
        public static void CreateColumnChart(IWorkbook wbook) {
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range2");
            //Create data range for a chart

            Chart chart = worksheet.Charts.Add(ChartType.ColumnClustered);
            chart.TopLeftCell = worksheet.Cells["H2"];
            chart.BottomRightCell = worksheet.Cells["N14"];

            // Add series
            chart.Series.Add(worksheet["D2"], worksheet["B3:B6"], worksheet["D3:D6"]);
            chart.Series.Add(worksheet["F2"], worksheet["B3:B6"], worksheet["F3:F6"]);

            // Display the chart title
            chart.Title.Visible = true;
            chart.Title.SetValue("Mobile OS market share");

            // Customize the appearance and scale of the axes
            Axis axis = chart.PrimaryAxes[0];
            axis.MajorTickMarks = AxisTickMarks.None;
            axis = chart.PrimaryAxes[1];
            axis.Outline.SetNoFill();
            axis.MajorTickMarks = AxisTickMarks.None;
            axis.NumberFormat.FormatCode = "0%";
            axis.NumberFormat.IsSourceLinked = false;
            axis.Scaling.AutoMax = false;
            axis.Scaling.Max = 1;
            axis.Scaling.AutoMin = false;
            axis.Scaling.Min = 0;

            // Set the gap width between data series
            ChartView view = chart.Views[0];
            view.GapWidth = 75;

            // Display data labels
            view.DataLabels.ShowValue = true;
            view.DataLabels.NumberFormat.FormatCode = "0%";
            view.DataLabels.NumberFormat.IsSourceLinked = false;

            // Set the chart style
            chart.Style = ChartStyle.ColorGradient;
        }
        public static void CreateComplexChart(IWorkbook wbook) {

            //Create data range for a chart
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range3");

            Chart chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet["B2:D8"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["L15"];

            // Change the chart type of the second series
            chart.Series[1].ChangeType(ChartType.Line);
            chart.Series[1].Smooth = true;

            // Use secondary axes
            chart.Series[1].AxisGroup = AxisGroup.Secondary;

            // Specify the chart style
            chart.Style = ChartStyle.ColorGradient;

            // Set the position of the legend
            chart.Legend.Position = LegendPosition.Bottom;
        }
        public static void CreateDoughnutChart(IWorkbook wbook) {

            //Create data range for a chart
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range2");

            // Create a chart and specify its location
            Chart chart = worksheet.Charts.Add(ChartType.Doughnut);
            chart.TopLeftCell = worksheet.Cells["H2"];
            chart.BottomRightCell = worksheet.Cells["N14"];

            // Add the data series
            chart.Series.Add(worksheet["E2"], worksheet["B3:B6"], worksheet["E3:E6"]);

            // Display the chart title
            chart.Title.Visible = true;
            chart.Title.SetValue("Mobile OS market share Q4'13");

            // Change the hole size
            chart.Views[0].HoleSize = 60;

            // Display the data labels
            chart.Views[0].DataLabels.ShowPercent = true;
        }
        public static void CreatePie3dChart(IWorkbook wbook) {

            Worksheet worksheet = SetActiveWorksheet(wbook, "Range2");

            Chart chart = worksheet.Charts.Add(ChartType.Pie3D);
            chart.TopLeftCell = worksheet.Cells["H2"];
            chart.BottomRightCell = worksheet.Cells["N14"];

            // Add the data series
            chart.Series.Add(worksheet["E2"], worksheet["B3:B6"], worksheet["E3:E6"]);

            // Set the explosion value for the slice
            chart.Series[0].CustomDataPoints.Add(2).Explosion = 25;

            // Set the rotation of the  3-D chart view
            chart.View3D.YRotation = 255;

            // Set the chart style
            chart.Style = ChartStyle.ColorGradient;

        }
        public static void CreateScatterChart(IWorkbook wbook) {
            //Create data range for a chart
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range4");

            Chart chart = worksheet.Charts.Add(ChartType.ScatterLineMarkers, worksheet["C2:D52"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["L15"];

            // Set the marker symbol
            chart.Series[0].Marker.Symbol = MarkerStyle.Circle;

            // Set appearance and scale of the X axis
            Axis axis = chart.PrimaryAxes[0];
            axis.Scaling.AutoMax = false;
            axis.Scaling.AutoMin = false;
            axis.Scaling.Max = 60.0;
            axis.Scaling.Min = -60.0;
            axis.MajorGridlines.Visible = true;

            // Set appearance and scale of the Y axis
            axis = chart.PrimaryAxes[1];
            axis.Scaling.AutoMax = false;
            axis.Scaling.AutoMin = false;
            axis.Scaling.Max = 50.0;
            axis.Scaling.Min = -50.0;
            axis.MajorUnit = 10.0;



        }
        public static void CreateStockChart(IWorkbook wbook) {
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range5");
            Chart chart = worksheet.Charts.Add(ChartType.StockOpenHighLowClose, worksheet["B2:F7"]);
            chart.TopLeftCell = worksheet.Cells["H2"];
            chart.BottomRightCell = worksheet.Cells["N15"];

            // Display the chart title
            chart.Title.Visible = true;
            chart.Title.SetValue("NASDAQ:MSFT");

            // Hide the legend
            chart.Legend.Visible = false;

            // Set appearance and scale of the value axis
            Axis axis = chart.PrimaryAxes[1];
            axis.Scaling.AutoMax = false;
            axis.Scaling.Max = 40.5;
            axis.Scaling.AutoMin = false;
            axis.Scaling.Min = 38.5;
            axis.MajorUnit = 0.25;

            // Format the axis labels
            axis.NumberFormat.FormatCode = "#0.00";
            axis.NumberFormat.IsSourceLinked = false;

            // Display the axis title
            axis.Title.Visible = true;
            axis.Title.SetValue("Price in USD");


        }
        public static void CreateBubbleChart(IWorkbook wbook) {
            Worksheet worksheet = SetActiveWorksheet(wbook, "Range6");

            // Create a chart and specify its location
            Chart chart = worksheet.Charts.Add(ChartType.Bubble3D);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["L15"];

            Series s1 = chart.Series.Add(worksheet["A3"], worksheet["C3:C7"], worksheet["D3:D7"]);
            s1.BubbleSize = ChartData.FromRange(worksheet["E3:E7"]);
            Series s2 = chart.Series.Add(worksheet["A9"], worksheet["C9:C13"], worksheet["D9:D13"]);
            s2.BubbleSize = ChartData.FromRange(worksheet["E9:E13"]);

            // Set the chart style
            chart.Style = ChartStyle.ColorGradient;
            // Set the bubble size 1.5x relative to the default setting.
            chart.Views[0].BubbleScale = 150;

            // Hide the legend
            chart.Legend.Visible = false;

            // Display data labels
            DataLabelOptions dataLabels = chart.Views[0].DataLabels;
            dataLabels.ShowBubbleSize = true;

            // Set the minimum and maximum values for the chart value axis.
            Axis axis = chart.PrimaryAxes[1];
            axis.Scaling.AutoMax = false;
            axis.Scaling.Max = 82;
            axis.Scaling.AutoMin = false;
            axis.Scaling.Min = 64;

        }
        public static Worksheet SetActiveWorksheet(IWorkbook workbook, string sheetName) {
            if (workbook.Worksheets.ActiveWorksheet != workbook.Worksheets[sheetName])
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            Worksheet worksheet = workbook.Worksheets.ActiveWorksheet;
            worksheet.Charts.Clear();
            return worksheet;
        }


    }
}
