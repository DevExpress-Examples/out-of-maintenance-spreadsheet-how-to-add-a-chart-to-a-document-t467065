Imports DevExpress.Web.Mvc
Imports DevExpress.Web.Office
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.Mvc
Imports DevExpress.Spreadsheet


Namespace T449201.Controllers
	Public Class HomeController
		Inherits Controller

		Public Function Index() As ActionResult
			Return View()
		End Function

		Public Function SpreadsheetPartial() As ActionResult
			Return PartialView("_SpreadsheetPartial")
		End Function
		Public Function SpreadsheetCustomPartial(ByVal diagramName As String) As ActionResult
			Dim wbook As IWorkbook = SpreadsheetExtension.GetCurrentDocument("Spreadsheet")

			Select Case diagramName
				Case "PieChart"
					DiagramCreationHelper.CreatePieChart(wbook)
				Case "BarChart"
					DiagramCreationHelper.CreateBarChart(wbook)
				Case "ColumnChart"
					DiagramCreationHelper.CreateColumnChart(wbook)
				Case "ComplexChart"
					DiagramCreationHelper.CreateComplexChart(wbook)
				Case "DoughnutChart"
					DiagramCreationHelper.CreateDoughnutChart(wbook)
				Case "Pie3dChart"
					DiagramCreationHelper.CreatePie3dChart(wbook)
				Case "ScatterChart"
					DiagramCreationHelper.CreateScatterChart(wbook)
				Case "StockChart"
					DiagramCreationHelper.CreateStockChart(wbook)
				Case "BubbleChart"
					DiagramCreationHelper.CreateBubbleChart(wbook)
			End Select

					Return PartialView("_SpreadsheetPartial")
		End Function
	End Class
End Namespace