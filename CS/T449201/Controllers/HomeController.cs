using DevExpress.Web.Mvc;
using DevExpress.Web.Office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DevExpress.Spreadsheet;


namespace T449201.Controllers {
    public class HomeController : Controller {
        public ActionResult Index() {
            return View();
        }

        public ActionResult SpreadsheetPartial() {
            return PartialView("_SpreadsheetPartial");
        }
        public ActionResult SpreadsheetCustomPartial(string diagramName) {
            IWorkbook wbook =  SpreadsheetExtension.GetCurrentDocument("Spreadsheet");
            
            switch (diagramName) {
                case "PieChart":
                    DiagramCreationHelper.CreatePieChart(wbook);
                    break;
                case "BarChart":
                    DiagramCreationHelper.CreateBarChart(wbook);
                    break;
                case "ColumnChart":
                    DiagramCreationHelper.CreateColumnChart(wbook);
                    break;
                case "ComplexChart":
                    DiagramCreationHelper.CreateComplexChart(wbook);
                    break;
                case "DoughnutChart":
                    DiagramCreationHelper.CreateDoughnutChart(wbook);
                    break;
                case "Pie3dChart":
                    DiagramCreationHelper.CreatePie3dChart(wbook);
                    break;
                case "ScatterChart":
                    DiagramCreationHelper.CreateScatterChart(wbook);
                    break;
                case "StockChart":
                    DiagramCreationHelper.CreateStockChart(wbook);
                    break;
                case "BubbleChart":
                    DiagramCreationHelper.CreateBubbleChart(wbook);
                    break;
            }

                    return PartialView("_SpreadsheetPartial");
        }
    }
}