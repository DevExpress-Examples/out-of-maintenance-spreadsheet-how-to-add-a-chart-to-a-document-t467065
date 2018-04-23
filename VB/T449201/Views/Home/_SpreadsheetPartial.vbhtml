@Imports DevExpress.Web.Office
@Imports T449201

@Html.DevExpress().Spreadsheet(Sub(settings)
    settings.Name = "Spreadsheet"
    settings.CallbackRouteValues = New With {Key .Controller = "Home", Key .Action = "SpreadsheetPartial"}
    settings.CustomActionRouteValues = New With {Key .Controller = "Home", Key .Action = "SpreadsheetCustomPartial"}
    settings.Width = 1100
    settings.Height = 600
    settings.ReadOnly = false
    settings.RibbonMode = SpreadsheetRibbonMode.Ribbon

   settings.PreRender = Function(s, e) 
Dim spreadsheert As MVCxSpreadsheet = DirectCast(s, MVCxSpreadsheet)
spreadsheert.CreateDefaultRibbonTabs(True)
spreadsheert.RibbonTabs(0).Visible = False
'Hide the file tab


DiagramCreationHelper.CreatePieChart(spreadsheert.Document)

End Function
End Sub).Open(Server.MapPath("~/App_Data/WorkDirectory/Document.xlsx")).GetHtml()
