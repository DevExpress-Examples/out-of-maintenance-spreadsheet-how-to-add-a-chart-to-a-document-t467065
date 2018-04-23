@Code
    ViewBag.Title = "Spreadsheet - How to add a chart to a document"
End Code

<script type="text/javascript">
    function OnSelectedIndexChanged(s, e) {
        Spreadsheet.PerformCallback({ diagramName: s.GetValue()});
    }
</script>

@Html.DevExpress().RadioButtonList(Function(settings)
settings.Name = "RadioButtonList"
settings.Properties.RepeatColumns = 3
settings.Properties.Items.Add("PieChart")
settings.Properties.Items.Add("BarChart")
settings.Properties.Items.Add("ColumnChart")
settings.Properties.Items.Add("ComplexChart")
settings.Properties.Items.Add("DoughnutChart")
settings.Properties.Items.Add("Pie3dChart")
settings.Properties.Items.Add("ScatterChart")
settings.Properties.Items.Add("StockChart")
settings.Properties.Items.Add("BubbleChart")

settings.PreRender = Function(s, e)
Dim list As MVCxRadioButtonList = DirectCast(s, MVCxRadioButtonList)
list.Value = "PieChart"

End Function
settings.Properties.ClientSideEvents.SelectedIndexChanged = "OnSelectedIndexChanged"


End Function).GetHtml()

<br/>
<br/>
@Code
Using Html.BeginForm()
	Html.RenderAction("SpreadsheetPartial")
End Using
End Code

