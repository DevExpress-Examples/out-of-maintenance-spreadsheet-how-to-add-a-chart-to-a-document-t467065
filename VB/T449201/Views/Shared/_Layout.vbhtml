<!DOCTYPE html>

<html>
<head>
    <meta charset="UTF-8" />
    <title>@ViewBag.Title</title>

    @Html.DevExpress().GetStyleSheets(
            New StyleSheet With {.ExtensionSuite = ExtensionSuite.NavigationAndLayout },
New StyleSheet With {.ExtensionSuite = ExtensionSuite.Editors },
New StyleSheet With {.ExtensionSuite = ExtensionSuite.Spreadsheet } )
    @Html.DevExpress().GetScripts(
            New Script With {.ExtensionSuite = ExtensionSuite.NavigationAndLayout },
New Script With {.ExtensionSuite = ExtensionSuite.Editors },
New Script With {.ExtensionSuite = ExtensionSuite.Spreadsheet }    )

</head>

<body>
    @RenderBody()
</body>
</html>