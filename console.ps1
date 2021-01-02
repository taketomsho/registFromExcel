add-type -path ./bin/Debug/net5.0/ClosedXML.dll
add-type -path ./bin/Debug/net5.0/DocumentFormat.OpenXml.dll
add-type -path ./bin/Debug/net5.0/ExcelNumberFormat.dll
add-type -path ./bin/Debug/net5.0/Microsoft.Win32.SystemEvents.dll
add-type -path ./bin/Debug/net5.0/System.Drawing.Common.dll
add-type -path ./bin/Debug/net5.0/System.IO.Packaging.dll


function regist($row) {
    $colinfo = @{
        column1 = @{row = 1 ; column =3 }
        column2 = @{row = 2 ; column =4 }
    }
    foreach($colname in $colinfo.keys){
        "set `"$($row.Field($colname).GetString())`" at ( $($colinfo.$colname.row) , $($colinfo.$colname.column) )"
    }
    "press ENTER"
    "get message below"

    if( (Get-Random) % 2 -eq 1 )
    {
        # 奇数の場合エラーなし
        "no error message"
        "press Enter again"
        "press F2 to go to main menu"
        $row.Field("status").Value = "OK"
    }else{
        # 偶数の場合エラー
        "error"
        "get screen shot"
        $row.Field("status").Value = "NG"
        "press F2 to go to main menu"
    }
}

$wb = new-object ClosedXML.Excel.XLWorkbook
$ws = $wb.Worksheets.Add("data");
$ws.Cell("A1").Value = "status";
$ws.Cell("B1").Value = "column1";
$ws.Cell("B2").Value = "1";
$ws.Cell("B3").Value = "2";
$ws.Cell("B4").Value = "3";
$ws.Cell("C1").Value = "column2";
$ws.Cell("C2").Value = "1";
$ws.Cell("C3").Value = "2";
$ws.Cell("C4").Value = "3";

$table = $ws.RangeUsed().CreateTable()
$first = $table.RangeAddress.FirstAddress.RowNumber
$last  = $table.RangeAddress.LastAddress.RowNumber
foreach($row in $table.DataRange.Rows(1,$last - $first ) ){
    regist($row)
}
$wb.SaveAs( $(Get-Date -Format "yyyyMMddHHmmss") + ".xlsx")

<#
$wb = new-object ClosedXML.Excel.XLWorkbook("Showcase.xlsx")
$r = $wb.Table("Table1").DataRange.Row(1)
c:\Users\kanek\Desktop\console\Showcase.xlsx
#>