add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/ClosedXML.dll
add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/DocumentFormat.OpenXml.dll
add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/ExcelNumberFormat.dll
add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/Microsoft.Win32.SystemEvents.dll
add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/System.Drawing.Common.dll
add-type -path /home/kaneko/consoleapp/bin/Debug/net5.0/System.IO.Packaging.dll

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
        $row.Field("SEQ").Value = "OK"
    }else{
        # 偶数の場合エラー
        "error"
        "get screen shot"
        $row.Field("SEQ").Value = "NG"
        "press F2 to go to main menu"
    }
}

$wb = new-object ClosedXML.Excel.XLWorkbook("Showcase.xlsx")
$table = $wb.Table("Table1")
$first = $table.RangeAddress.FirstAddress.RowNumber
$last  = $table.RangeAddress.LastAddress.RowNumber
foreach($row in $table.DataRange.Rows(1,$last - $first ) ){
    regist($row)
}
$wb.Save()

