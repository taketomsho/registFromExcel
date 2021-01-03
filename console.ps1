add-type -path ./bin/Debug/net5.0/ClosedXML.dll
add-type -path ./bin/Debug/net5.0/DocumentFormat.OpenXml.dll
add-type -path ./bin/Debug/net5.0/ExcelNumberFormat.dll
add-type -path ./bin/Debug/net5.0/Microsoft.Win32.SystemEvents.dll
add-type -path ./bin/Debug/net5.0/System.Drawing.Common.dll
add-type -path ./bin/Debug/net5.0/System.IO.Packaging.dll


function regist($row) {
    $colinfo = @{
        columnC = @{row = 1 ; column =3 }
        columnD = @{row = 2 ; column =4 }
        columnE = @{row = 5 ; column =7 }
    }
    foreach($colname in $colinfo.keys){
        $values = @(
            $row.Field($colname).GetString(),
            $colinfo.$colname.row,
            $colinfo.$colname.column
        )
        "set `"{0}`" at ( {1} , {2} )" -f $values
    }
    try{
        "press ENTER"
        "get message below"
        # 奇数の場合エラー
        $errornum = (Get-Random) % 4
        if( $errornum -ne 1 ) { throw "error code is $errornum" }
        "no error message"
        "press Enter again"
        $row.Field("status").Value = "OK"
    }catch{
        $PSItem.ToString()
        "get screen shot"
        $row.Field("status").Value = "NG"
    }finally{
        "press F2 to go to main menu"
    }
}

$wb = new-object ClosedXML.Excel.XLWorkbook
$ws = $wb.Worksheets.Add("data");
$ws.Cell("A1").Value = "SEQ";
$ws.Cell("A2").Value = "1";
$ws.Cell("A3").Value = "2";
$ws.Cell("A4").Value = "3";
$ws.Cell("A5").Value = "4";
$ws.Cell("A6").Value = "5";
$ws.Cell("B1").Value = "status";    
$ws.Cell("C1").Value = "columnC";
$ws.Cell("C2").Value = "C1";
$ws.Cell("C3").Value = "C2";
$ws.Cell("C4").Value = "C3";
$ws.Cell("C5").Value = "C4";
$ws.Cell("C6").Value = "C5";
$ws.Cell("D1").Value = "columnD";
$ws.Cell("D2").Value = "D1";
$ws.Cell("D3").Value = "D2";
$ws.Cell("D4").Value = "D3";
$ws.Cell("D5").Value = "D4";
$ws.Cell("D6").Value = "D5";
$ws.Cell("E1").Value = "columnE";
$ws.Cell("E2").Value = "E1";
$ws.Cell("E3").Value = "E2";
$ws.Cell("E4").Value = "E3";
$ws.Cell("E5").Value = "E4";
$ws.Cell("E6").Value = "E5";
$table = $ws.RangeUsed().CreateTable()

$first = $table.RangeAddress.FirstAddress.RowNumber
$last  = $table.RangeAddress.LastAddress.RowNumber
foreach($row in $table.DataRange.Rows(1,$last - $first ) ){
    regist($row)
}
$wb.SaveAs( $(Get-Date -Format "yyyyMMddHHmmss") + ".xlsx")
