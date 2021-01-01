using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Console;

namespace consoleapp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open Workbook and make a sheet (sheet name is timestamp)
            var wb = new XLWorkbook("Showcase.xlsx");
            var ws = wb.Worksheets.Add( DateTime.Now.ToString("yyyyMMddHHmmss") );

            // set data
            ws.Cell(1, 1).Value = "columnA";
            ws.Cell(2, 1).Value = "A1";
            ws.Cell(3, 1).Value = "A2";
            ws.Cell(4, 1).Value = "A3";

            ws.Cell(1, 2).Value = "columnB";
            ws.Cell(2, 2).Value = "B1";
            ws.Cell(3, 2).Value = "B2";
            ws.Cell(4, 2).Value = "B3";

            // make table
            var table = ws.RangeUsed().CreateTable();

            wb.Save();

            Int32 currentRow = table.RangeAddress.FirstAddress.RowNumber; // reset the currentRow
            foreach (var row in table.DataRange.Rows())
            {
            // currentRow++;
            var A = row.Field("columnA").GetString();
            var B = row.Field("columnB").GetString();
            var AandB = String.Format("{0} {1}", A, B);
            Console.WriteLine(AandB);
            }


            // var posNums = table.DataRange.Rows().Where(row => row.Field("columnA").GetString() == "A1").Select(r => r);
            var posNums = from row in table.DataRange.Rows()
                            where row.Field("columnA").GetString() == "A1"
                            select row;

            foreach(var row2 in posNums ) Console.WriteLine( row2.Field("columnB").GetString() );
        }
    }

    // class View {
    //     // regist class
    //     private colinfo

    //     public void regist(row) {
    //         foreach (var colname in colinfo.keys)
    //         {
    //             Console.WriteLine("set (row.Field(colname).GetString()) at ( (colinfo.colname.row) , (colinfo.colname.column) )");
                
    //         }
    //     } 
    // }
}
