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
        using(var wb = new XLWorkbook())
        {
            var data = Prepare.MakeTable(wb);
            View View = new View();

            foreach (var row in data.DataRange.Rows())
            {
                View.row = row;
                View.regist();
            }            
            wb.SaveAs( DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx" );
        }
        }
    }

    class View {
        private Dictionary<string,(int row, int column)> colinfo 
            = new Dictionary<string,(int row, int column)>
        {
            {"columnC",(1,3)},
            {"columnD",(2,4)},
            {"columnE",(5,7)}
        };

        public IXLTableRow row;

        public void regist() {

            foreach (string colname in colinfo.Keys)
            {
                string[] values = {
                    row.Field(colname).GetString(),
                    colinfo[colname].row.ToString(),
                    colinfo[colname].column.ToString()
                };
                Console.WriteLine( String.Format("set {0} at ({1},{2})",values) );
            }
            try {
                // 
                Console.WriteLine("press ENTER");
                Console.WriteLine("get message below");
                Random r1 = new System.Random();
                switch ( r1.Next(0,4) )
                {
                    case 0:
                        break;
                    case 1:
                        throw new registException("error1");
                    case 2:
                        throw new registException("error2");
                    default:
                        throw new registException("error3");
                }
                Console.WriteLine("no error message");
                Console.WriteLine("press Enter again");
                row.Field("status").Value = "OK";

            }catch(registException){
                // 
                Console.WriteLine("error");
                Console.WriteLine("get screen shot");
                row.Field("status").Value = "NG";
            }finally{
                Console.WriteLine("press F2 to go to main menu");
            }
        }
    }
    class registException : Exception {
        // 
        public registException() :base() {}
        public registException(string str) :base(str) {}

        public override string ToString()
        {
            return Message;
        }


    }

    static class Prepare {
        // 
        public static IXLTable MakeTable(XLWorkbook wb)
        {
            // 
            var ws = wb.Worksheets.Add("data");
            // set data
            ws.Cell("A1").Value = "SEQ";
            ws.Cell("A2").Value = "1";
            ws.Cell("A3").Value = "2";
            ws.Cell("A4").Value = "3";
            ws.Cell("A5").Value = "4";
            ws.Cell("A6").Value = "5";

            ws.Cell("B1").Value = "status";

            ws.Cell("C1").Value = "columnC";
            ws.Cell("C2").Value = "C1";
            ws.Cell("C3").Value = "C2";
            ws.Cell("C4").Value = "C3";
            ws.Cell("C5").Value = "C4";
            ws.Cell("C6").Value = "C5";

            ws.Cell("D1").Value = "columnD";
            ws.Cell("D2").Value = "D1";
            ws.Cell("D3").Value = "D2";
            ws.Cell("D4").Value = "D3";
            ws.Cell("D5").Value = "D4";
            ws.Cell("D6").Value = "D5";

            ws.Cell("E1").Value = "columnE";
            ws.Cell("E2").Value = "E1";
            ws.Cell("E3").Value = "E2";
            ws.Cell("E4").Value = "E3";
            ws.Cell("E5").Value = "E4";
            ws.Cell("E6").Value = "E5";

            return ws.RangeUsed().CreateTable();
        }
        public static void testLinq(IXLTable table){
            Int32 currentRow = table.RangeAddress.FirstAddress.RowNumber; // reset the currentRow
            var Linq1 = table.DataRange.Rows().Where(row => row.Field("columnA").GetString() == "A1").Select(r => r);
            var Linq2 = from row in table.DataRange.Rows()
                            where row.Field("columnC").GetString() == "C1"
                            select row;
            foreach(var row2 in Linq1 ) Console.WriteLine( row2.Field("columnC").GetString() );
        }
    }
}
