using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OpenQA.Selenium;
using ClosedXML;
using ClosedXML.Excel;

namespace XlsToRPA
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 1) return;

            var sttms = new List<RPAStatement>();
            sttms = readXls(args[0]);

            foreach (var itm in sttms)
            {
                itm.exec();
            }
            Console.WriteLine("処理が完了しました。");
            Console.ReadLine();
            RPAStatement.MyDispose();
        }
        static List<RPAStatement> readXls(string pasu)
        {
            var ret = new List<RPAStatement>();
            //wbよみこみ
            using (var wb = new XLWorkbook(pasu))
            {
                var ws = wb.Worksheets.Where(s => s.Name == "Sheet1").FirstOrDefault();
                //一行目はヘッダ
                for (int idx = 2; idx < ws.RowCount(); idx++)
                {
                    var row = ws.Row(idx);
                    var itm = new RPAStatement();
                    itm.id = row.Cell(1).Value.ToString();//id

                    if (itm.id.Equals("")) break;

                    itm.line = row.RowNumber();//

                    itm.operation1 = row.Cell(2).Value.ToString();
                    itm.operation2 = row.Cell(3).Value.ToString();
                    itm.operation3 = row.Cell(4).Value.ToString();

                    itm.argument1 = row.Cell(5).Value.ToString();
                    itm.argument2 = row.Cell(6).Value.ToString();
                    itm.argument3 = row.Cell(7).Value.ToString();

                    itm.variable = row.Cell(8).Value.ToString();
                    itm.memo = row.Cell(9).Value.ToString();
                    ret.Add(itm);
                }
                return ret;
            }
        }
    }
}
