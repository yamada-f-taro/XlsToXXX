using System;
using System.Xml;
using ClosedXML.Excel;
//using System.Xml.Linq;
//using System.Collections.Generic;
//using System.Linq;
using System.IO;

namespace XmlToXls
{
    class Program
    {
        static void Main(string[] args)
        {
            var fn = "";
            if (args.Length == 1) fn = args[0];
            else fn = @".\XMLFile1.xml";

            if (!File.Exists(@".\result.xlsx")) File.Delete(@".\result.xlsx");


            using (var book = new XLWorkbook())
            {
                // ワークシートを作成し、シートを取得
                var sheet = book.Worksheets.Add("Sheet1");

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(fn);

                XmlNodeList nodeList = xmlDocument.SelectNodes("//ScriptCatalog/Script");

                var rownum = 1;

                foreach (var item in nodeList)
                {

                    XmlElement elem = (XmlElement)item;

                    sheet.Cell(rownum, 1).Value = elem.GetAttribute("name");

                    XmlNodeList nodeList2 = elem.SelectNodes("//StepList/Step");
                    //Script一つについての処理

                    foreach (var item2 in nodeList2)
                    {
                        XmlElement elem2 = (XmlElement)item2;
                        sheet.Cell(rownum, 2).Value = elem2.GetAttribute("name");

                        XmlElement elem3 = (XmlElement)elem2.ChildNodes[0];

                        sheet.Cell(rownum, 3).Value = elem3.InnerText;

                        rownum++;
                    }
                }
                // ファイルに保存
                book.SaveAs(@".\result.xlsx");
            }
            Console.WriteLine("end");
            Console.ReadKey();
        }
    }
}
