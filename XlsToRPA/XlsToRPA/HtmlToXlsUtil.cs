using AngleSharp;
using AngleSharp.Html.Parser;
using ClosedXML.Excel;
using System;
using System.IO;
using System.Threading.Tasks;
using AngleSharp.Dom;

namespace XlsToRPA
{
    class HtmlToXlsUtil
    {
        /*
         * 
         */
        public static void HtmlToWorkbook(string source,string path) {
            
            var parser = new HtmlParser();
           // IXLWorksheet ws;

            using (var wb = new XLWorkbook()) 
            using (var tbl = parser.ParseDocument(source))
            {
                wb.Worksheets.Add("Sheet1");
                var ws = wb.Worksheet("Sheet1");

                var idxx = 1;
                var idxy = 1;

                foreach (var item in tbl.GetElementsByTagName("TR"))
                {
                    idxx = 1;
                    foreach (var item2 in item.ChildNodes)
                    {
                        if (item2 is IElement)
                        {
                            Console.WriteLine(item2.NodeType);
                            ws.Cell(idxy, idxx).Value = item2.TextContent;
                            idxx++;
                        }
                    }
                    idxy++;
                }
                //return ws;

                //                wb.SaveAs(this.variable);
                wb.SaveAs(path);
            }
        }
        /*        public static async Task<int> hoge(string htmlCode,string selector) {
                    //Use the default configuration for AngleSharp
                    var config = Configuration.Default;

                    //Create a new context for evaluating webpages with the given config
                    var context = BrowsingContext.New(config);

                    //Parse the document from the content of a response to a virtual request
                    var document = await context.OpenAsync(req => req.Content(htmlCode));

                    var elem = document.QuerySelector(selector);


                    return 0;
                }*/
    }
}
