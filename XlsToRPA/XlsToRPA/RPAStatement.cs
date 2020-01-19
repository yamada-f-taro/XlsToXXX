using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using OpenQA.Selenium.Support.UI;

using System.Collections.ObjectModel;

//使ってないけど参考：
//https://tech.sanwasystem.com/entry/2016/06/27/180903

namespace XlsToRPA
{
    public class RPAStatement
    {
        public string id { get; set; }
        public int line { get; set; }
        public string operation1 { get; set; }
        public string operation2 { get; set; }
        public string operation3 { get; set; }
        public string argument1 { get; set; }
        public string argument2 { get; set; }
        public string argument3 { get; set; }
        public string variable { get; set; }
        public string memo { get; set; }

        private static ChromeDriver driver { get; set; }

        private static Dictionary<string,string> vari { get; set; }

        public RPAStatement() {
            if (driver == null)
            {
                var options = new ChromeOptions();
                // --headlessを追加します。
                options.AddArgument("--headless");
                driver = new ChromeDriver(options);
                vari = new Dictionary<string,string>();
            }
        }

        public static void MyDispose()
        {
            if (driver != null)
            {
                driver.Dispose();
            }
        }

        public Boolean exec() {
            var ret = false;
            if (operation1.Equals("open") && operation2.Equals("")) this.openUrl();
            else if (operation1.Equals("close") && operation2.Equals("")) this.close();
            else if (operation1.Equals("click") && operation2.Equals("link")) this.click();
            else if (operation1.Equals("select") && operation2.Equals("option")) this.select();

            else if (operation1.Equals("get") && operation2.Equals("table")) this.getTable();
            else if (operation1.Equals("get") && operation2.Equals("text")) this.getText();
            else if (operation1.Equals("get") && operation2.Equals("screenShot")) this.getss();

            else if (operation1.Equals("set") && operation2.Equals("text")) this.setText();
            return ret;
        }

        private void openUrl() {
            Console.WriteLine(this.argument1);
            driver.Navigate().GoToUrl(this.argument1);
        }

        private void click()
        {
            ReadOnlyCollection<IWebElement> elems = null;// driver.FindElementsByXPath(this.argument1);
            IWebElement elem = null;// driver.FindElementsByXPath(this.argument1);

            if (this.operation3.Equals("id"))
            {
                elem = driver.FindElementById(this.argument1);
            }
            else if(this.operation3.Equals("xpath"))
            {
                elems = driver.FindElementsByXPath(this.argument1);
                elem = elems.First();
            }

            if (elem == null) throw new Exception("elementの検索でなにも見つからなかった(click)");

            elem.Click();

        }

        private void close() {
            driver.Close();
        }

        private void getTable() {

            if (this.operation3.Equals("excel")) {
                //excelにする
                var tmp = driver.FindElementsByXPath(this.argument1);

                if (tmp.Count() == 0) throw new Exception("xpathでなにも見つからなかった(getTable)");

                //一つ上のtableタグを指定する
                var elem = tmp[0].FindElements(By.XPath(".."));

                HtmlToXlsUtil.HtmlToWorkbook(elem[0].GetAttribute("innerHTML"), this.variable);
            }
        }
        private static IWebElement myGetElement(string ope, string arg) {
            IWebElement ret = null;
            if (ope.Equals("xpath")) {
                var elems = driver.FindElementsByXPath(arg);
                if (elems.Count() == 0)
                {
                    throw new Exception("なにも見つからなかった・・・");
                }
                ret = elems[0];
            }
            else if (ope.Equals("id")){
                var elems = driver.FindElementsById(arg);
                if (elems.Count() == 0)
                {
                    throw new Exception("なにも見つからなかった・・・");
                }
                ret = elems[0];
            }
            else //なかったらnameとして判断
            {
                var elems = driver.FindElementsByName(arg);
                if (elems.Count() == 0)
                {
                    throw new Exception("なにも見つからなかった・・・");
                }
                ret = elems[0];
            }

            return ret;
        }

        private void select() {
            var elems = driver.FindElementsByXPath(this.argument1);

            if (elems.Count() == 0)
            {
                throw new Exception("xpathでなにも見つからなかった(select)");
            }
            IWebElement element = elems[0];
            SelectElement selectElement = new SelectElement(element);
            selectElement.SelectByText(this.argument2);
        }

        private void setText() {
            var elems = driver.FindElementsByXPath(this.argument1);

            if (elems.Count() == 0)
            {
                throw new Exception("xpathでなにも見つからなかった(setText)");
            }
            IWebElement element = elems[0];
            element.SendKeys(this.variable);

        }

        private void getText() {
            var fn = this.argument1.ToString();//ファイル名
            var sb = new StringBuilder();

            //Excel ファイルを開く
            using (var wb = new XLWorkbook(this.argument1)){
                //シート名を指定してシートを取得
                var ws = wb.Worksheets.Where(s => s.Name == this.argument2).FirstOrDefault();

                for (var idx = 1; idx < ws.RowCount(); idx++) {
                    var row = ws.Row(idx);
                    sb.Append(row.Cell(1).Value.ToString());
                    if (row.Cell(1).Value.Equals("")) {
                        break;
                    }
                }
            }
            vari.Add(this.variable, sb.ToString());//変数名：値のDic
        }

        /*
        private void getTableOld()
        {
            var elems = driver.FindElementsByXPath(this.argument1);
            var elems2 = driver.FindElementsByXPath("//meta[@charset]").FirstOrDefault();
            var enc = elems2.GetAttribute("charset");

            if (elems.Count() == 0 ) {
                throw new Exception("xpathでなにも見つからなかった(get)");
            }
            using (var sr = new StreamWriter(this.variable))
            {
                sr.Write("<html>");
                sr.Write("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=" + enc + "\">");
                sr.Write(elems[0].GetAttribute("innerHTML"));
                sr.Write("</html>");
            }
        }
        */
        private void getss()
        {
            //var hogee =  HtmlToXlsUtil.hoge("", this.argument1);
            //hogee.Wait();
            var ss = driver.GetScreenshot();
            Console.WriteLine("SaveAsFile : " + this.variable);
            ss.SaveAsFile(this.variable);
        }
    }
}
