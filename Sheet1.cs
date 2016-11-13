using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;

namespace StockAnalyzerWB
{
    public partial class Sheet1
    {
        ILogger logger;
        Workbook workbook;
        Excel.Application app;
        enum OptionType { PUT, CALL }
        enum TemplateSheet { putTemplate, callTemplate };
        enum OptionLetter { P, C }

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            ILoggerFactory loggerFactory = new LoggerFactory()
            .AddDebug();
            logger = loggerFactory.CreateLogger($"XL-Addin-{this.Name}");

            logger.LogInformation("**************** starting *********************");

            workbook = (Workbook)this.Parent;
            app = workbook.Application;

        }


        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
            logger.LogInformation("**************** shutting down *********************");

        }


        private void buttonGetPutOptions_Click(object sender, EventArgs e)
        {
            startAnalysis(OptionType.PUT);
        }

        private void buttonGetCallOptions_Click(object sender, EventArgs e)
        {
            startAnalysis(OptionType.CALL);
        }


     
        void startAnalysis(OptionType type)
        {
            const int symbolColumn = 4;
            const int priceColumn = 5;

            double? lastPrice;
            int row = app.Selection.row;
            string stockSymbol = this.Cells[row, symbolColumn].Value;

            try
            {
                lastPrice = this.Cells[row, priceColumn].Value;
            }
            catch
            {
                lastPrice = null;
            }
            string lastPriceFormula = this.Cells[row, priceColumn].Formula;

            if (stockSymbol == null || lastPrice == null)
            {
                Error("Missing stock symbol or a last price on the current row");
                return;
            }
            logger.LogInformation($"getting option data for {stockSymbol} {lastPrice} {type.ToString()}");
            Range sel = app.Selection;
            getOptionData(stockSymbol, lastPriceFormula, type);
            //workbook.
        }

        void Error(string message)
        {
            logger.LogError(message, "Whoops");
            MessageBox.Show(message);
        }



        void getOptionData(string stockSymbol, string lastPriceFormula, OptionType type)
        {

            string sheetName = ((TemplateSheet)type).ToString();
            string typeID = (type).ToString();

            Workbook outputWorkbook = getWorkbook();
            outputWorkbook.Activate();

            workbook.Sheets[sheetName].Copy(outputWorkbook.Sheets[1]);// .ActiveSheet);
            Worksheet sheet = outputWorkbook.Sheets[sheetName];
            try
            {
                sheet.Name = stockSymbol + $"-{typeID}-" + DateTime.Now.ToString("d.h.m");

            }
            catch
            {
                sheet.Name = stockSymbol + $"-{typeID}-" + DateTime.Now.ToString("d.h.m.s");

            }
            sheet.Cells[2, 2].value = stockSymbol;
            sheet.Cells[2, 3].Formula = lastPriceFormula;
            int iYear = DateTime.Today.Year;
            int iMonth = DateTime.Today.Month;
            Range r = sheet.Cells[3, 1];
            r = GetOptionStrikePrices(sheet, stockSymbol, iYear + 2, 1, r, type);
            r = r.Offset[2, 0];
            r = GetOptionStrikePrices(sheet, stockSymbol, iYear + 1, 1, r, type);
            r = r.Offset[2, 0];
            r = GetOptionStrikePrices(sheet, stockSymbol, iYear, iMonth + 1, r, type);
            r = r.Offset[2, 0];
            r = GetOptionStrikePrices(sheet, stockSymbol, iYear, iMonth, r, type);
            sheet.Range["A3"].Select();
        }
        Range GetOptionStrikePrices(Worksheet sheet, string stockSymbol, int iYear, int iMonth, Range destRange, OptionType type)
        {
            string url = $"http://finance.yahoo.com/q/op?s={stockSymbol}&m={iYear}-{iMonth:00}"; //ex:  "URL;http://finance.yahoo.com/q/op?s=MSFT&m=2018-01"
            QueryTable webQuery = sheet.QueryTables.Add("URL;" + url, destRange);
            webQuery.WebSelectionType = XlWebSelectionType.xlSpecifiedTables;
            if (type == OptionType.PUT)
                webQuery.WebTables = "12,14"; //14 is puts , 12 is header/date
            else
                webQuery.WebTables = "12,13"; //14 is puts , 12 is header/date

            webQuery.WebFormatting = XlWebFormatting.xlWebFormattingRTF;
            webQuery.BackgroundQuery = false;
            webQuery.AdjustColumnWidth = false;
            webQuery.Refresh();

            int firstRow = destRange.Row + 3;
            Range r = destRange.Offset[2, 0]; // app.Selection;
            r = r.End[XlDirection.xlDown];
            int lastRow = r.Row;

            fillInFormulas(sheet, firstRow, lastRow, type);
            return r;
        }

        void fillInFormulas(Worksheet sheet, int firstRow, int lastRow, OptionType type)
        {
            const int symbolColumn = 2;
            const char bidColumn = 'J';
            const char askColumn = 'K';
            //=RTD("tos.rtd", , E$1, ".UPRO180119P"&$A6)
            //=RTD("tos.rtd", , E$1, ".UPRO180119P"&Strike_Price)

            string optionLetter = ((OptionLetter)type).ToString();
            string optionSymbol = sheet.Cells[firstRow, symbolColumn].value;
            string baseSymbol = "." + optionSymbol.Remove(1 + optionSymbol.LastIndexOf(optionLetter));
            string bidCellFormula = $"=RTD(\"tos.rtd\", , {bidColumn}$1, \"{baseSymbol}\"&Strike_Price";
            string askCellFormula = $"=RTD(\"tos.rtd\", , {askColumn}$1, \"{baseSymbol}\"&Strike_Price";
            sheet.Range[$"{bidColumn}{firstRow}:{bidColumn}{lastRow}"].Value = bidCellFormula;
            sheet.Range[$"{askColumn}{firstRow}:{askColumn}{lastRow}"].Value = askCellFormula;

        }

        Workbook getWorkbook()
        {
            Workbook wb;
            string name = "optiondata " + DateTime.Now.ToString("yyyy-MMM") + ".xlsx";
            string filePath = workbook.Path + @"\"; // @"C:\Users\Steven\mystuff\";

            try    //see if already open
            {
                //app.Windows[newWorkbookName].Activate();
                wb = app.Workbooks[name];
            }
            catch
            {
                try  //to open it
                {
                    wb = app.Workbooks.Open(filePath + name);
                }
                catch
                {   //create it
                    wb = app.Workbooks.Add();
                    wb.SaveAs(filePath + name);
                }

            }
            return wb;


        }


        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.buttonGetPutOptions.Click += new System.EventHandler(this.buttonGetPutOptions_Click);
            this.buttonGetCallOptions.Click += new System.EventHandler(this.buttonGetCallOptions_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

    
    }
}
