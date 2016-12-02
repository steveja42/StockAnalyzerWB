using System;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;

namespace StockAnalyzerWB
{
    class CallOptionSheet : OptionSheet
    {
        public CallOptionSheet(Workbook sourceWorkbook, Stock? stockIn) : base(sourceWorkbook, stockIn)
        {
            optionType = "CALL";
            templateSheetName = "callTemplate";
            optionLetter = "C";
            webQueryTables = "8,10"; // 8 is header/date, 10 is calls

        }
    }

    class PutOptionSheet:OptionSheet
    {
        public PutOptionSheet(Workbook sourceWorkbook, Stock? stockIn) : base(sourceWorkbook, stockIn)
        {
            optionType = "PUT";
            templateSheetName = "putTemplate";
            optionLetter = "P";
            webQueryTables = "12,14"; // 12 is header/date, 14 is puts 

        }
       
    }

    class OptionSheet
    {
        enum OptionType { PUT, CALL }
        enum TemplateSheet { putTemplate, callTemplate };
        enum OptionLetter { P, C }

        protected string optionType;
        protected string templateSheetName;
        protected string optionLetter;
        protected string webQueryTables; // 12 is header/date, 14 is puts 

        const int symbolColumn = 2;

        Workbook outputWorkbook;
         Microsoft.Office.Interop.Excel.Application app;
         Worksheet sheet;
         Workbook sourceWorkbook;
        Stock stock;
        static ILogger logger;

        protected OptionSheet(Workbook sourceWorkbook, Stock? stockIn)
        {
            if (logger == null)
            {
                ILoggerFactory loggerFactory = new LoggerFactory()
                .AddDebug();
                logger = loggerFactory.CreateLogger($"XL-Addin-OptionSheet");

            }


            stock = (Stock)stockIn;
            this.sourceWorkbook = sourceWorkbook;
        }
        //makeSheet - fills in a new sheet with option prices for the coming 2 Januarys and the upcoming 2 months
        public void makeSheet()
        { 
            app = sourceWorkbook.Application;


            outputWorkbook = getOpenOrCreateWorkbook(sourceWorkbook);
            outputWorkbook.Activate();

            sourceWorkbook.Sheets[templateSheetName].Copy(outputWorkbook.Sheets[1]);// .ActiveSheet);
            sheet = outputWorkbook.Sheets[templateSheetName];
            try
            {
                sheet.Name = stock.symbol + $"-{optionType}-" + DateTime.Now.ToString("d.h.m");

            }
            catch
            {
                sheet.Name = stock.symbol + $"-{optionType}-" + DateTime.Now.ToString("d.h.m.s");

            }
            sheet.Cells[2, 2].value = stock.symbol;
            sheet.Cells[2, 3].Formula = stock.lastPriceFormula;
            

            int iYear = DateTime.Today.Year;
            int iMonth = DateTime.Today.Month;
            Range r = sheet.Cells[3, 1];
            r = GetOptionStrikePrices(stock.symbol, iYear + 2, 1, r);
            r = r.Offset[2, 0];
            r = GetOptionStrikePrices(stock.symbol, iYear + 1, 1, r);
            r = r.Offset[2, 0];

            if (iMonth != 12)  //january is already done above
            {
                r = GetOptionStrikePrices(stock.symbol, (iMonth == 12 ? iYear + 1 : iYear), (iMonth + 1) % 12, r);
                r = r.Offset[2, 0];
            }
            r = GetOptionStrikePrices(stock.symbol, iYear, iMonth, r);


            foreach (Range r2 in sheet.Range["N5:Y300"].Cells)
            {
                r2.Errors[XlErrorChecks.xlInconsistentFormula].Ignore = true;
            }

            sheet.Range["A3"].Select();
        }
         Range GetOptionStrikePrices(string stockSymbol, int iYear, int iMonth, Range destRange)
        {
            string url = $"http://finance.yahoo.com/q/op?s={stockSymbol}&m={iYear}-{iMonth:00}"; //ex:  "URL;http://finance.yahoo.com/q/op?s=MSFT&m=2018-01"
            QueryTable webQuery = sheet.QueryTables.Add("URL;" + url, destRange);
            webQuery.WebSelectionType = XlWebSelectionType.xlSpecifiedTables;
            webQuery.WebTables = webQueryTables;  
      
            webQuery.WebFormatting = XlWebFormatting.xlWebFormattingRTF;
            webQuery.BackgroundQuery = false;
            webQuery.AdjustColumnWidth = false;
            webQuery.Refresh();

            int firstDataRow = destRange.Row + 3;   //first 3 rows are date info
            Range r = destRange.Offset[2, 0]; 
            if (sheet.Cells[firstDataRow, symbolColumn].value == null)
            {
                sheet.Cells[firstDataRow, symbolColumn].value = $"ERROR: Expected data from web query here- url: ${url}";
                return r;
            }

            r = r.End[XlDirection.xlDown];
            int lastRow = r.Row;

            fillInFormulas(firstDataRow, lastRow);
            return r;
        }

         void fillInFormulas(int firstRow, int lastRow)
        {
            
            const char bidColumn = 'J';
            const char askColumn = 'K';
            //=RTD("tos.rtd", , E$1, ".UPRO180119P"&$A6)
            //=RTD("tos.rtd", , E$1, ".UPRO180119P"&Strike_Price)


            string optionSymbol = sheet.Cells[firstRow, symbolColumn].value;
           
            string baseSymbol = "." + optionSymbol.Remove(1 + optionSymbol.LastIndexOf(optionLetter));
            string bidCellFormula = $"=RTD(\"tos.rtd\", , {bidColumn}$1, \"{baseSymbol}\"&Strike_Price";
            string askCellFormula = $"=RTD(\"tos.rtd\", , {askColumn}$1, \"{baseSymbol}\"&Strike_Price";
            sheet.Range[$"{bidColumn}{firstRow}:{bidColumn}{lastRow}"].Value = bidCellFormula;
            sheet.Range[$"{askColumn}{firstRow}:{askColumn}{lastRow}"].Value = askCellFormula;

        }
        
         Workbook getOpenOrCreateWorkbook(Workbook sourceWorkbook)
        {
            Workbook wb;
            string name = "optiondata " + DateTime.Now.ToString("yyyy-MMM") + ".xlsx";
            string filePath = sourceWorkbook.Path + @"\"; // @"C:\Users\Steven\mystuff\";
            
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

    }
}

/* Copyright (C) 2016 Steve Janke - All Rights Reserved
 * You may use, distribute and modify this code under the
 * terms of the GNU GENERAL PUBLIC LICENSE Version 3,
 *
 * You should have received a copy of the GNU GENERAL PUBLIC LICENSE Version 3 with
 * this file. If not, please visit : https://www.gnu.org/licenses/gpl-3.0.en.html
 */
