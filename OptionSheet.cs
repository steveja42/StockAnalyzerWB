using System;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StockAnalyzerWB
{
    class CallOptionSheet : OptionSheet
    {
        public CallOptionSheet(Workbook sourceWorkbook, string stockIn) : base(sourceWorkbook, stockIn)
        {
            optionType = "CALL";
            templateSheetName = "callTemplate";
            optionLetter = "C";
        }
    }

    class PutOptionSheet:OptionSheet
    {
        public PutOptionSheet(Workbook sourceWorkbook, string stockIn) : base(sourceWorkbook, stockIn)
        {
            optionType = "PUT";
            templateSheetName = "putTemplate";
            optionLetter = "P";
        }
       
    }

    class OptionSheet
    {
        enum OptionType { PUT, CALL }
        enum TemplateSheet { putTemplate, callTemplate };
        enum OptionLetter { P, C }
        enum SheetColumn { strike=1,symbol,last,chg,bid,ask};
        enum SheetRows { header=1,underlying,dataStart};
        protected string optionType;
        protected string templateSheetName;
        protected string optionLetter;
        OptionChain ochain;

        const int UnderlyingRow = 2;
        const int FirstDataRow = 3;

        Workbook outputWorkbook;
         Microsoft.Office.Interop.Excel.Application app;
         Worksheet sheet;
         Workbook sourceWorkbook;
        string stockSymbol;
        static ILogger logger;

        protected OptionSheet(Workbook sourceWorkbook, string stockIn)
        {
            if (logger == null)
            {
                ILoggerFactory loggerFactory = new LoggerFactory()
                .AddDebug();
                logger = loggerFactory.CreateLogger($"XL-Addin-OptionSheet");

            }

            stockSymbol = stockIn;
            this.sourceWorkbook = sourceWorkbook;
        }
        //makeSheet - fills in a new sheet with option prices for the provided option dates
        public void makeSheet(List<string> OptionDates)
        {
            if (OptionDates.Count == 0)
                return;
            try
            {
                ochain = TDAmeritrade.GetOptionChain(stockSymbol, $"{optionType}").Result;
            }
            catch (Exception e)
            {
                string sz = $"ERROR: {e.Message} ";
                if (e.InnerException != null)
                    sz+= e.InnerException.Message;
                MessageBox.Show(sz, "Error");
                //sheet.Cells[SheetRows.dataStart, SheetColumn.symbol].value = sz; //x
                return;
            }
            if (ochain == null)
                return;

            app = sourceWorkbook.Application;
            outputWorkbook = openOrCreateWorkbook(sourceWorkbook);
            outputWorkbook.Activate();

            sourceWorkbook.Sheets[templateSheetName].Copy(outputWorkbook.Sheets[1]);// .ActiveSheet);
            sheet = outputWorkbook.Sheets[templateSheetName];
            try
            {
                sheet.Name = stockSymbol + $"-{optionType}-" + DateTime.Now.ToString("d.h.m");

            }
            catch
            {
                sheet.Name = stockSymbol + $"-{optionType}-" + DateTime.Now.ToString("d.h.m.s");

            }
            sheet.Cells[SheetRows.underlying, SheetColumn.symbol].value = stockSymbol;
            sheet.Cells[SheetRows.underlying, SheetColumn.last].Formula = @"= RTD(""tos.rtd"", , ""LAST"", Symbol)";
            int Row = FirstDataRow;

            foreach (string s in OptionDates)
            {
                Row = AddOptionData(stockSymbol, s, Row) + 2;

            }
            /*
            int iYear = DateTime.Today.Year;
            int iMonth = DateTime.Today.Month;


            Row = AddOptionData(stockSymbol, iYear + 2, 1, Row);
            Row += 2;
            Row = AddOptionData(stockSymbol, iYear + 1, 1, Row);
            Row += 2;

            if (iMonth != 12)  //january is already done above
            {
                Row = AddOptionData(stockSymbol, (iMonth == 12 ? iYear + 1 : iYear), (iMonth + 1) % 12, Row);
                Row += 2;
            }
            Row = AddOptionData(stockSymbol, iYear, iMonth, Row);

    */
            foreach (Range r2 in sheet.Range["N5:Y300"].Cells)
            {
                r2.Errors[XlErrorChecks.xlInconsistentFormula].Ignore = true;
            }

            sheet.Range["A3"].Select();
        }

     
        //https://api.tdameritrade.com/v1/marketdata/chains GET /v1/marketdata/chains?apikey=SIRSNEEZ%40AMER.OAUTHAP&symbol=TSLA&contractType=PUT&strikeCount=1&optionType=S


//iRow is where to put the data

        int AddOptionData(string stockSymbol, string date, int iRow)
        {
            int firstDataRow = iRow + 1;   //first  row is Epiration Date info
            //string date = $"{iYear}-{iMonth:00}";
            ExpirationDate ExpDateItem = null;
            string ExpDate = null;
            
            Dictionary<string, ExpirationDate> ExpDates = optionType == "PUT" ? ochain.putExpDateMap : ochain.callExpDateMap;

            
            foreach (var item in ExpDates)
            {
                if (0 == string.Compare(item.Key, 0, date,0,date.Length)) {
                    ExpDate = item.Key;
                    ExpDateItem = item.Value;
                    break;
                }
            }

            if (ExpDateItem == null)
            {
                sheet.Cells[iRow, 2].Value = $"Expiration Date not found: {ExpDate}";
                return iRow;
            }

            string SymbolPrefix = $".{stockSymbol}" + ExpDate.Substring(2, 2) + ExpDate.Substring(5, 2) + ExpDate.Substring(8, 2) + optionLetter;

            sheet.Cells[iRow, 1].Value = optionType;
            sheet.Cells[iRow, 2].Value = ExpDate;
            sheet.Cells[iRow, 2].Font.Bold = true;
            sheet.Cells[iRow, 2].Font.Size = 12;

            foreach (var item in ExpDateItem)
            {
                ++iRow;
                sheet.Cells[iRow, SheetColumn.strike].Value = item.Key;
                string strikePrice = item.Key.Remove(item.Key.LastIndexOf('.'));
                sheet.Cells[iRow, SheetColumn.symbol].Value = SymbolPrefix + strikePrice;
                sheet.Cells[iRow, SheetColumn.last].Value = item.Value[0].last;
                sheet.Cells[iRow, SheetColumn.bid].Value = item.Value[0].bid;
                sheet.Cells[iRow, SheetColumn.ask].Value = item.Value[0].ask;
            }

            fillInFormulas(firstDataRow, iRow);
            return iRow;
        }

         void fillInFormulas(int firstRow, int lastRow)
        {
            
            const char bidColumn = 'J';
            const char askColumn = 'K';
            
            string bidCellFormula = $"=RTD(\"tos.rtd\", , {bidColumn}$1, Symbol";
            string askCellFormula = $"=RTD(\"tos.rtd\", , {askColumn}$1, Symbol";
            sheet.Range[$"{bidColumn}{firstRow}:{bidColumn}{lastRow}"].Value = bidCellFormula;
            sheet.Range[$"{askColumn}{firstRow}:{askColumn}{lastRow}"].Value = askCellFormula;

        }
        
        //opens workbook if it exists, else creates it
         Workbook openOrCreateWorkbook(Workbook sourceWorkbook)
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
