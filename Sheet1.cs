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
    public struct Stock
    {
        public string symbol;
        public string lastPriceFormula;
    }
    public partial class Sheet1
    {
        ILogger logger;
        Workbook workbook;
        Excel.Application app;
      
    
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
            Stock? stock = getStockDataFromSelectedRow();
            if (stock != null)
               new PutOptionSheet(workbook, stock).makeSheet();
        }

        private void buttonGetCallOptions_Click(object sender, EventArgs e)
        {
            Stock? stock = getStockDataFromSelectedRow();
            if (stock != null)
                new CallOptionSheet(workbook, stock).makeSheet();
        }




        Stock? getStockDataFromSelectedRow()
        {
            const int symbolColumn = 1;
            const int priceColumn = 2;

            Stock stock;

            double? lastPrice;
            int row = app.Selection.row;
            stock.symbol = this.Cells[row, symbolColumn].Value;

            try
            {
                lastPrice = this.Cells[row, priceColumn].Value;
            }
            catch
            {
                lastPrice = null;
            }

            if (stock.symbol == null || lastPrice == null)
            {
                Error("Missing stock symbol or a last price on the current row");
                return null;
            }

            stock.lastPriceFormula = this.Cells[row, priceColumn].Formula;

            return stock;
        }

        void Error(string message)
        {
            logger.LogError(message);
            MessageBox.Show(message, "Whoops");
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



/* Copyright (C) 2016 Steve Janke - All Rights Reserved
 * You may use, distribute and modify this code under the
 * terms of the GNU GENERAL PUBLIC LICENSE Version 3,
 *
 * You should have received a copy of the GNU GENERAL PUBLIC LICENSE Version 3 with
 * this file. If not, please visit : https://www.gnu.org/licenses/gpl-3.0.en.html
 */
