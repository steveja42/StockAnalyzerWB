using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
//a change for practicegit - a second change for practicegit
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
        const int symbolColumn = 1;
        const int priceColumn = 2;

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {

            ILoggerFactory loggerFactory = new LoggerFactory()
            .AddDebug();
            logger = loggerFactory.CreateLogger($"XL-Addin-{this.Name}");

            logger.LogInformation("**************** starting *********************");

            workbook = (Workbook)this.Parent;
            app = workbook.Application;
            makeLabelsTransparentHack();

        }
        private void makeLabelsTransparentHack()
        {
            var color = System.Drawing.ColorTranslator.FromWin32((int)this.Cells[1, 1].Interior.Color);
            label0.BackColor = color; //   System.Drawing.Color.FromArgb((int) color);
            label1.BackColor = color; //   System.Drawing.Color.FromArgb((int) color);
            label2.BackColor = color; //   System.Drawing.Color.FromArgb((int) color);
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
            logger.LogInformation("**************** shutting down *********************");

        }


        private void buttonGetPutOptions_Click(object sender, EventArgs e)
        {
            var OptionDates = new List<string>();
            if (listBoxOptionDates.SelectedItems.Count == 0)
                return;
            foreach (string s in listBoxOptionDates.SelectedItems)
            {
                OptionDates.Add(s);
            }

            new PutOptionSheet(workbook, textBoxStockSymbol.Text).makeSheet(OptionDates);
        }

        private void buttonGetCallOptions_Click(object sender, EventArgs e)
        {
            var OptionDates = new List<string>();
            if (listBoxOptionDates.SelectedItems.Count == 0)
                return;
            foreach (string s in listBoxOptionDates.SelectedItems)
            {
                OptionDates.Add(s);
            }
            new CallOptionSheet(workbook, textBoxStockSymbol.Text).makeSheet(OptionDates);
        }




        Stock? getStockDataFromSelectedRow()
        {


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
            this.label0.Click += new System.EventHandler(this.label0_Click);
            this.buttonGetPutOptions.Click += new System.EventHandler(this.buttonGetPutOptions_Click);
            this.buttonGetCallOptions.Click += new System.EventHandler(this.buttonGetCallOptions_Click);
            this.textBoxStockSymbol.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxStockSymbol_KeyDown);
            this.textBoxStockSymbol.Leave += new System.EventHandler(this.textBox1_Leave);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Sheet1_SelectionChange);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }


        #endregion

        private void showOptionDates()
        {



            listBoxOptionDates.Items.Clear();
            string symbol = textBoxStockSymbol.Text;
            if (symbol == null || symbol =="")
                return;
            List<string> ExpDates = TDAmeritrade.getOptionDates(symbol).Result;

            if (ExpDates == null || ExpDates.Count == 0)
                return;
            //System.Windows.Forms.ListBox lb = listBoxOptionDates;
            listBoxOptionDates.BeginUpdate();
            foreach (var item in ExpDates)
            {
                listBoxOptionDates.Items.Add(item);

            }
            listBoxOptionDates.EndUpdate();

            listBoxOptionDates.SelectionMode = SelectionMode.MultiExtended;
            listBoxOptionDates.SetSelected(0, true);
            listBoxOptionDates.SetSelected(Math.Min(1, listBoxOptionDates.Items.Count), true);
            listBoxOptionDates.SetSelected(Math.Min(listBoxOptionDates.Items.Count - 2, listBoxOptionDates.Items.Count), true);
            listBoxOptionDates.SetSelected(listBoxOptionDates.Items.Count - 1, true);

        }


        private void textBoxStockSymbol_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                showOptionDates();
            }

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            showOptionDates();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Sheet1_SelectionChange(Range Target)
        {

            if (Target.Column == symbolColumn && Target.Row > 4)
            {
                string stockSymbol = this.Cells[Target.Row, Target.Column].Value;
                if (stockSymbol != null)
                {
                    textBoxStockSymbol.Text = stockSymbol;
                    showOptionDates();
                }
            }


            double? lastPrice;
            int row = app.Selection.row;

            try
            {
                lastPrice = this.Cells[row, priceColumn].Value;
            }
            catch
            {
                lastPrice = null;
            }

            
        }

        private void label0_Click(object sender, EventArgs e)
        {

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
