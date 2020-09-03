using System;
using Microsoft.Office.Tools.Ribbon;
using MNB.LogDatabaseDataSetTableAdapters;
using MNB.MNBWebService;
using MNB.Models;
using MNB.Services;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.IO;

namespace MNB
{
    public partial class MNBRibbon
    {
        private MNBArfolyamServiceSoapClient _mnbService;
        private IOperatorService _operatorService;
        private LogDatabaseDataSet _logDatabaseDataSet;
        private LogTableAdapter _logTableAdapter;
        private ListObject _logListObject;
        private BindingSource _logBindingSource;

        private void MNBRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _mnbService = new MNBArfolyamServiceSoapClient();
            _operatorService = new OperatorService();
            _logDatabaseDataSet = new LogDatabaseDataSet();
            _logTableAdapter = new LogTableAdapter();
            _logTableAdapter.Fill(_logDatabaseDataSet.Log);
            _logBindingSource = new BindingSource();
        }

        private void mnbDownload_Click(object sender, RibbonControlEventArgs e)
        {
            //Check MNB service
            if (_mnbService.State.ToString() != "Created")
            {
                _mnbService = new MNBArfolyamServiceSoapClient();
                _mnbService.Open();
            }
            //Log to the database
            LogDatabaseDataSet.LogRow newLogRow = _logDatabaseDataSet.Log.NewLogRow();
            newLogRow.Név = Environment.UserName;
            newLogRow.TimeStamp = DateTime.Now;
            _logDatabaseDataSet.Log.Rows.Add(newLogRow);
            _logTableAdapter.Update(_logDatabaseDataSet.Log);

            //Get currencies xml and map it to its model
            GetCurrenciesResponseBody currenciesXML = _mnbService.GetCurrencies(new GetCurrenciesRequestBody());
            MNBCurrencies currenciesModel = (MNBCurrencies)_operatorService.XmlToModel<MNBCurrencies>(currenciesXML.GetCurrenciesResult);

            //Exchange rates query config
            GetExchangeRatesRequestBody getExchangeRatesRequestBody = new GetExchangeRatesRequestBody
            {
                startDate = "2015.01.01.",
                endDate = "2020.04.01.",
                currencyNames = string.Join(",", currenciesModel.Currencies)
            };
            //Get exchange rates xml and map it to its model
            GetExchangeRatesResponseBody exchangeRatesXML = _mnbService.GetExchangeRates(getExchangeRatesRequestBody);
            MNBExchangeRates exchangeRatesModel = (MNBExchangeRates)_operatorService.XmlToModel<MNBExchangeRates>(exchangeRatesXML.GetExchangeRatesResult);

            //Create datatable from model and import it to Excel then save it
            DataTable dataTable = _operatorService.ModelToDataTable(exchangeRatesModel, currenciesModel);
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            _operatorService.DataSetToExcel(dataSet);

            //Format worksheet and save the workbook to the user's documents
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeWorksheet.Range["A2:CV2"].NumberFormatLocal = "";
            activeWorksheet.Range["B3:CV10000"].NumberFormatLocal = "0";
            activeWorksheet.Range["A3:A10000"].NumberFormatLocal = "éééé\\.hh\\.nn\\.";
            string savePath = Directory.GetCurrentDirectory() + "\\arfolyam-letoltes.xlsx";
            Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(savePath);

            _mnbService.Close();
        }

        private void mnbLog_Click(object sender, RibbonControlEventArgs e)
        {
            //Create an extended worksheet
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Sheets sheets = activeWorkBook.Sheets;
            Excel.Worksheet logWorkSheet = sheets.Add();
            Worksheet extendedLogWorksheet = Globals.Factory.GetVstoObject(logWorkSheet);
            Excel.Range cell = extendedLogWorksheet.Range["$A$1"];
            _logListObject = extendedLogWorksheet.Controls.AddListObject(cell, "logList");
            
            //Bind the database
            _logBindingSource.DataSource = _logDatabaseDataSet.Log;
            _logListObject.AutoSetDataBoundColumnHeaders = true;
            _logListObject.SetDataBinding(
                _logBindingSource, "", "Név", "TimeStamp",
                "Indoklás");

            //Set columns read only
            extendedLogWorksheet.Range["A2:B10000"].Locked = true;
        }

        private void mnbLogSave_Click(object sender, RibbonControlEventArgs e)
        {
            _logTableAdapter.Update(_logDatabaseDataSet.Log);
        }
    }
}
