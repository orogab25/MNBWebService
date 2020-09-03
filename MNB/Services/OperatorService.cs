using MNB.Models;
using System.Data;
using System.IO;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace MNB.Services
{
    class OperatorService : IOperatorService
    {
        public object XmlToModel<T>(string xml)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            T model;

            using (TextReader reader = new StringReader(xml))
            {
                model = (T)serializer.Deserialize(reader);
            }

            return model;
        }

        public DataTable ModelToDataTable(MNBExchangeRates exchangeRatesModel, MNBCurrencies currenciesModel)
        {
            DataTable exchangeRatesDataTable = new DataTable("exchangeRatesDataTable");
            exchangeRatesDataTable.Clear();

            //Columns
            exchangeRatesDataTable.Columns.Add("Dátum/ISO");

            foreach (string currency in currenciesModel.Currencies)
            {
                exchangeRatesDataTable.Columns.Add(currency);
            }

            //First row (units)
            DataRow FirstRow = exchangeRatesDataTable.NewRow();
            FirstRow["Dátum/ISO"] = "Egység";

            foreach (MNBExchangeRatesDayRate rate in exchangeRatesModel.Day[0].Rate)
            {
                FirstRow[rate.curr] = rate.unit;
            }
            exchangeRatesDataTable.Rows.Add(FirstRow);

            //Rate rows
            foreach (MNBExchangeRatesDay day in exchangeRatesModel.Day)
            {
                DataRow RateRow = exchangeRatesDataTable.NewRow();
                RateRow["Dátum/ISO"] = day.date.ToShortDateString();

                foreach (MNBExchangeRatesDayRate rate in day.Rate)
                {
                    RateRow[rate.curr] = rate.Value;
                }
                exchangeRatesDataTable.Rows.Add(RateRow);
            }

            return exchangeRatesDataTable;
        }

        public void DataSetToExcel(DataSet dataSet)
        {
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Sheets sheets = activeWorkBook.Sheets;
            sheets.Add();
            activeWorkBook.XmlMaps.Add(dataSet.GetXmlSchema(), dataSet.DataSetName);
            activeWorkBook.XmlImportXml(dataSet.GetXml(), out _, true, "$A1");
            activeWorkBook.RefreshAll();
        }
    }
}
