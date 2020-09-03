using MNB.Models;
using System.Data;

namespace MNB.Services
{
    interface IOperatorService
    {
        object XmlToModel<T>(string xml);
        /*
        * Input: type,string (type like MNBExchangeRates, xml string like exchangeRatesXML)
        * Output: T type object (e.g. exchange rates model)
        * Deserialize an xml string to a given object
        * Object model needed
        */

        DataTable ModelToDataTable(MNBExchangeRates exchangeRatesModel, MNBCurrencies currenciesModel);
        /*
        * Input: MNBExchangeRates, MNBCurrencies (e.g. exchangeRatesModel, e.g currenciesModel)
        * Output: DataTable
        * Create a new DataTable from exchange rates and currencies
        * Load units and rates for each currency
        * Currencies used for the columns
        * Currencies with no rates will not be loaded
        */

        void DataSetToExcel(DataSet dataSet);
        /*
        * Input: DataSet (e.g. LogDatabaseDataSet)
        * Output: void
        * Create a new sheet in the active workbook
        * Import the given DataSet to this workbook generating a new table
        */
    }
}
