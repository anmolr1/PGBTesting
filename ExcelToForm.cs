using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PGBtesting
{
    class ExceltToForm
    {
        public DataTable ExcelToDataTable(string filename)
        {
            //Open file and return steam
            FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);

            //CreateOpenXmlReader Via ExcelReaderFactory
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            // DataSet result = excelReader.AsDataSet();
            //  excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            DataTableCollection table = result.Tables;
            DataTable resultTable = table["Sheet1"];

            return resultTable;
        }
        List<DataCollection> datacol = new List<DataCollection>();

        // Convert data from the excel file into a data table.
        public void PopulateInCollection(string filename)
        {
            DataTable table = ExcelToDataTable(filename);

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    DataCollection dtTable = new DataCollection()
                    {
                        rowNum = row,
                        columnName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    datacol.Add(dtTable);
                }
            }
        }

        //Function to match the input row number and column from the excel file.
        public string ReadData(int rowNumber, string columnName)
        {
            try
            {
                var datas = datacol.Where(x => x.columnName == columnName && x.rowNum == rowNumber).SingleOrDefault().colValue;
                return datas.ToString();
            }
            catch (Exception e)
            {
                return null;
            }

        }
    }
    public class DataCollection
    {
        public int rowNum { get; set; }
        public string columnName { get; set; }
        public string colValue { get; set; }
    }
}


