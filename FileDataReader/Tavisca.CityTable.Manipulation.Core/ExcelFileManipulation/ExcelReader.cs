using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Tavisca.CityTable.Manipulation.Core.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation
{
    public class ExcelReader
    {
        private string _filePath;

        public ExcelReader(string filepath)
        {
           this._filePath = filepath;
        }

        public List<City> ReadCityFromExcelFile(string filePath)
        {
            Excel.Application xlApp = new Excel.Application();
            try
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int row = xlRange.Rows.Count;
                int columnl = xlRange.Columns.Count;

                List<City> cities = new List<City>();
                ExcelManipulator excelManipulator = new ExcelManipulator();

                for (int i = 1; i <= row; i++)
                {
                    cities.Add(new City
                    {
                        CityName = Convert.ToString((xlRange.Cells[i, 1] as Excel.Range).Value2),
                        StateCode = Convert.ToString((xlRange.Cells[i, 2] as Excel.Range).Value2),
                        CountryCode = Convert.ToString((xlRange.Cells[i, 3] as Excel.Range).Value2),
                        Latitude = Convert.ToString((xlRange.Cells[i, 4] as Excel.Range).Value2),
                        Longitude = Convert.ToString((xlRange.Cells[i, 5] as Excel.Range).Value2),
                        IsEnabled = Convert.ToString((xlRange.Cells[i, 6] as Excel.Range).Value2),
                        IataCityCode = Convert.ToString((xlRange.Cells[i, 7] as Excel.Range).Value2),
                        FullTextColumn = (i == 1 ? "FullTextSearch" : excelManipulator.GetFullTextSearch((string)(xlRange.Cells[i, 1] as Excel.Range).Value2, Convert.ToString((xlRange.Cells[i, 7] as Excel.Range).Value2)))
                    });
                }
                return cities;
            }
            catch
            {
                throw new Exception("File Not Found");
            }
            finally
            {
                // Quit Excel application
                xlApp.Quit();

                // Release COM objects ()
                if (xlApp != null)
                Marshal.ReleaseComObject(xlApp);
                
                // Empty variables
                xlApp = null;
 
                // Force garbage collector cleaning
                GC.Collect();
            }


        }

    }
}
