using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Tavisca.CityTable.Manipulation.Core.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation
{
    public class ExcelExporter
    {
        public bool ExportToExcel(List<City> cities)
        {
            // Load Excel application
            Excel.Application excel = new Excel.Application();

            // Load Excel application
            excel.Workbooks.Add();

            // Create Worksheet from active sheet
            Excel._Worksheet workSheet = excel.ActiveSheet;

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!
            try
            {
                // Creation of header cells
                workSheet.Cells[1, "A"] = "CityName";
                workSheet.Cells[1, "B"] = "StateCode";
                workSheet.Cells[1, "C"] = "CountryCode";
                workSheet.Cells[1, "D"] = "Latitude";
                workSheet.Cells[1, "E"] = "Longitude";
                workSheet.Cells[1, "F"] = "IsEnabled";
                workSheet.Cells[1, "G"] = "IataCityCode";
                workSheet.Cells[1, "H"] = "FullTextColumn";

                // Populate sheet with some real data from "city" list
                int row = 1;
                foreach (City city in cities)
                {
                    workSheet.Cells[row, "A"] = city.CityName;
                    workSheet.Cells[row, "B"] = city.StateCode;
                    workSheet.Cells[row, "C"] = city.CountryCode;
                    workSheet.Cells[row, "D"] = city.Latitude;
                    workSheet.Cells[row, "E"] = city.Longitude;
                    workSheet.Cells[row, "F"] = city.IsEnabled;
                    workSheet.Cells[row, "G"] = city.IataCityCode;
                    workSheet.Cells[row, "H"] = city.FullTextColumn;

                    row++;
                }

                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelCityData.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);
                return true;
            }
            catch 
            {
                throw new Exception("Unable to save file");
            }
            finally
            {
                // Quit Excel application
                excel.Quit();

                // Release COM objects ()
                if (excel != null)
                    Marshal.ReleaseComObject(excel);
               
                if (workSheet != null)
                    Marshal.ReleaseComObject(workSheet);

                // Empty variables
                excel = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }

        }
    }
}
