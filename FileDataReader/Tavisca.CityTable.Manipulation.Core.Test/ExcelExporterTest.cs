using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation;
using Tavisca.CityTable.Manipulation.Core.Model;
using System.Collections.Generic;

namespace Tavisca.CityTable.Manipulation.Core.Test
{
    [TestClass]
    public class ExcelExporterTest
    {
        [TestMethod]
        public void ExcelExporter_Should_Save_File_In_Excel()
        {
            string filePath = @"C:\Users\aprakash\Desktop\CLTS_Cities_Data.xlsx";
            ExcelReader excelReader = new ExcelReader(filePath);
            List<City> cities = excelReader.ReadCityFromExcelFile(filePath);
            ExcelExporter excelExporter = new ExcelExporter();
            var result = excelExporter.ExportToExcel(cities);
            Assert.AreEqual(true, result);

        }
       
    }
}
