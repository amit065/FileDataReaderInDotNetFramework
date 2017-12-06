using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation;
using System.IO;

namespace Tavisca.CityTable.Manipulation.Core.Test
{
    [TestClass]
    public class ExcelReaderTest
    {
        [TestMethod]
        public void Reader_Should_Give_File_When_Found()
        {
            
            string filePath = @"C:\Users\aprakash\Desktop\CLTS_Cities_Data.xlsx";
            ExcelReader excelReader = new ExcelReader(filePath);
            var cities = excelReader.Read(filePath);
            Assert.IsNotNull(cities);

        }

        [TestMethod]
        public void Reader_Should_Throw_Exception_When_File_Not_Found()
        {
            string filePath = @"C:\Users\aprakash\Desktop\CLTS_Cities.xlsx";
            ExcelReader excelReader = new ExcelReader(filePath);
            var exception = Assert.ThrowsException<Exception>(() => excelReader.Read(filePath));
            Assert.AreEqual("File Not Found", exception.Message);
        }
    }
}
