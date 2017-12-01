using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation;

namespace Tavisca.CityTable.Manipulation.Core.Test
{
    [TestClass]
    public class ExcelManipulatorTest
    {
        [TestMethod]
        public void ExcelManipulator_Should_Return_Full_Text_Search()
        {
            var cityName = "Mumbai";
            var IataCityCode = "MUM";
            var fullTextSearch = "MUM mum mumb mumba mumbai";
            ExcelManipulator excelManipulator = new ExcelManipulator();
            var generatedFullTextSearch = excelManipulator.GetFullTextSearch(cityName, IataCityCode);
            Assert.AreEqual(fullTextSearch, generatedFullTextSearch);
        }

        [TestMethod]
        public void ExcelManipulator_Should_Return_Full_Text_Search_For_More_Than_One_Word_CityName()
        {
            var cityName = "Santa Maria";
            var IataCityCode = "SMX";
            var fullTextSearch = "SMX san sant santa mar mari maria";
            ExcelManipulator excelManipulator = new ExcelManipulator();
            var generatedFullTextSearch = excelManipulator.GetFullTextSearch(cityName, IataCityCode);
            Assert.AreEqual(fullTextSearch, generatedFullTextSearch);
        }
    }
}
