using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation;

namespace Tavisca.CityTable.Manipulation.Core.Test
{
    [TestClass]
    public class ScriptGeneratorTest
    {
        [TestMethod]
        public void Script_Generator_Should_Generate_Insert_Script_For_Cities()
        {
            string filePath = @"C:\Users\aprakash\Desktop\CLTS_Cities_Data.xlsx";
            ExcelReader excelReader = new ExcelReader(filePath);
            var cities = excelReader.Read(filePath);
            ScriptGenerator scriptGenerator = new ScriptGenerator();
            Assert.ThrowsException<Exception>(() => scriptGenerator.GenerateInsertScript(cities));

        }
    }
}
