using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tavisca.CityTable.Manipulation.Core.Model;

namespace Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation
{
    public class ScriptGenerator
    {
        public void GenerateInsertScript(List<City> cities)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(@"C:\Users\aprakash\Desktop\CityScript.txt", false))
                {

                    for(int i=1;i<cities.Count;i++)
                    {
                        writer.WriteLine("Insert into Cities (CityName, StateCode, CountryCode, Latitude, Longitude, IsEnabled, IataCityCode, FullTextColumn) values('" + cities[i].CityName + "' , '" + cities[i].StateCode + "' , '" + cities[i].CountryCode + "' , '" + cities[i].Latitude + "' , '" + cities[i].Longitude + "' , '" + cities[i].IsEnabled + "' , '" + (cities[i].IataCityCode != null ? cities[i].IataCityCode : cities[i].IataCityCode = "NULL") + "' , '" + cities[i].FullTextColumn + "');");

                    }

                }
            }
            catch
            {

                throw new Exception("Failed to generate script");
            }

        }
    }
}
