using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tavisca.CityTable.Manipulation.Core.ExcelFileManipulation
{
    public class ExcelManipulator
    {

        public  string GetFullTextSearch(string cityName, string IataCityCode)
        {
            string textSearch = IataCityCode != null ? IataCityCode : string.Empty;

            var CityNameloweCase = cityName.ToLower();

            foreach (var part in CityNameloweCase.Split(' '))
            {
                if (part.Length > 3)
                {

                    for (int i = 3; i <= part.Length; i++)
                    {
                        textSearch += " " + part.Substring(0, i);
                    }
                }
                else
                {
                    textSearch += " " + part;
                }
            }

            return textSearch;
        }
    }
}
