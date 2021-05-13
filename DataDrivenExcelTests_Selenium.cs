using NUnit.Framework;
using RestSharp;
using RestSharp.Serialization.Json;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataDrivenExcelTests
{
    [TestFixture]
    class DataDrivenExcelTests_Selenium
    {
        const string zippopotamusUrl = "https://api.zippopotam.us/";
        const string excelFilePath = @"..\..\..\zipCodes.xlsx";

        [TestCaseSource("LoadLocationTestData")]
        public void GetDataFromFile_ReturnsPlaceName(string countryCode, string zipCode, string city)
        {
            // Arrange
            var restClient = new RestClient(zippopotamusUrl);
            var httpRequest = new RestRequest(countryCode + "/" + zipCode);

            // Act
            var httpResponce = restClient.Execute(httpRequest);
            var location = new JsonDeserializer().Deserialize<Location>(httpResponce);

            // Assert
            StringAssert.Contains(city, location.Places[0].PlaceName);
        }
        static IEnumerable<TestCaseData> LoadLocationTestData()
        {
            using (var sheet = new SLDocument("zipCodes.xlsx"))
            {
                int endRowIndex = sheet.GetWorksheetStatistics().EndColumnIndex;
                for (int row = 2; row <= endRowIndex; row++)
                {
                    string countryCode = sheet.GetCellValueAsString(row, 1);
                    string zipCode = sheet.GetCellValueAsString(row, 2);
                    string city = sheet.GetCellValueAsString(row, 3);

                    yield return new TestCaseData(countryCode, zipCode, city);
                }
            }
        }


    }

}
                    
     