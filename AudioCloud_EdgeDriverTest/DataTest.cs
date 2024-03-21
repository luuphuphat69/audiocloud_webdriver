using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AudioCloud_EdgeDriverTest
{
    public class DataTest
    {
        // Login data
        public static IEnumerable<object[]> ReadTestDataLoginFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["LoginDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var account = worksheet.Cells[row, 1].Value?.ToString();
                    var password = worksheet.Cells[row, 2].Value?.ToString();
                    testData.Add(new object[] { account, password });
                }
            }

            return testData;
        }

        // SignUp data
        public static IEnumerable<object[]> ReadTestDataSignUpFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["SignUpDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var account = worksheet.Cells[row, 1].Value?.ToString();
                    var password = worksheet.Cells[row, 2].Value?.ToString();
                    var repassword = worksheet.Cells[row, 3].Value?.ToString();
                    var email = worksheet.Cells[row, 4].Value?.ToString();
                    testData.Add(new object[] { account, password, repassword, email });
                }
            }

            return testData;
        }

        public static IEnumerable<object[]>  ReadTestDataSearchFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["SearchDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var query = worksheet.Cells[row, 1].Value?.ToString();
                    testData.Add(new object[] { query});
                }
            }

            return testData;
        }
    }
}
