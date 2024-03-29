using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace AudioCloud_ChromeDriverTest
{
    [TestClass]
    public class ChromeDriverTest
    {
        private ChromeDriver _driver;
        public static string url = "http://54.161.251.210:8080";
        private static int rowIndex = 1;

        [TestInitialize]
        public void ChromeDriverInitialize()
        {
            // Initialize Chrome driver 
            var options = new ChromeOptions
            {
                PageLoadStrategy = PageLoadStrategy.Normal
            };
            _driver = new ChromeDriver(options);
        }


        [TestCleanup]
        public void ChromeDriverCleanup()
        {
            _driver.Quit();
        }

        // Update Test Status in Excel
        private void UpdateTestResultInExcel(string sheetName, int row, int column, string result)
        {
            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];
                worksheet.Cells[row, column].Value = result;
                excelPackage.Save();
            }
        }

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

        public static IEnumerable<object[]> ReadTestDataSearchFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["SearchDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var query = worksheet.Cells[row, 1].Value?.ToString();
                    testData.Add(new object[] { query });
                }
            }

            return testData;
        }

        [TestMethod]
        [DataRow("http://54.161.251.210:8080/home")]
        public void URLTest(string url)
        {
            _driver.Url = url;
            Assert.AreEqual(url, _driver.Url);
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataLoginFromExcel), DynamicDataSourceType.Method)]
        public async Task LoginTest(string account, string password)
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys(account);
            _driver.FindElement(By.Id("password")).SendKeys(password);

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            try
            {

                IAlert alert = _driver.SwitchTo().Alert();
                alert.Accept();
                await Task.Delay(TimeSpan.FromSeconds(5));

                Assert.IsTrue(_driver.Url.Contains("/home"), "Login failed. User was not redirected to the home page.");
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, $"Failed + {ex}");
                throw ex;
            }
            catch (ArgumentNullException e)
            {
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, $"Failed + {e}");
                throw e;
            }
            catch(Exception e)
            {
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, $"Failed + {e}");
                throw e;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataSignUpFromExcel), DynamicDataSourceType.Method)]
        public async Task SignUpTest(string account, string password, string repassword, string email)
        {
            _driver.Url = url + "/register";

            await Task.Delay(TimeSpan.FromSeconds(5));

            _driver.FindElement(By.XPath("//*[@id=\"account\"]")).SendKeys(account);
            _driver.FindElement(By.XPath("//*[@id=\"password\"]")).SendKeys(password);
            _driver.FindElement(By.XPath("//*[@id=\"signup-form\"]/div[3]/input")).SendKeys(repassword);
            _driver.FindElement(By.XPath("//*[@id=\"email\"]")).SendKeys(email);

            _driver.FindElement(By.XPath("//*[@id=\"submit\"]")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));
            try
            {
                IAlert alert = _driver.SwitchTo().Alert();

                if(alert == null)
                {
                    UpdateTestResultInExcel("SignUpDataTest", rowIndex, 5, "Failed");
                    return;
                }

                var alertText = alert.Text.ToString();
                await Task.Delay(TimeSpan.FromSeconds(5));
                var ifTrue = alertText.Contains("Đăng ký thành công");

                if (ifTrue)
                {
                    UpdateTestResultInExcel("SignUpDataTest", rowIndex, 5, "Passed");
                    return;
                }
                else
                {
                    UpdateTestResultInExcel("SignUpDataTest", rowIndex, 5, "Failed");
                    return;
                }
            }
            catch (Exception ex)
            {
                UpdateTestResultInExcel("SignUpDataTest", rowIndex, 5, $"Failed + {ex}");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataSearchFromExcel), DynamicDataSourceType.Method)]
        public async Task SearchSongs(string query)
        {
            _driver.Url = url + "/home";
            var searchInput = _driver.FindElement(By.CssSelector("input.form-control.pl-3"));
            searchInput.SendKeys(query);
            searchInput.Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/search"), "Search failed");
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, $"Failed + {ex}");
                throw ex;
            }
            catch(Exception e)
            {
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, $"Failed + {e}");
                throw e;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        public async Task PlayAudio()
        {
            _driver.Url = url + "/home";
            _driver.FindElement(By.CssSelector("button.btn-95")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            var aplayer = _driver.FindElement(By.Id("aplayer"));
            Assert.IsTrue(aplayer.Displayed, "Failed");
        }
    }
}
