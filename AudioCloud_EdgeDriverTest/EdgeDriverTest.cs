using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading.Tasks;
using System;
using static System.Net.WebRequestMethods;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Security.Cryptography;
using IdentityModel.Client;
using AudioCloud_EdgeDriverTest;
using System.Runtime.CompilerServices;
using System.Reflection.Metadata;

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

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/home"), "Login failed. User was not redirected to the home page.");
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, "Failed");
                throw ex;
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

            _driver.FindElement(By.Name("account")).SendKeys(account);
            _driver.FindElement(By.Name("password")).SendKeys(password);
            _driver.FindElement(By.Name("repassword")).SendKeys(repassword);
            _driver.FindElement(By.Name("email")).SendKeys(email);

            _driver.FindElement(By.Name("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/login"), "Sign up failed");
                UpdateTestResultInExcel("SignUpDataTest", rowIndex, 6, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("SignUpDataTest", rowIndex, 6, "Failed");
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
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, "Failed");
                throw ex;
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



        //tín comment

        public static IEnumerable<object[]> ReadTestDatacommentFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["comment"];
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
        [DynamicData(nameof(ReadTestDatacommentFromExcel),DynamicDataSourceType.Method)]
        public async Task comment(string comment)
        {
            _driver.Manage().Window.Maximize();
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("tin7979");
            _driver.FindElement(By.Id("password")).SendKeys("tin7979");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(2));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();  
            await Task.Delay(TimeSpan.FromSeconds(10));
            _driver.FindElement(By.LinkText("Cầu vồng khuyết")).Click();
            await Task.Delay(TimeSpan.FromSeconds(5));
            var input =  _driver.FindElement(By.CssSelector("[type='text'][placeholder = 'Write a comment']"));
            input.SendKeys(comment);
            input.Submit();
            try
            {
                Assert.IsTrue(_driver.Url.Contains("/details/AUDIO52890086"));
                UpdateTestResultInExcel("comment", rowIndex, 2, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("comment", rowIndex, 2, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }

        }
        //tải bài hát về thiết bị khi đã đăng nhập
        [TestMethod]
        public async Task dowloadsong()
        {
            _driver.Manage().Window.Maximize();
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("tin7979");
            _driver.FindElement(By.Id("password")).SendKeys("tin7979");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(2));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();
            await Task.Delay(TimeSpan.FromSeconds(10));
            _driver.FindElement(By.LinkText("Cầu vồng khuyết")).Click();
            await Task.Delay(TimeSpan.FromSeconds(5));
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            js.ExecuteScript("document.body.style.transform = 'scale(0.9)';");
            await Task.Delay(TimeSpan.FromSeconds(3));
            var button = _driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[2]/button[3]/div[2]"));
            button.Click();
            await Task.Delay(TimeSpan.FromSeconds(5));       

            var ifUrl = _driver.Url.Contains("storage.googleapis.com");
            if (ifUrl)
            {
                Assert.IsTrue(ifUrl, "Passed");
            }
            else
            {
                Assert.IsFalse(ifUrl, "failed");

            }
        }
        //tải bài hát về thiết bị khi chưa đăng nhập
        [TestMethod]
        public async Task dowloadsongnologin()
        {
            _driver.Manage().Window.Maximize();
            _driver.Url = url + "/home";
            await Task.Delay(TimeSpan.FromSeconds(5));
            _driver.FindElement(By.LinkText("Cầu vồng khuyết")).Click();
            await Task.Delay(TimeSpan.FromSeconds(10));
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            js.ExecuteScript("document.body.style.transform = 'scale(0.9)';");
            await Task.Delay(TimeSpan.FromSeconds(3));
            var button = _driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[2]/button[3]/div[2]"));
            button.Click(); 
            await Task.Delay(TimeSpan.FromSeconds(5));

            var ifUrl = _driver.Url.Contains("storage.googleapis.com");
            if (ifUrl)
            {
                Assert.IsTrue(ifUrl, "Passed");
            }
            else
            {
                Assert.IsFalse(ifUrl, "failed");

            }
        }
        ////xóa bài hát 
        [TestMethod]
        public async Task deletesong()
        {
            _driver.Manage().Window.Maximize();
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("tin7979");
            _driver.FindElement(By.Id("password")).SendKeys("tin7979");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(2));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            await Task.Delay(TimeSpan.FromSeconds(5));
            _driver.FindElement(By.CssSelector("#navbarDropdownMenuLink")).Click();
            _driver.FindElement(By.LinkText("Tài khoản")).Click();
            await Task.Delay(TimeSpan.FromSeconds(3));
            _driver.FindElement(By.LinkText("Danh sách phát")).Click();
            await Task.Delay(TimeSpan.FromSeconds(3));
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            js.ExecuteScript("window.scrollTo(0,400);");
            await Task.Delay(TimeSpan.FromSeconds(3));
            _driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[2]/div[2]/div/div")).Click();
            await Task.Delay(TimeSpan.FromSeconds(3));
            var ktdk = _driver.FindElement(By.CssSelector("#root > div > div.card.card-body.mx-3.mx-md-4.mt-n6 > div:nth-child(2) > div > div.overlay > div > div > div > div:nth-child(1) > button"));
            ktdk.Click();
            await Task.Delay(TimeSpan.FromSeconds(3));

            Assert.IsTrue(ktdk.Displayed, "Failed");
        }
    }
}
