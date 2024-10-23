using AngleSharp.Common;
using ExcelDataReader;
using NUnit.Framework;
using NUnitAutomationFramework.Base;
using NUnitAutomationFramework.Pages;
using NUnitAutomationFramework.Utility;
using System.Collections.Generic;
using System;
using System.Data;
using static System.IO.Directory;
using NUnit.Framework.Internal.Execution;
using System.Reflection;

namespace NUnitAutomationFramework.TestSuites
{
    [Parallelizable(ParallelScope.Children)]
    public class RegressionDataDrivenCopy : BaseSetup
    {
        [TestCaseSource(nameof(GetData))]
        [Test, Category("ExciseTaxTestData"), Category("ExciseTax"), Category("Regression"), Category("Smoke")]
        public void TC001_OpenTabData(object[] data)
        {
            string sheetName = "";
            //var properties = "null";
            string? testcase = TestContext.CurrentContext.Test.MethodName;
            string testcase_id = testcase.Substring(0,testcase.IndexOf('_'));
             sheetName = (string)TestContext.CurrentContext.Test.Properties.Get("Category").ToString();
            
                        var categories = TestContext.CurrentContext.Test?.Properties["Category"];
                        //IList<string> categories = (IList<string>)TestContext.CurrentContext.Test.Properties["Category"];

                        foreach (var category in categories)
                        {
                            string cat= category.ToString();

                            if (cat.Contains("TestData"))
                            {
                                sheetName = cat.Replace("TestData", "");
                                break;
                            }
                        }

           /* var s = data;
            String valll = data["Column2"];*/



                  DataTable dt2 = ExcelUtils.GetTestData(sheetName, testcase_id);

            //  DataTable dt = ExcelUtils.ExcelFileReader(sheetName);

            string testdata = ReadTestData.GetTestData(testcase, "TestData");
            extent_test.Value.Info("Testdata is : " +testdata);
           /* String basePath = @"C:\Users\ADMIN\test\Project\";
            Directory.SetCurrentDirectory(basePath +@"Reports");

            String reports = Directory.GetCurrentDirectory();
            Console.WriteLine(reports);*/

            string user = ReadUsers.UserList("Registered_User");
            extent_test.Value.Info("Testdata is : " + user);
         
            HomePage page = new(GetDriver(), extent_test.Value);
            page.OpenTab();
            extent_test.Value.Pass("Open Tab Testcase is passed");
           // Assert.Fail();
        }

        [Test, Category("RegressionData")]
        //[TestCaseSource(nameof(LoadTestDataFromExcel))]
       // [DatapointSource(nameof(TestData))]
       // [DynamicData(nameof(AdditionData))]
        public void TC002_MouseOverData()
        {
            //To get testdata from xml file
            string? testcase = TestContext.CurrentContext.Test.MethodName;

            var properties = TestContext.CurrentContext.Test.Properties["Category"];
            foreach (var property in properties)
            {
                Console.WriteLine(property);
            }
            //string testdata = ReadTestData.GetTestData(testcase, "TestData");
            //extent_test.Value.Info("Testdata is : " + testdata);



           /* //To get user from testdata file
            string username = ReadTestData.GetTestData(testcase, "Username");
            extent_test.Value.Info("Username from xml file is : " + username);

            // To get User from UserList file
            string user = ReadUsers.UserList("Registered_User");
            extent_test.Value.Info("user is : " + user);

            HomePage page = new(GetDriver(), extent_test.Value);
            page.MouseOver();
            extent_test.Value.Pass("MouseOver Testcase is passed");*/
        }


        public static IEnumerable<object[]> GetData()
        {
            string? testcase = TestContext.CurrentContext.Test.MethodName;
            string testcase_id = testcase.Substring(0, testcase.IndexOf('_'));
            var sheetName = (string)TestContext.CurrentContext.Test.Properties.Get("Category");

            var dataTable =  ExcelUtils.GetTestData(sheetName, testcase_id);
            yield return new object[] { dataTable };
        }

        public static IEnumerable<object[]> TestData
        {


            get
            {
                return new[]
                {
                       new object[] { 1, 1, 2 },
                    new object[] { 2, 2, 4 },
                    new object[] { 3, 3, 6 },
                    new object[] { 0, 0, 1 }, // The test run with this row fails
                };
            }
        }


    }
}
