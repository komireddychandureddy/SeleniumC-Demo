using ExcelDataReader;
using NUnit.Framework;
using NUnitAutomationFramework.Base;
using NUnitAutomationFramework.Pages;
using NUnitAutomationFramework.Utility;
using static System.IO.Directory;

namespace NUnitAutomationFramework.TestSuites
{
    [Parallelizable(ParallelScope.Children)]
    public class Regression : BaseSetup
    {
        [Test, Category("Regression")]
        public void TC001_OpenTab()
        {


            string? testcase = TestContext.CurrentContext.Test.MethodName;
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

        [Test, Category("Regression")]
        //[TestCaseSource(nameof(LoadTestDataFromExcel))]
       // [DatapointSource(nameof(TestData))]
       // [DynamicData(nameof(AdditionData))]
        public void TC002_MouseOver()
        {
            //To get testdata from xml file
            string? testcase = TestContext.CurrentContext.Test.MethodName;
            string testdata = ReadTestData.GetTestData(testcase, "TestData");
            extent_test.Value.Info("Testdata is : " + testdata);

            //To get user from testdata file
            string username = ReadTestData.GetTestData(testcase, "Username");
            extent_test.Value.Info("Username from xml file is : " + username);

            // To get User from UserList file
            string user = ReadUsers.UserList("Registered_User");
            extent_test.Value.Info("user is : " + user);

            HomePage page = new(GetDriver(), extent_test.Value);
            page.MouseOver();
            extent_test.Value.Pass("MouseOver Testcase is passed");
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

        /*static IEnumerable<TestCaseData> LoadTestDataFromExcel()
        {
           // using (var sheet = new SLDocument("../../../ZippopotamousTestData.xlsx"))
            {



                int endRowIndex = sheet.GetWorksheetStatistics().EndRowIndex;
                for (int row = 2; row <= endRowIndex; row++)
                {
                    string countryCode = sheet.GetCellValueAsString(row, 1);
                    string zipCode = sheet.GetCellValueAsString(row, 2);
                    string expectedPlace = sheet.GetCellValueAsString(row, 3);
                    string expectedState = sheet.GetCellValueAsString(row, 4);
                    yield return new TestCaseData(countryCode, zipCode, expectedPlace, expectedState);
                }
            }
        }*/


    }
}
