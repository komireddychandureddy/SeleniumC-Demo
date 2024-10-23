using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NUnitAutomationFramework.Utility
{
    public class ExcelUtils
    {

        public static DataTable ReadExcelFiles(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var encoding = Encoding.GetEncoding("UTF-8");
                //IExcelDataReader es = ExcelReaderFactory.CreateReader("");

                using (var reader = ExcelReaderFactory.CreateReader(stream,
                  new ExcelReaderConfiguration() { FallbackEncoding = encoding }))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                    });

                    if (result.Tables.Count > 0)
                    {
                        return result.Tables[0];
                    }
                }
            }
            return null;
        }


        public static DataTable ReadDataExcelFiles(string filePath)
        {

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            //...
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            // IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();

            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            // DataTable dt = excelReader
            //...
            //4. DataSet - Create column names from first row
            //excelReader. = true;
            //DataSet result = excelReader.AsDataSet();

            //5. Data Reader methods
            while (excelReader.Read())
            {
                //excelReader.GetInt32(0);
            }

            //6. Free resources (IExcelDataReader is IDisposable)
            // excelReader.Close();
            return null;

        }


        public static void renameReport()
        {



            String basePath = @"C:\Users\ADMIN\test\Project\";
            Directory.SetCurrentDirectory(basePath + @"Reports");

            String reports = Directory.GetCurrentDirectory();

            String reportsArchive = reports + @"Archive";

            if (!Directory.Exists(reportsArchive))
            {

                Directory.CreateDirectory(reports);
            }

        }



        public static DataTable ExcelFileReader(string SheetName)
        {

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string path = @"C:\\Users\\ADMIN\\source\\repos\\NUnitSeleniumAutomationFramework\\NUnitAutomationFramework\\TestData\\TestData.xlsx";
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var rowCount = reader.RowCount;
            DataTable dt = new DataTable();
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (DataTableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            // var result = reader.AsDataSet();

            var tables = result.Tables.Cast<DataTable>();
            var dataTable = result.Tables[0].Columns[0].ToString();
            foreach (DataTable table in tables)
            {
                //dataGridView1.DataSource = table;
                if (table.TableName == SheetName)
                {
                    dt = table;
                    /*
                  //  DataColumn dc = new DataColumn();
                        DataColumnCollection columns = table.Columns;
                       // dt.Columns.AddRange(columns)
                        foreach (DataColumn column in columns)
                            {
                            dt.Columns.Add(column);
                            }
                         DataRowCollection rows = table.Rows;
                                foreach (DataRow row in rows)
                                {
                                    string TestcaseName = "Product 3";
                                   
                                    if (row["TestcaseId"].ToString() == TestcaseName)
                                    {
                                        // string cellvalue = row.ToString();
                                        //Console.WriteLine(celldata);
                                        //return dataRow

                                        dt.Rows.Add(row);
                                    }


                                }
                           

                     */

                }
            }


            return dt;
        }













        public static DataTable GetTestData(string SheetName, string testcaseId)
        {

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string path = @"C:\\Users\\ADMIN\\source\\repos\\NUnitSeleniumAutomationFramework\\NUnitAutomationFramework\\TestData\\TestData.xlsx";
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var rowCount = reader.RowCount;
            DataTable dt = new DataTable();
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (DataTableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            // var result = reader.AsDataSet();

            var tables = result.Tables.Cast<DataTable>();
            var dataTable = result.Tables[0].Columns[0].ToString();
            foreach (DataTable table in tables)
            {
                //dataGridView1.DataSource = table;
                if (table.TableName == SheetName)
                {
                 // dt = table.Clone();
                    dt = table.Copy();
                    dt.Clear();

                    /* DataColumnCollection columns = table.Columns;

                     foreach (DataColumn column in columns)
                         {
                         dt.Columns.Add(column);
                         }
                      */
                    foreach (DataRow row in table.Rows)
                                {
                                  
                                    if (row["TestCase_ID"].ToString() == testcaseId)
                                    {

                                        dt.NewRow();
                                        dt.ImportRow(row);
                                       // dt.Rows.Add(row);
                        }

                                }


                 

                }
            }


            return dt;
        }



    }
}
