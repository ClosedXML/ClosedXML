#if !APPVEYOR && NETFRAMEWORK
using ClosedXML.Excel;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.OleDb
{
    [TestFixture]
    public class OleDbTests
    {
        [Test]
        public void TestOleDbValues()
        {
            using (var tf = new TemporaryFile(CreateTestFile()))
            {
                Console.Write("Using temporary file\t{0}", tf.Path);
                var connectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';", tf.Path);
                using (var connection = new OleDbConnection(connectionString))
                {
                    // Install driver from https://www.microsoft.com/en-za/download/details.aspx?id=13255 if required
                    // Also check that test runner is running under correct architecture:
                    connection.Open();
                    using (var command = new OleDbCommand("select * from [Sheet1$]", connection))
                    using (var dataAdapter = new OleDbDataAdapter())
                    {
                        dataAdapter.SelectCommand = command;
                        var dt = new DataTable();
                        dataAdapter.Fill(dt);

                        Assert.AreEqual("Base", dt.Columns[0].ColumnName);
                        Assert.AreEqual("Ref", dt.Columns[1].ColumnName);

                        Assert.AreEqual(2, dt.Rows.Count);

                        Assert.AreEqual(42, dt.Rows.Cast<DataRow>().First()[0]);
                        Assert.AreEqual(42, dt.Rows.Cast<DataRow>().First()[1]);

                        Assert.AreEqual(41, dt.Rows.Cast<DataRow>().Last()[0]);
                        Assert.AreEqual(41, dt.Rows.Cast<DataRow>().Last()[1]);
                    }

                    using (var command = new OleDbCommand("select * from [Sheet2$]", connection))
                    using (var dataAdapter = new OleDbDataAdapter())
                    {
                        dataAdapter.SelectCommand = command;
                        var dt = new DataTable();
                        dataAdapter.Fill(dt);

                        Assert.AreEqual("Ref1", dt.Columns[0].ColumnName);
                        Assert.AreEqual("Ref2", dt.Columns[1].ColumnName);
                        Assert.AreEqual("Sum", dt.Columns[2].ColumnName);
                        Assert.AreEqual("SumRef", dt.Columns[3].ColumnName);

                        var expected = new Dictionary<string, double>()
                        {
                            {"Ref1", 42 },
                            {"Ref2", 41 },
                            {"Sum", 83 },
                            {"SumRef", 83 },
                        };

                        foreach (var col in dt.Columns.Cast<DataColumn>())
                            foreach (var row in dt.Rows.Cast<DataRow>())
                            {
                                Assert.AreEqual(expected[col.ColumnName], row[col]);
                            }

                        Assert.AreEqual(2, dt.Rows.Count);
                    }

                    connection.Close();
                }
            }
        }

        private string CreateTestFile()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = "Base";
                ws.Cell("B1").Value = "Ref";

                ws.Cell("A2").Value = 42;
                ws.Cell("A3").Value = 41;

                ws.Cell("B2").FormulaA1 = "=A2";
                ws.Cell("B3").FormulaA1 = "=A3";

                ws = wb.AddWorksheet("Sheet2");
                ws.Cell("A1").Value = "Ref1";
                ws.Cell("B1").Value = "Ref2";
                ws.Cell("C1").Value = "Sum";
                ws.Cell("D1").Value = "SumRef";

                ws.Cell("A2").FormulaA1 = "=Sheet1!A2";
                ws.Cell("B2").FormulaA1 = "=Sheet1!A3";
                ws.Cell("C2").FormulaA1 = "=SUM(A2:B2)";
                ws.Cell("D2").FormulaA1 = "=SUM(Sheet1!A2:Sheet1!A3)";

                ws.Cell("A3").FormulaA1 = "=Sheet1!B2";
                ws.Cell("B3").FormulaA1 = "=Sheet1!B3";
                ws.Cell("C3").FormulaA1 = "=SUM(A3:B3)";
                ws.Cell("D3").FormulaA1 = "=SUM(Sheet1!B2:Sheet1!B3)";

                var path = Path.ChangeExtension(Path.GetTempFileName(), "xlsx");
                wb.SaveAs(path, true, true);

                return path;
            }
        }
    }
}
#endif
