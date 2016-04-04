using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Collections.Generic;
using ExcelDemo;
using ExcelDemo.Spreadsheet;
using ExcelDemoTest.SpreadsheetMock;
using ExcelDemo.TableContainer;
using System.Reflection;
using System.Xml.Linq;
using System.Linq;
using ExcelDemo.TableContainer.ListTableContainer;

namespace ExcelDemoTest
{
    [TestClass]
    public class CExcelDemoTest
    {
        private CTableContainer fileCon;
        private CTableContainer addFileCon;
        private CTableContainer assertFileCon;

        private CTableConOpener opener;
        private CSpreadsheetAdapterMock spreadSheetMock;
        private Thread mainThread;

        [TestInitialize]
        public void setUp()
        {
            initFileCon();
            opener = new CTableConOpener(fileCon);
            spreadSheetMock = new CSpreadsheetAdapterMock();
            opener.openFile(spreadSheetMock.mock.Object);
            mainThread = new Thread(() => opener.waitForClosing());
            mainThread.Start();

            initAddFileCon();
            initAssertFileCon();
        }

        private void initFileCon()
        {
            string[,] standardTable = new string[,]
            {
                { "col1", "col2", "col3", "col4", "col5" },
            {"val1.1", "val2.1", "val3.1", "val4.1", "val5.1" },
            {"val1.2", "val2.2", "val3.2", "val4.2", "val5.2"  }
            };

            string[,] numTable = new string[,] 
            {
                { "num1", "num2", "num3", "num4", "num5" },
            {"1", "2", "3", "4", "5" },
            {"1,1", "2,2", "3,3", "4,4", "5,5"  }
            };

            List<KeyValuePair<string, string[,]>> list = new List<KeyValuePair<string, string[,]>>();
            list.Add(new KeyValuePair<string, string[,]>("standardTable", standardTable));
            list.Add(new KeyValuePair<string, string[,]>("numTable", numTable));

            fileCon = new CListTableContainer("fileCon", list);
        }

        private void initAddFileCon()
        {
            string[,] standardTable = new string[,]
            {
                {"", "", "", "", "" },
            {"val1.3", "val2.3", "val3.3", "val4.3", "val5.3" },
            {"val1.4", "val2.4", "val3.4", "val4.4", "val5.4" }
            };

            string[,] numTable = new string[,]
            {
                {"", "", "", "", "" },
            {"100", "200", "300", "400", "500" },
            {"111", "222", "333", "444", "555"  }
            };

            string[,] personTable = new string[,]
            {
                {"Name", "Age"},
                {"Person1", "60" },
                {"Person2", "19" }
            };

            List<KeyValuePair<string, string[,]>> list = new List<KeyValuePair<string, string[,]>>();
            list.Add(new KeyValuePair<string, string[,]>("standardTable", standardTable));
            list.Add(new KeyValuePair<string, string[,]>("numTable", numTable));
            list.Add(new KeyValuePair<string, string[,]>("EmptyTable", new string[0, 0]));
            list.Add(new KeyValuePair<string, string[,]>("personTable", personTable));

            addFileCon = new CListTableContainer("addFileCon", list);
        }

        private void initAssertFileCon()
        {
            string[,] standardTable = new string[,]
            {
                { "col1", "col2", "col3", "col4", "col5" },
            {"val1.1", "val2.1", "val3.1", "val4.1", "val5.1" },
            {"val1.2", "val2.2", "val3.2", "val4.2", "val5.2"  },
            {"val1.3", "val2.3", "val3.3", "val4.3", "val5.3" },
            {"val1.4", "val2.4", "val3.4", "val4.4", "val5.4" }
            };

            string[,] numTable = new string[,]
            {
                { "num1", "num2", "num3", "num4", "num5" },
            {"1", "2", "3", "4", "5" },
            {"1,1", "2,2", "3,3", "4,4", "5,5"  },
            {"100", "200", "300", "400", "500" },
            {"111", "222", "333", "444", "555"  }
            };

            string[,] personTable = new string[,]
            {
                {"Name", "Age"},
                {"Person1", "60" },
                {"Person2", "19" }
            };

            List<KeyValuePair<string, string[,]>> list = new List<KeyValuePair<string, string[,]>>();
            list.Add(new KeyValuePair<string, string[,]>("standardTable", standardTable));
            list.Add(new KeyValuePair<string, string[,]>("numTable", numTable));
            list.Add(new KeyValuePair<string, string[,]>("EmptyTable", new string[0, 0]));
            list.Add(new KeyValuePair<string, string[,]>("personTable", personTable));

            assertFileCon = new CListTableContainer("assertFileCon", list);
        }

        [TestCleanup]
        public void endTest()
        {
            spreadSheetMock.mock.Object.closeSpreadsheet();
            Thread.Sleep(1000);
            Assert.AreEqual(mainThread.IsAlive, false);
        }

        //[TestMethod]
        public void testEditExcelSheet()
        {
            editFileWithAddFile();
            Assert.AreEqual(spreadSheetMock.mock.Object.tableContainer, assertFileCon);
        }

        private void editFileWithAddFile()
        {
            foreach (CTable table in addFileCon)
            {
                int numOfSheet = spreadSheetMock.getNumberOfSheet(table.Name);
                if (numOfSheet != -1)
                    addRowsToFile(table, numOfSheet);
                else
                {
                    numOfSheet = spreadSheetMock.mock.Object.addSheetBehind(table.Name);
                    spreadSheetMock.mock.Object.setContentOfSheet(numOfSheet, table.toArray(), true);
                }
            }
        }

        private void addRowsToFile(CTable table, int numOfSheet)
        {
            foreach (CRow row in table)
                spreadSheetMock.addRowToSheet(numOfSheet, row.convertToArray());
        }
    }
}
