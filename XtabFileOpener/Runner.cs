using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDemo.Spreadsheet;
using ExcelDemo.Spreadsheet.SpreadsheetAdapter;
using ExcelDemo.TableContainer.ListTableContainer;
using ExcelDemo.TableContainer;

namespace ExcelDemo
{
    public class Runner
    {
        public static void Main(String[] args)
        {
            CTableConOpener opener = new CTableConOpener(createContainer());
            opener.openFile(new CExcelAdapter());
            opener.waitForClosing();
        }

        private static CTableContainer createContainer()
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

            return new CListTableContainer("assertFileCon", list);
        }
    }
}
