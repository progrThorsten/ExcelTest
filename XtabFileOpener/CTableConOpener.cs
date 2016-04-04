using System;
using System.Text;
using ExcelDemo.Spreadsheet.SpreadsheetAdapter;
using ExcelDemo.Spreadsheet;
using ExcelDemo.TableContainer;

namespace ExcelDemo
{
    public class CTableConOpener
    {
        private CTableContainer tableContainer;
        private SpreadsheetHandler spreadsheetHandler;

        public CTableConOpener(CTableContainer _tableContainer)
        {
            tableContainer = _tableContainer;
        }

        public void openFile(ISpreadsheetAdapter _spreadsheet)
        {
            spreadsheetHandler = new SpreadsheetHandler(tableContainer, _spreadsheet, true);
            spreadsheetHandler.open();
        }

        public void waitForClosing()
        {
            spreadsheetHandler.waitForClosing();
        }
    }
}
