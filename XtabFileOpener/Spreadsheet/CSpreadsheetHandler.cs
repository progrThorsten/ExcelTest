using System;
using System.Text;
using System.Threading;
using ExcelDemo.TableContainer;
using ExcelDemo.Spreadsheet.SpreadsheetAdapter;

namespace ExcelDemo.Spreadsheet
{
    /// <summary>
    /// enables open the tables with a program that can show and edit tables
    /// </summary>
    internal class SpreadsheetHandler
    {
        internal ISpreadsheetAdapter spreadsheet
        {
            get; private set;
        }

        internal SpreadsheetHandler(CTableContainer tableContainer, ISpreadsheetAdapter _spreadsheet, bool autosize)
        {
            spreadsheet = _spreadsheet;
            spreadsheet.createSpreadsheet(tableContainer.Name);
            spreadsheet.saveSpreadsheet();
            createWorkbookFromTableContainer(tableContainer, autosize);
        }
        
        /// <summary>
        /// opens the spreadsheet
        /// </summary>
        internal void open()
        {
            //spreadsheet.save += onSave;
            spreadsheet.startListeningToSaving();
            spreadsheet.closed += onClosed;
            spreadsheet.startListeningToClosing();
            spreadsheet.show();
        }

        /// <summary>
        /// creates the workbook out of the TableContainer
        /// </summary>
        /// <param name="tableContainer">the TableContainer that should the workbook be created with</param>
        /// <param name="autoSize">defines whether the column width should be autosized</param>
        private void createWorkbookFromTableContainer(CTableContainer tableContainer, bool autoSize)
        {
            //necessary if Excel opens several sheets at the beginning
            closeExistingSheetsExceptOne();

            bool standardSheetExists = spreadsheet.SheetCount >= 1;
            //necessary to rename the standard sheet with an unusual name, so that no table of the xtab-file has the same name
            if (standardSheetExists)
                spreadsheet.renameSheet(0, "DefaultSheetWithInimitableName");

            foreach (CTable table in tableContainer)
                addTableToWorkbook(table, autoSize);

            //should never occur: if there is no table, then add default table
            if (tableContainer.Count == 0)
                spreadsheet.addSheetBehind("DefaultSheet");

            if (standardSheetExists)
                spreadsheet.deleteSheet(0);

            spreadsheet.activateSheet(0);

            spreadsheet.createTableContainer();
        }

        private void closeExistingSheetsExceptOne()
        {
            while (spreadsheet.SheetCount >= 2)
                spreadsheet.deleteSheet(0);
        }

        private void addTableToWorkbook(CTable table, bool autosize)
        {
            int number = spreadsheet.addSheetBehind(table.Name);
            if (!table.Empty)
                spreadsheet.setContentOfSheet(number, table.toArray(), autosize);
        }

        private bool closed = false;
        private void onClosed()
        {
            Monitor.Enter(this);
            closed = true;
            Monitor.Pulse(this);
            Monitor.Exit(this);
            spreadsheet.destroySpreadsheet();
        }

        /// <summary>
        /// lets the current thread wait until the Excel workbook is closed
        /// </summary>
        public void waitForClosing()
        {
            while (!closed)
                wait();
        }

        private void wait()
        {
            Monitor.Enter(this);
            Monitor.Wait(this);
            Monitor.Exit(this);
        }
    }
}
