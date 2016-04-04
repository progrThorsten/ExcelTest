using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDemo.TableContainer;
using ExcelDemo.TableContainer.ExcelTableContainer;

namespace ExcelDemo.Spreadsheet.SpreadsheetAdapter
{
    internal class CExcelAdapter : ISpreadsheetAdapter
    {
        private event applyChanges saveEvent;
        event applyChanges ISpreadsheetAdapter.save
        {
            add { saveEvent += value; }
            remove { saveEvent -= value; }
        }

        private event spreadsheetClosed closedEvent;
        event spreadsheetClosed ISpreadsheetAdapter.closed
        {
            add { closedEvent += value; }
            remove { closedEvent -= value; }
        }

        private FileInfo tmpFile;
        private Application excel;
        private Workbook workbook;

        internal CExcelAdapter() { }

        public void createSpreadsheet(string name)
        {
            excel = new Application();
            workbook = excel.Workbooks.Add(Type.Missing);
            excel.Caption = name;
        }

        public void closeSpreadsheet()
        {
            workbook.Close();
        }

        public void destroySpreadsheet()
        {
            tmpFile.Delete();
            //necessary to stop process
            deleteObject(workbook);
            deleteObject(excel);
        }

        private void deleteObject(Object o)
        {
            try { while (Marshal.ReleaseComObject(o) > 0) ; }
            catch { }
            finally { o = null; }
        }

        public void saveSpreadsheet()
        {
            if (tmpFile == null)
                saveWorkbookAs();
            else
                workbook.Save();
        }

        private void saveWorkbookAs()
        {
            tmpFile = new FileInfo(Path.GetTempPath() + Path.GetRandomFileName());
            workbook.SaveAs(tmpFile);
        }

        public void startListeningToSaving()
        {
            excel.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(onSave);   
        }

        private void onSave(Workbook savedWorkbook, bool SaveAsUI, ref bool Cancel)
        {
            
        }

        public void startListeningToClosing()
        {
            excel.WorkbookDeactivate += (onClose);
        }

        private void onClose(Workbook wb)
        {
            new Thread(waitForRealClosing).Start();
        }

        /// <summary>
        /// waits until the Excel workbook is really closed 
        /// (when the WorkbookDeactivate-event is fired, the application is not yet closed completely)
        /// </summary>
        private void waitForRealClosing()
        {
            //wait until the file is really closed (if this is missing, there can occur exceptions)
            while (excel.Workbooks.Count != 0)
                Thread.Sleep(10);
            raiseClosedEvent();
        }

        public void show()
        {
            excel.Visible = true;
        }

        public int SheetCount
        {
            get { return workbook.Sheets.Count; }
        }

        public CTableContainer tableContainer
        {
            get { return container; }
        }

        public int addSheetBehind(string name)
        {
            Worksheet sheet = workbook.Sheets.Add(After: workbook.Sheets.get_Item(workbook.Sheets.Count));
            sheet.Name = name;
            return workbook.Sheets.Count - 1;
        }

        public void setContentOfSheet(int number, string[,] content, bool autosize)
        {
            Worksheet sheet = workbook.Sheets[number + 1];
            Range startCell = sheet.Cells.get_Item(1, 1);
            Range endCell = sheet.Cells.get_Item(content.GetLength(0), content.GetLength(1));
            Range rangeToSet = sheet.get_Range(startCell, endCell);
            rangeToSet.Value = content;

            if (autosize)
                sheet.UsedRange.Columns.AutoFit();
        }

        public void deleteSheet(int number)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Delete();
        }

        public void activateSheet(int number)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Select(Type.Missing);
        }

        public void renameSheet(int number, string name)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Name = name;
        }

        private void raiseClosedEvent()
        {
            if (closedEvent != null) closedEvent();
        }

        private CTableContainer container;
        public void createTableContainer()
        {
            container = new CExcelTableContainer(tmpFile.Name, workbook.Sheets);
        }
    }
}
