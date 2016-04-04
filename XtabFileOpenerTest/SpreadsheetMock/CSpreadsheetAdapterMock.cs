using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDemo.TableContainer.ListTableContainer;
using ExcelDemo.TableContainer;
using ExcelDemo.Spreadsheet.SpreadsheetAdapter;
using Moq;

namespace ExcelDemoTest.SpreadsheetMock
{
    internal class CSpreadsheetAdapterMock
    {
        private string name;
        private List<KeyValuePair<string, string[,]>> tables;
        private CTableContainer tableContainer;

        internal bool Visible { get; private set; }

        private int activeSheet;

        private event applyChanges save;
        private event spreadsheetClosed closed;

        private const string nameAlreadyExistsMessage = "The name already exists";

        internal Mock<ISpreadsheetAdapter> mock
        {
            get;
            private set;
        }

        internal CSpreadsheetAdapterMock()
        {
            mock = new Mock<ISpreadsheetAdapter>(MockBehavior.Strict);
            tables = new List<KeyValuePair<string, string[,]>>();

            initProperties();
            initMethods();
        }

        private void initProperties()
        {
            mock.SetupGet(spreadsheet => spreadsheet.SheetCount).Returns(tables.Count);
            mock.SetupGet(spreadSheet => spreadSheet.tableContainer).Returns(new CListTableContainer(name, tables));
        }

        private void initMethods()
        {
            mock.Setup(spreadsheet => spreadsheet.createSpreadsheet(It.IsAny<string>())).Callback<string>(_name => name = _name);
            mock.Setup(spreadsheet => spreadsheet.closeSpreadsheet()).Callback(() => mock.Raise(spreadsheet => spreadsheet.closed += null));
            mock.Setup(spreadsheet => spreadsheet.destroySpreadsheet());
            mock.Setup(spreadsheet => spreadsheet.saveSpreadsheet()).Callback(() => mock.Raise(spreadsheet => spreadsheet.save+=null, tableContainer));
            mock.Setup(spreadsheet => spreadsheet.startListeningToSaving());
            mock.Setup(spreadsheet => spreadsheet.startListeningToClosing());
            mock.Setup(spreadsheet => spreadsheet.show()).Callback(() => Visible = true);
            mock.Setup(spreadsheet => spreadsheet.addSheetBehind(It.IsAny<string>())).Returns<string>(addSheetBehind);
            mock.Setup(spreadsheet => spreadsheet.setContentOfSheet(It.IsAny<int>(), It.IsAny<string[,]>(), It.IsAny<bool>())).
                Callback<int, string[,], bool>(setContentOfSheet);
            mock.Setup(spreadsheet => spreadsheet.deleteSheet(It.IsAny<int>())).Callback<int>(number => tables.RemoveAt(number));
            mock.Setup(spreadsheet => spreadsheet.activateSheet(It.IsAny<int>())).Callback<int>(number => activeSheet = number);
            mock.Setup(spreadsheet => spreadsheet.renameSheet(It.IsAny<int>(), It.IsAny<string>())).Callback<int, string>(renameSheet);
            mock.Setup(spreadsheet => spreadsheet.createTableContainer()).Callback(() => tableContainer = new CListTableContainer(name, tables));
        }

        private void close()
        {
            closed();
            Visible = false;
        }

        private int addSheetBehind(string name)
        {
            tables.Add(new KeyValuePair<string, string[,]>(name, null));
            return tables.Count() - 1;
        }

        internal void addSheetAt(string name, int number)
        {
            tables.Insert(number, new KeyValuePair<string, string[,]>(name, null));
        }

        private void setContentOfSheet(int number, string[,] content, bool autosize)
        {
            string key = tables.ElementAt(number).Key;
            tables[number] = new KeyValuePair<string, string[,]>(key, content);
        }

        internal void setCellOfSheet(int sheetNumber, int x, int y, string value)
        {
            string[,] table = tables.ElementAt(sheetNumber).Value;
            if (y >= table.GetLength(0) || x >= table.GetLength(1))
                table = Array2D.changeArraySize(table, y + 1, x + 1);
            table[y, x] = value;
            setContentOfSheet(sheetNumber, table, true);
        }

        internal void addRowToSheet(int sheetNumber, string[] row)
        {
            string[,] table = tables.ElementAt(sheetNumber).Value;
            table = Array2D.addRowToArray(table, row);
            setContentOfSheet(sheetNumber, table, true);
        }

        private void renameSheet(int number, string name)
        {
            if (doesNameExist(name)) throw new OperationCanceledException(nameAlreadyExistsMessage);
            string[,] content = tables.ElementAt(number).Value;
            tables[number] = new KeyValuePair<string, string[,]>(name, content);
        }

        public bool doesNameExist(string name)
        {
            return getNumberOfSheet(name) != -1;
        }

        public int getNumberOfSheet(string name)
        {
            int num = 0;
            foreach (KeyValuePair<string, string[,]> table in tables)
            {
                if (table.Key.Equals(name))
                    return num;
                num++;
            }
            return -1;
        }
    }
}
