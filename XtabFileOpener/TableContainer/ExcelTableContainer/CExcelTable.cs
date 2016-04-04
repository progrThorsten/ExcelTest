using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelDemo.TableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of CTable, that manages an Excel table
    /// </summary>
    internal class CExcelTable : CTable
    {
        internal CExcelTable(string _name, Range _cells) : base(_name, new CExcelColumns(_cells))
        {
            cells = _cells;
        }

        private Range cells;
        public override IEnumerator<CRow> GetEnumerator()
        {
            //begin with the second row in the Excel sheet, because the first row contains the column names
            for (int i = 2; i <= cells.Rows.Count; i++)
                yield return new CExcelRow(cells, i);
        }

        public override int Width
        {
            get { return cells.Columns.Count; }
        }

        public override int maxWidth
        {
            get { return Width; }
        }

        public override int Height
        {
            get { return cells.Rows.Count; }
        }
    }
}
