using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of CColumns, that manages the columns of an Excel table
    /// </summary>
    internal class CExcelColumns : CColumns
    {
        internal CExcelColumns(Range _cells)
        {
            cells = _cells;
        }

        private Range cells;
        public override IEnumerator<string> GetEnumerator()
        {
            for (int i = 1; i <= cells.Columns.Count; i++)
            {
                Object value = (cells[1, i] as Range).Value2;
                if (value != null)
                    yield return value.ToString();
                else
                    yield return "";
            }
        }

        public override bool Exists
        {
            //every Excel-table with at least one row has columns (the first row)
            get { return cells.Rows.Count >= 1; }
        }
    }
}
