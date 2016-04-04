using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelDemo.TableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of CRow, that manages a row of an Excel table
    /// </summary>
    internal class CExcelRow : CRow
    {
        internal CExcelRow(Range _cells, int _rowNumber)
        {
            cells = _cells;
            rowNumber = _rowNumber;
        }
        
        private int rowNumber;
        private Range cells;
        public override IEnumerator<string> GetEnumerator()
        {
            for (int i = 1; i <= cells.Columns.Count; i++)
            {
                Object value = (cells[rowNumber, i] as Range).Value2;
                if (value != null)
                    yield return value.ToString();
                else
                    yield return ""; //return an empty string in case of an empty cell
            }
        }
    }
}
