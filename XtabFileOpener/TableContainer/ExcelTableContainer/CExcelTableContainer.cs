using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelDemo.TableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of CTableContainer, that contains Excel database tables
    /// </summary>
    internal class CExcelTableContainer : CTableContainer
    {
        internal CExcelTableContainer(string _name, Sheets _sheets) : base(_name)
        {
            sheets = _sheets;
        }

        private Sheets sheets;
        public override IEnumerator<CTable> GetEnumerator()
        {
            //tale only the used range in the Excel sheet
            foreach (Worksheet sheet in sheets)
                yield return new CExcelTable(sheet.Name, sheet.UsedRange);
        }

        public override int Count
        {
            get { return sheets.Count; }
        }
    }
}
