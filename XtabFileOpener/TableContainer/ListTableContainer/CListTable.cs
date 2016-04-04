using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ListTableContainer
{
    internal class CListTable : CTable
    {
        private string[,] rows;

        internal CListTable(string _name, string[,] _rows, string[] columns):base(_name, new CListColumns(columns))
        {
            rows = _rows;
        }

        public override int Height
        {
            get { return rows.GetLength(0) + (Columns.Exists ? 1 : 0); }
        }

        public override int Width
        {
            get
            {
                return Columns.Count();
            }
        }

        public override int maxWidth
        {
            get
            {
                if (Width > rows.GetLength(1))
                    return Width;
                else
                    return rows.GetLength(1);
            }
        }

        public override IEnumerator<CRow> GetEnumerator()
        {
            for (int i = 0; i < rows.GetLength(0); i++)
                yield return new CListRow(Array2D.getRow(rows, i));
        }
    }
}
