using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ListTableContainer
{
    internal class CListColumns : CColumns
    {
        internal CListColumns(string[] _columns)
        {
            columns = _columns;
        }

        private readonly string[] columns;
        public override IEnumerator<string> GetEnumerator()
        {
            foreach (string column in columns)
                yield return column;
        }

        public override bool Exists
        {
            get { return columns.Length > 0; }
        }
    }
}
