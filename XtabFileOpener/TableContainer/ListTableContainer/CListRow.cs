using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ListTableContainer
{
    internal class CListRow : CRow
    {
        private string[] values;

        internal CListRow(string[] _values)
        {
            values = _values;
        }

        public override IEnumerator<string> GetEnumerator()
        {
            foreach (string value in values)
                yield return value;
        }
    }
}
