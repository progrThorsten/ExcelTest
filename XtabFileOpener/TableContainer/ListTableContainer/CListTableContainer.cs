using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ListTableContainer
{
    /// <summary>
    /// TableContainer that manages its tables using lists
    /// </summary>
    public class CListTableContainer : CTableContainer
    {
        private IEnumerable<KeyValuePair<string, string[,]>> tables;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_name">name of the TableContainer</param>
        /// <param name="_tables">tables of the TableContainer</param>
        public CListTableContainer(string _name, IEnumerable<KeyValuePair<string, string[,]>> _tables) : base(_name)
        {
            tables = _tables;
        }

        public override int Count
        {
            get { return tables.Count(); }
        }

        public override IEnumerator<CTable> GetEnumerator()
        {
            foreach (KeyValuePair<string, string[,]> table in tables)
                yield return new CListTable(table.Key, Array2D.removeRow(table.Value, 0), Array2D.getRow(table.Value, 0));
        }
    }
}
