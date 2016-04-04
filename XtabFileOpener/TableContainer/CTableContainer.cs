using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer
{
    /// <summary>
    /// Represents a container of database tables
    /// </summary>
    public abstract class CTableContainer : IEnumerable<CTable>
    {
        public CTableContainer(string _name)
        {
            name = _name;
        }
        
        private string name;
        public string Name
        {
            get { return name; }
        }

        public abstract IEnumerator<CTable> GetEnumerator();

        /// <summary>
        /// enables iterating over the tables of this container
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        /// <summary>
        /// number of the tables
        /// </summary>
        public abstract int Count
        {
            get;
        }

        public override bool Equals(object obj)
        {
            CTableContainer con = obj as CTableContainer;
            var bothCons = this.Zip(con, (t1, t2) => new { Table1 = t1, Table2 = t2 });
            foreach (var com in bothCons)
            {
                if (!com.Table1.Equals(com.Table2)) return false;
            }
            return true;
        }
    }
}
