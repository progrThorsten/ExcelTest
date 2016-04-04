using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer
{
    /// <summary>
    /// represents a row (a record) of a database table
    /// </summary>
    public abstract class CRow : IEnumerable<string>
    {
        public abstract IEnumerator<string> GetEnumerator();

        /// <summary>
        /// enables iterating over the values in the row
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public string[] convertToArray()
        {
            List<string> list = new List<string>();

            foreach (string s in this)
                list.Add(s);

            return list.ToArray();
        }
    }
}
