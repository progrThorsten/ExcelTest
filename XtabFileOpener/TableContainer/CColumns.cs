using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer
{
    /// <summary>
    /// represents the columns of a database table
    /// </summary>
    public abstract class CColumns : IEnumerable<string>
    {
        public abstract IEnumerator<string> GetEnumerator();

        /// <summary>
        /// enables iterating over the columns
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        /// <summary>
        /// checks whether these columns exist; if they are not existing, the corresponding table has no column definition
        /// </summary>
        public abstract bool Exists
        {
            get;
        }
    }
}
