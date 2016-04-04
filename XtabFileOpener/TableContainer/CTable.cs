using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDemo.TableContainer.ListTableContainer;

namespace ExcelDemo.TableContainer
{
    /// <summary>
    /// represents a database tables
    /// </summary>
    public abstract class CTable : IEnumerable<CRow>
    {
        public CTable(string _name, CColumns _columns)
        {
            name = _name;
            columns = _columns;
        }

        private string name;
        public string Name
        {
            get { return name; }
        }

        public abstract IEnumerator<CRow> GetEnumerator();

        /// <summary>
        /// enables iterating over the rows of this table excluding the row with the column names
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        private CColumns columns;
        public CColumns Columns
        {
            get { return columns; }
        }

        /// <summary>
        /// number of columns of the table
        /// </summary>
        public abstract int Width
        {
            get;
        }

        /// <summary>
        /// number of rows of the table including the row with the column names
        /// </summary>
        public abstract int Height
        {
            get;
        }

        /// <summary>
        /// Length of the longest row, including the columns
        /// </summary>
        public abstract int maxWidth
        {
            get;
        }

        /// <summary>
        /// creates a string array of this table, including the column names
        /// </summary>
        /// <returns></returns>
        public string[,] toArray()
        {
            string[,] result = new string[Height, maxWidth]; //+ (columns.Exists ? 1 : 0)

            int i = 0;
            int j = 0;
            foreach (string value in Columns)
            {
                result[i, j] = value;
                j++;
            }

            i += Columns.Exists ? 1 : 0;
            foreach (CRow row in this)
            {
                j = 0;
                foreach (string value in row)
                {
                    result[i, j] = value;
                    j++;
                }
                i++;
            }
            Array2D.replaceNullByEmptyString(result);
            return result;
        }

        /// <summary>
        /// checks whether the table is empty
        /// </summary>
        public bool Empty
        {
            get { return Height == 0 || Width == 0; }
        }

        public override bool Equals(object obj)
        {
            CTable tab = obj as CTable;
            string[,] arr1 = toArray();
            string[,] arr2 = tab.toArray();
            return ListTableContainer.Array2D.areEqual(arr1, arr2);
        }
    }
}
