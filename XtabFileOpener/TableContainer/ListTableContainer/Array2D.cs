using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.TableContainer.ListTableContainer
{
    /// <summary>
    /// contains static methods that are useful for handling with two dimensional arrays
    /// </summary>
    public static class Array2D
    {
        internal static string[] getRow(string[,] array, int rowNum)
        {
            string[] row = new string[array.GetLength(1)];
            for (int i = 0; i < row.Length; i++)
                row[i] = array[rowNum, i];
            return row;
        }

        internal static string[,] removeRow(string[,] array, int rowNum)
        {
            if (array.GetLength(0) == 0)
                return array;

            string[,] newArray = new string[array.GetLength(0) - 1, array.GetLength(1)];
            int col = 0;
            for (int i = 0; i < newArray.GetLength(0); i++)
            {
                if (col == rowNum) col++; ;

                for (int j = 0; j < newArray.GetLength(1); j++)
                    newArray[i, j] = array[col, j];

                col++;
            }
            return newArray;
        }

        /// <summary>
        /// creates a new array from an exisiting array with new dimensions
        /// </summary>
        /// <param name="array">old array</param>
        /// <param name="height">new height</param>
        /// <param name="width">new width</param>
        /// <returns>new array</returns>
        public static string[,] changeArraySize(string[,] array, int height, int width)
        {
            int maxHeight = Math.Max(height, array.GetLength(0));
            int minHeight = Math.Min(height, array.GetLength(0));
            int maxWidth = Math.Max(width, array.GetLength(1));
            int minWidth = Math.Min(width, array.GetLength(1));

            string[,] newArray = new string[maxHeight, maxWidth];

            for (int i = 0; i < minHeight; i++)
            {
                for (int j = 0; j < minWidth; j++)
                    newArray[i, j] = array[i, j];
            }
            replaceNullByEmptyString(newArray);
            return newArray;
        }

        public static void replaceNullByEmptyString(string[,] array)
        {
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                    if (array[i, j] == null)
                        array[i, j] = String.Empty;
            }
        }

        /// <summary>
        /// creates a new array from an existing array, that has one more row
        /// </summary>
        /// <param name="array">old array</param>
        /// <param name="row">row that should be added</param>
        /// <returns>new array</returns>
        public static string[,] addRowToArray(string[,] array, string[] row)
        {
            int indexOfNewRow = array.GetLength(0);
            string[,] newArray = changeArraySize(array, array.GetLength(0) + 1, array.GetLength(1));//new string[indexOfNewRow + 1, array.GetLength(1)];
            if (row.Length > array.GetLength(1))
            {
                newArray = changeArraySize(newArray, newArray.GetLength(0), row.Length);
            }
            for (int i = 0; i < row.Length; i++)
                newArray[indexOfNewRow, i] = row[i];
            return newArray;
        }

        public static bool areEqual(string[,] arr1, string[,] arr2)
        {
            if (arr1.GetLength(0) == arr2.GetLength(0) && arr1.GetLength(1) == arr2.GetLength(1))
            {
                for (int i = 0; i < arr1.GetLength(0); i++)
                {
                    for (int j = 0; j < arr2.GetLength(1); j++)
                        if (!arr1[i, j].Equals(arr2[i, j])) return false;
                }
                return true;
            }
            else return false;
        }
    }
}
