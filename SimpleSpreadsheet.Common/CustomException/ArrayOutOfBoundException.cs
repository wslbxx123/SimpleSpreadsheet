using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleSpreadsheet.Common.CustomException
{
    public class ArrayOutOfBoundException : Exception
    {
        public ArrayOutOfBoundException(int maxRow, int maxColumn) 
            : base(string.Format("The row and column you input are out of bound. " +
                "The max row is {0}, The max column is {1}", maxRow, maxColumn))
        {
        }
    }
}
