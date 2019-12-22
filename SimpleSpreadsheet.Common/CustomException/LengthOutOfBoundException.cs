using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleSpreadsheet.Common.CustomException
{
    public class LengthOutOfBoundException : Exception
    {
        public LengthOutOfBoundException(int length)
            : base(string.Format("The input value is over the max length {0}", length))
        {
        }
    }
}
