using SimpleSpreadsheet.Common.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleSpreadsheet.Common.CustomException
{
    public class InvalidValueException : Exception
    {
        public InvalidValueException(Cell cell)
           : base(string.Format("Cell [{0}, {1}] has invalid value that caused the operation failure.", cell.Row, cell.Column))
        {
        }

        public InvalidValueException(string message) : base(message)
        {

        }

        public InvalidValueException()
            : base("The input values are invalid.")
        {

        }
    }
}
