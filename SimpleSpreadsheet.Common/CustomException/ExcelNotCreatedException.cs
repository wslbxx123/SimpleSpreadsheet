using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleSpreadsheet.Common.CustomException
{
    public class ExcelNotCreatedException : Exception
    {
        public ExcelNotCreatedException()
            : base("The excel hasn't been created yet.")
        {
        }
    }
}
