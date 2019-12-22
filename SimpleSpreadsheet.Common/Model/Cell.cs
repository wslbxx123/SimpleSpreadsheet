using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleSpreadsheet.Common.Model
{
    public class Cell
    {
        public int Row { get; set; }

        public int Column { get; set; }

        public string Value { get; set; }

        //public CellType Type { get; set; }
    }

    public enum CellType
    {
        NUM,
        STRING
    }
}
