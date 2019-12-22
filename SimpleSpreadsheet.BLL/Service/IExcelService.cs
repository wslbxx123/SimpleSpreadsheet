using SimpleSpreadsheet.Common.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleSpreadsheet.BLL.Service
{
    public interface IExcelService
    {
        int MaxRowCount { get; }
        int MaxColumnCount { get; }
        int MaxCharColumnCount { get; }
        int MaxCharRowCount { get; }

        void CreateExcel(int row, int column, int cellLength);
        void UpdateCell(int row, int column, string value);
        void RemoveCell(int row, int column);
        char GetCellChar(int charRow, int charColumn);
        bool ExcelExists();
        void UpdateCellsByLine(int beginRow, int beginColumn, int endRow,
            int endColumn, string value = "x");
        void UpdateCellsBySquare(int beginRow, int beginColumn, int endRow,
            int endColumn, string value = "x");
        Cell GetCell(int row, int count);
        void UpdateCellsByBucket(int row, int column, string value);
        int SumCells(int beginRow, int beginColumn, int endRow, int endColumn);
    }
}
