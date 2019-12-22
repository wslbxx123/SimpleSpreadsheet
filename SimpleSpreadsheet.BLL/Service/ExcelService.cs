using SimpleSpreadsheet.Common.CustomException;
using SimpleSpreadsheet.Common.Model;
using System;
using System.Configuration;

namespace SimpleSpreadsheet.BLL.Service
{
    public class ExcelService: IExcelService
    {
        public int CellLength { get; set; }

        private Cell[,] CellList { get; set; }

        public char[,] CharList { get; set; }

        public int MaxCharRowCount { get; set; }
        public int MaxCharColumnCount { get; set; }

        public int MaxRowCount { get; set; }
        public int MaxColumnCount { get; set; }

        public ExcelService()
        {   
        }

        public void CreateExcel(int maxRowCount, int maxColumnCount, int cellLength)
        {
            this.CellList = new Cell[maxRowCount, maxColumnCount];
            this.CellLength = cellLength;
            this.MaxRowCount = maxRowCount;
            this.MaxColumnCount = maxColumnCount;
            this.MaxCharRowCount = maxRowCount * CellLength + 2;
            this.MaxCharColumnCount = maxColumnCount + 2;
            this.CharList = new char[this.MaxCharRowCount, this.MaxCharColumnCount];
            InitBorder();
        }

        public bool ExcelExists()
        {
            return this.CellList != null;
        }

        private void InitBorder()
        {
            for (int i = 0; i < MaxCharRowCount; i++)
            {
                for (int j = 0; j < MaxCharColumnCount; j++)
                {
                    this.CharList[i, j] = default(char);

                    if (j == 0 || j == MaxCharColumnCount - 1)
                    {
                        this.CharList[i, j] = '-';
                        continue;
                    }

                    if (i == 0 || i == MaxCharRowCount - 1)
                    {
                        this.CharList[i, j] = '|';
                    }
                }
            }
        }

        //
        // Summary:
        //     Update value of cell with the row and column index.
        //
        // Parameters:
        //   row:
        //     A 32-bit signed integer containing the row index.
        //   column:
        //     A 32-bit signed integer containing the column index.
        //   value:
        //     A 32-bit signed integer containing the new cell value.
        //
        // Returns:
        //     A 32-bit signed integer equivalent to the number contained in s.
        //
        // Exceptions:
        //   T:System.ArrayOutOfBoundException:
        //     row is bigger than row capacity, column is bigger than column capacity.
        public void UpdateCell(int row, int column, int value)
        {
            UpdateCell(row, column, value.ToString());
        }

        public void UpdateCell(int row, int column, string value)
        {
            if (row > this.MaxRowCount || column > this.MaxColumnCount)
            {
                throw new ArrayOutOfBoundException(this.MaxCharRowCount, this.MaxCharColumnCount);
            }

            if (value.Length > CellLength)
            {
                throw new LengthOutOfBoundException(CellLength);
            }

            var cell = this.CellList[row - 1, column - 1];

            if (cell == null)
            {
                cell = new Cell
                {
                    Row = row,
                    Column = column
                };
                this.CellList[row - 1, column - 1] = cell;
            }
            cell.Value = value.ToString();

            for (int i = 0; i < CellLength; i++)
            {
                if (value.Length > i)
                {
                    var chars = value.ToCharArray();
                    this.CharList[row * CellLength - i, column] = chars[value.Length - i - 1];
                }
            }
        }

        public void RemoveCell(int row, int column)
        {
            this.CellList[row - 1, column - 1] = null;
        }

        public Cell GetCell(int row, int column)
        {
            if (row > MaxRowCount || column > MaxColumnCount)
            {
                throw new ArrayOutOfBoundException(MaxRowCount, MaxColumnCount);
            }
            return this.CellList[row - 1, column - 1];
        }

        public char GetCellChar(int charRow, int charColumn)
        {
            return CharList[charRow, charColumn];
        }

        public bool HasValue(int row, int column)
        {
            var cell = GetCell(row, column);

            return cell != null && !string.IsNullOrEmpty(cell.Value);
        }


        public int SumCells(int beginRow, int beginColumn, int endRow, int endColumn)
        {
            var sum = 0;
            if (beginRow > MaxRowCount || endRow > MaxRowCount ||
                beginColumn > MaxColumnCount || endColumn > MaxColumnCount)
            {
                throw new IndexOutOfRangeException();
            }

            var cells = new Cell[endColumn - beginColumn + 1];
            for (int j = beginColumn; j <= endColumn; j++)
            {
                cells[j - beginColumn] = GetCell(endRow, j);
            }

            foreach (Cell cell in cells)
            {
                try
                {
                    var cellValue = Convert.ToInt32(cell.Value);
                    sum += cellValue;
                }
                catch
                {
                    throw new InvalidValueException(cell);
                }

            }

            return sum;
        }

        public void UpdateCellsByLine(int beginRow, int beginColumn, int endRow,
            int endColumn, string value = "x")
        {
            for (int i = beginRow; i <= endRow; i++)
            {
                for (int j = beginColumn; j <= endColumn; j++)
                {
                    UpdateCell(i, j, value);
                }
            }
        }

        public void UpdateCellsBySquare(int beginRow, int beginColumn, int endRow,
            int endColumn, string value = "x")
        {
            UpdateCellsByLine(beginRow, beginColumn, beginRow, endColumn, value);
            UpdateCellsByLine(beginRow, beginColumn, endRow, beginColumn, value);
            UpdateCellsByLine(endRow, beginColumn, endRow, endColumn, value);
            UpdateCellsByLine(beginRow, endColumn, endRow, endColumn, value);
        }

        public void UpdateCellsByBucket(int row, int column, string value = "x")
        {
            if (row > MaxRowCount || column > MaxColumnCount
                || row <= 0 || column <= 0)
            {
                return;
            }

            if (HasValue(row, column))
            {
                return;
            }

            UpdateCell(row, column, value);
            UpdateCellsByBucket(row + 1, column, value);
            UpdateCellsByBucket(row - 1, column, value);
            UpdateCellsByBucket(row, column + 1, value);
            UpdateCellsByBucket(row, column - 1, value);
        }
    }
}
