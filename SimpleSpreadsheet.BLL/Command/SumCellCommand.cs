using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL
{
    public class SumCellCommand : BaseCommand
    {
        public SumCellCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMaxCount = this.ParamMinCount = 7;
        }

        public override void Execute(string[] args)
        {
            if (!_excelService.ExcelExists())
            {
                throw new ExcelNotCreatedException();
            }

            this.CheckParamsValid(args);

            var row1 = Convert.ToInt32(args[1]);
            var column1 = Convert.ToInt32(args[2]);

            var row2 = Convert.ToInt32(args[3]);
            var column2 = Convert.ToInt32(args[4]);

            var row3 = Convert.ToInt32(args[5]);
            var column3 = Convert.ToInt32(args[6]);

            if (row1 != row2)
            {
                throw new InvalidValueException("You can only sum the value in the same column");
            }

            var sum = _excelService.SumCells(row1, column1, row2, column2);
            _excelService.UpdateCell(row3, column3, sum.ToString());
            base.PrintExcel();
        }
    }
}
