using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL
{
    public class InsertSquareCommand : BaseCommand
    {
        public InsertSquareCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMaxCount = 6;
            this.ParamMinCount = 5;
        }

        public override void Execute(string[] args)
        {
            if (!_excelService.ExcelExists())
            {
                throw new ExcelNotCreatedException();
            }

            this.CheckParamsValid(args);

            var value = "x";
            if (args.Length == this.ParamMaxCount)
            {
                value = args[5];
            }

            _excelService.UpdateCellsBySquare(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]),
                    Convert.ToInt32(args[3]), Convert.ToInt32(args[4]), value);
            PrintExcel();
        }
    }
}