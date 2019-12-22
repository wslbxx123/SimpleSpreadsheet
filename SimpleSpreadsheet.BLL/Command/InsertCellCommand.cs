using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL
{
    public class InsertCellCommand : BaseCommand
    {
        public InsertCellCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMinCount = this.ParamMaxCount = 4;
        }

        public override void Execute(string[] args)
        {
            if(!_excelService.ExcelExists())
            {
                throw new ExcelNotCreatedException();
            }

            this.CheckParamsValid(args);

            _excelService.UpdateCell(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), args[3]);
            base.PrintExcel();
        }
    }
}
