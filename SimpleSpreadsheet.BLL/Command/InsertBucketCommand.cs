using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL
{
    public class InsertBucketCommand : BaseCommand
    {
        public InsertBucketCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMaxCount = this.ParamMinCount = 4;
        }

        public override void Execute(string[] args)
        {
            if (!_excelService.ExcelExists())
            {
                throw new ExcelNotCreatedException();
            }

            this.CheckParamsValid(args);

            _excelService.UpdateCellsByBucket(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]),
                    args[3]);
            base.PrintExcel();
        }
    }
}