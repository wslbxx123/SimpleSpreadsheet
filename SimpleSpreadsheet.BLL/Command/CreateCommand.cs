using SimpleSpreadsheet.BLL.Service;
using System;

namespace SimpleSpreadsheet.BLL
{
    public class CreateCommand : BaseCommand
    {
        public CreateCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMaxCount = 4;
            this.ParamMinCount = 3;
        }

        public override void Execute(string[] args)
        {
            this.CheckParamsValid(args);

            var cellLength = 1;
            if(args.Length == this.ParamMaxCount)
            {
                cellLength = Convert.ToInt32(args[3]);
            }
            _excelService.CreateExcel(Convert.ToInt32(args[1]), 
                Convert.ToInt32(args[2]), cellLength);
            PrintExcel();
        }

        public override string GetDescription()
        {
            return "The arguments should be C [Row Number] [Column Number]";
        }
    }
}
