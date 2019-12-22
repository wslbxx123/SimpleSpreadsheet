using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL
{
    public abstract class BaseCommand
    {
        protected IExcelService _excelService { get; set; }
        public int ParamMaxCount { get; set; }
        public int ParamMinCount { get; set; }

        public BaseCommand(IExcelService ExcelRepository)
        {
            this._excelService = ExcelRepository;
        }

        public abstract void Execute(string[] args);

        public virtual string GetDescription()
        {
            return "The arguments you input are invalid";
        }
        
        public void CheckParamsValid(string[] args)
        {
            if (args.Length < this.ParamMinCount || args.Length > this.ParamMaxCount)
            {
                throw new InvalidValueException();
            }
        }

        public void PrintExcel()
        {
            for (int j = 0; j < this._excelService.MaxCharColumnCount; j++)               
            {
                for (int i = 0; i < this._excelService.MaxCharRowCount; i++)
                {
                    Console.Write(this._excelService.GetCellChar(i, j));
                }

                Console.Write("\n");
            }
        }
    }
}
