
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common;
using SimpleSpreadsheet.Common.CustomException;

namespace SimpleSpreadsheet.BLL
{
    public class QuitCommand : BaseCommand

    {
        public QuitCommand(IExcelService excelService) : base(excelService)
        {
            this.ParamMaxCount = this.ParamMinCount = 1;
        }

        public override void Execute(string[] args)
        {
            this.CheckParamsValid(args);

            throw new QuitException();
        }
    }
}