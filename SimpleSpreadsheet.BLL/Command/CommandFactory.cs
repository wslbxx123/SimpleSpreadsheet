using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common;
using SimpleSpreadsheet.Common.CustomException;

namespace SimpleSpreadsheet.BLL
{
    public class CommandFactory
    {
        private IExcelService _excelService;

        public CommandFactory(IExcelService excelService)
        {
            this._excelService = excelService;
        }

        public BaseCommand GetCommand(string argType)
        {
            BaseCommand command;

            switch (argType.ToLowerInvariant())
            {
                case "c":
                    command = new CreateCommand(_excelService);
                    break;
                case "n":
                    command = new InsertCellCommand(_excelService);
                    break;
                case "s":
                    command = new SumCellCommand(_excelService);
                    break;
                case "q":
                    command = new QuitCommand(_excelService);
                    break;
                case "l":
                    command = new InsertLineCommand(_excelService);
                    break;
                case "r":
                    command = new InsertSquareCommand(_excelService);
                    break;
                case "b":
                    command = new InsertBucketCommand(_excelService);
                    break;
                default:
                    throw new InvalidValueException();
            }

            return command;
        }
    }
}
