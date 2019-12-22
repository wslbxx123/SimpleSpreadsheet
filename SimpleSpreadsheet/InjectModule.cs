using Ninject.Modules;
using SimpleSpreadsheet.BLL.Service;

namespace SimpleSpreadsheet
{
    public class InjectModule : NinjectModule
    {
        public override void Load()
        {
            this.Bind<IExcelService>().To<ExcelService>();
        }
    }
}
