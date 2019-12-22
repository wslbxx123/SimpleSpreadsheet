using Ninject;
using SimpleSpreadsheet;
using SimpleSpreadsheet.BLL;
using SimpleSpreadsheet.Common;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreedsheet
{
    class Program
    {
        private static IKernel kernel;
        static Program()
        {
            kernel = new Ninject.StandardKernel(new InjectModule());
        }

        static T GetType<T>()
        {
            return kernel.Get<T>();
        }

        static void Main(string[] args)
        {
            var factory = GetType<CommandFactory>();

            while (true)
            {
                Console.Write("enter command: ");
                string readStr = Console.ReadLine();

                var inputArgs = readStr.Split(' ');

                if(inputArgs.Length == 0)
                {
                    continue;
                }

                BaseCommand command = factory.GetCommand(inputArgs[0]);

                try
                {           
                    command.Execute(inputArgs);
                }
                catch(QuitException ex)
                {
                    break;
                }
                catch(ExcelNotCreatedException ex)
                {
                    Console.WriteLine(ex.Message);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(command.GetDescription());
                }
                
            }
        }
    }
}
