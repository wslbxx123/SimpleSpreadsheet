using System;
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class CreateCommandTest
    {
        [TestMethod]
        public void Test_CreateExcel_20_4_Success()
        {
            var excelService = A.Fake<IExcelService>();

            var command = new CreateCommand(excelService);
            command.Execute(new string[] { "C", "20", "4" });

            A.CallTo(() => excelService.CreateExcel(20, 4, 1)).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_CreateExcel_FiveParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            var command = new CreateCommand(excelService);
            command.Execute(new string[] { "C", "20", "4", "5", "1" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_CreateExcel_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            var command = new CreateCommand(excelService);
            command.Execute(new string[] { "C", "A", "4" });
        }
    }
}
