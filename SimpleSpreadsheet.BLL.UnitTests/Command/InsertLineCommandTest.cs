
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.Common.CustomException;
using SimpleSpreadsheet.BLL.Service;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class InsertLineCommandTest
    {
        [TestMethod]
        public void Test_InsertCellByLine_1_2_6_2_Success()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertLineCommand(excelService);
            command.Execute(new string[] { "L", "1", "2", "6", "2" });

            A.CallTo(() => excelService.UpdateCellsByLine(1, 2, 6, 2, "x")).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(ExcelNotCreatedException))]
        public void Test_InsertCellByLine_ExcelNotExists_ExcelNotCreatedException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(false);

            var command = new InsertLineCommand(excelService);
            command.Execute(new string[] { "L", "1", "2", "6", "2" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_UpdateCellsByLine_FourParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertLineCommand(excelService);
            command.Execute(new string[] { "L", "1", "2", "6" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_UpdateCellsByLine_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertLineCommand(excelService);
            command.Execute(new string[] { "L", "a", "2", "6", "2"});
        }
    }
}