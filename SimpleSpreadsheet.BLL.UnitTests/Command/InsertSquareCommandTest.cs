
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.Common.CustomException;
using SimpleSpreadsheet.BLL.Service;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class InsertSquareCommandTest
    {
        [TestMethod]
        public void Test_InsertCellBySquare_14_1_18_3_Success()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertSquareCommand(excelService);
            command.Execute(new string[] { "R", "14", "1", "18", "3" });

            A.CallTo(() => excelService.UpdateCellsBySquare(14, 1, 18, 3, "x")).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(ExcelNotCreatedException))]
        public void Test_InsertCellBySquare_ExcelNotExists_ExcelNotCreatedException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(false);

            var command = new InsertSquareCommand(excelService);
            command.Execute(new string[] { "R", "14", "1", "18", "3" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_InsertCellBySquare_FourParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertSquareCommand(excelService);
            command.Execute(new string[] { "R", "14", "1", "18" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_InsertCellBySquare_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertSquareCommand(excelService);
            command.Execute(new string[] { "R", "a", "1", "18", "3" });
        }
    }
}