
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class SumCellCommandTest
    {
        [TestMethod]
        public void Test_SumCells_1_2_1_3_1_4_Success()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);
            A.CallTo(() => excelService.SumCells(1, 2, 1, 3)).Returns(4);

            var command = new SumCellCommand(excelService);
            command.Execute(new string[] { "S", "1", "2", "1", "3", "1", "4" });

            A.CallTo(() => excelService.SumCells(1, 2, 1, 3)).MustHaveHappenedOnceExactly();
            A.CallTo(() => excelService.UpdateCell(1, 4, "4")).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(ExcelNotCreatedException))]
        public void Test_SumCells_ExcelNotExists_ExcelNotCreatedException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(false);

            var command = new SumCellCommand(excelService);
            command.Execute(new string[] { "S", "1", "2", "1", "3", "1", "4" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_SumCells_FourParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new SumCellCommand(excelService);
            command.Execute(new string[] { "S", "1", "2", "6" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_SumCells_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new SumCellCommand(excelService);
            command.Execute(new string[] { "S", "a", "2", "1", "3", "1", "4" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_SumCells_ColumnOrRowNotSame_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new SumCellCommand(excelService);
            command.Execute(new string[] { "S", "1", "2", "3", "4", "2", "2" });
        }
    }
}