
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL;
using SimpleSpreadsheet.Common.CustomException;
using SimpleSpreadsheet.BLL.Service;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class InsertCellCommandTest
    {
        [TestMethod]
        public void Test_UpdateCell_1_2_2_Success()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertCellCommand(excelService);
            command.Execute(new string[] { "N", "1", "2", "2" });

            A.CallTo(() => excelService.UpdateCell(1, 2, "2")).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(ExcelNotCreatedException))]
        public void Test_UpdateCell_ExcelNotExists_ExcelNotCreatedException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(false);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "N", "1", "2", "2" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_UpdateCell_FiveParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "N", "1", "2", "2", "p" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_UpdateCell_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "N", "a", "2", "2" });
        }
    }
}