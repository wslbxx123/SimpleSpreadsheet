
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL;
using SimpleSpreadsheet.Common.CustomException;
using SimpleSpreadsheet.BLL.Service;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class InsertBucketCommandTest
    {
        [TestMethod]
        public void Test_UpdateCellsByBucket_10_3_o_Success()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "B", "10", "3", "o" });

            A.CallTo(() => excelService.UpdateCellsByBucket(10, 3, "o")).MustHaveHappenedOnceExactly();
        }

        [TestMethod]
        [ExpectedException(typeof(ExcelNotCreatedException))]
        public void Test_UpdateCellsByBucket_ExcelNotExists_ExcelNotCreatedException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(false);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "B", "10", "3", "o" });
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_UpdateCellsByBucket_FiveParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "B", "10", "3", "o", "p" });
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void Test_UpdateCellsByBucket_ParameterInvalid_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            A.CallTo(() => excelService.ExcelExists()).Returns(true);

            var command = new InsertBucketCommand(excelService);
            command.Execute(new string[] { "B", "a", "3", "o" });
        }
    }
}