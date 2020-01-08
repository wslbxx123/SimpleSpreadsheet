using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;

namespace SimpleSpreadsheet.UnitTest
{
    [TestClass]
    public class ExcelServiceTest
    {
        [TestMethod]
        [ExpectedException(typeof(ArrayOutOfBoundException))]
        public void Test_UpdateCell_OutOfBound()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);

            excelService.UpdateCell(11, 1, 5);
        }

        [TestMethod]
        [ExpectedException(typeof(LengthOutOfBoundException))]
        public void Test_UpdateCell_ValueLengthFour_Throw()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);

            excelService.UpdateCell(1, 1, 2567);
        }

        [TestMethod]
        public void Test_UpdateCell_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);
            excelService.UpdateCell(1, 1, "5");

            var actual = excelService.GetCell(1, 1).Value;
            var expected = "5";

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void Test_ExcelExists_True()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);

            var actual = excelService.ExcelExists();

            Assert.IsTrue(actual);
        }

        [TestMethod]
        public void Test_ExcelExists_False()
        {
            var excelService = new ExcelService();

            var actual = excelService.ExcelExists();

            Assert.IsFalse(actual);
        }

        public void Test_GetCell_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);
            excelService.UpdateCell(5, 5, 10);

            var actual = excelService.GetCell(5, 5);
            var expected = 10;

            Assert.Equals(actual, expected);
        }

        public void Test_RemoveCell_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);
            excelService.UpdateCell(5, 5, 10);

            excelService.RemoveCell(5, 5);

            var actual = excelService.GetCell(5, 5);

            Assert.IsNull(actual);
        }

        public void Test_HasValue_True()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);
            excelService.UpdateCell(5, 5, 10);

            var actual = excelService.HasValue(5, 5);

            Assert.IsTrue(actual);
        }

        public void Test_HasValue_False()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);

            var actual = excelService.HasValue(5, 5);

            Assert.IsFalse(actual);
        }

        [TestMethod]
        public void Test_SumCells_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(10, 10, 1);
            excelService.UpdateCell(1, 1, 5);
            excelService.UpdateCell(1, 2, 5);

            var actual = excelService.SumCells(1, 1, 1, 2);
            var expected = 10;

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void Test_UpdateCellsByLine_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(20, 4, 1);

            excelService.UpdateCellsByLine(1, 2, 6, 2);

            Assert.AreEqual(excelService.GetCell(1, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(2, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(3, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(4, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(5, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(6, 2).Value, "x");
        }

        [TestMethod]
        public void Test_UpdateCellsBySquare_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(20, 4, 1);

            excelService.UpdateCellsBySquare(14, 1, 18, 3);

            Assert.AreEqual(excelService.GetCell(14, 1).Value, "x");
            Assert.AreEqual(excelService.GetCell(15, 1).Value, "x");
            Assert.AreEqual(excelService.GetCell(16, 1).Value, "x");
            Assert.AreEqual(excelService.GetCell(17, 1).Value, "x");
            Assert.AreEqual(excelService.GetCell(18, 1).Value, "x");
            Assert.AreEqual(excelService.GetCell(14, 2).Value, "x");
            Assert.AreEqual(excelService.GetCell(14, 3).Value, "x");
            Assert.AreEqual(excelService.GetCell(15, 3).Value, "x");
            Assert.AreEqual(excelService.GetCell(16, 3).Value, "x");
            Assert.AreEqual(excelService.GetCell(17, 3).Value, "x");
            Assert.AreEqual(excelService.GetCell(18, 3).Value, "x");
            Assert.AreEqual(excelService.GetCell(18, 2).Value, "x");
        }

        [TestMethod]
        public void Test_UpdateCellsByBucket_Success()
        {
            var excelService = new ExcelService();
            excelService.CreateExcel(20, 4, 1);
            excelService.UpdateCellsBySquare(14, 1, 18, 3);
            excelService.UpdateCellsByBucket(15, 2, "o");

            Assert.AreEqual(excelService.GetCell(15, 2).Value, "o");
            Assert.AreEqual(excelService.GetCell(16, 2).Value, "o");
            Assert.AreEqual(excelService.GetCell(17, 2).Value, "o");
        }
    }
}
