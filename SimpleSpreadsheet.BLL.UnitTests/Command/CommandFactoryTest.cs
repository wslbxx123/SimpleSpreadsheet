using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class CommandFactoryTest
    {
        [TestMethod]
        public void Test_GetCommand_C_GetCreateCommand()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("C");
            Assert.IsInstanceOfType(result, typeof(CreateCommand));
        }

        [TestMethod]
        public void Test_GetCommand_N_GetInsertCellCommand()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("N");
            Assert.IsInstanceOfType(result, typeof(InsertCellCommand));
        }

        [TestMethod]
        public void Test_GetCommand_S_GetSumCellCommand()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("S");
            Assert.IsInstanceOfType(result, typeof(SumCellCommand));
        }

        [TestMethod]
        public void Test_GetCommand_Q_GetQuitCommand()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("Q");
            Assert.IsInstanceOfType(result, typeof(QuitCommand));
        }

        [TestMethod]
        public void Test_GetCommand_L_GetInsertLineCommand()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("L");
            Assert.IsInstanceOfType(result, typeof(InsertLineCommand));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_GetCommand_Z_ThrowInvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            var factory = new CommandFactory(excelService);
            var result = factory.GetCommand("Z");
        }
    }
}
