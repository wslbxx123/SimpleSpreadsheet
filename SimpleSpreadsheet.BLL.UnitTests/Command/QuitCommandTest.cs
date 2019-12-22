
using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimpleSpreadsheet.BLL.Service;
using SimpleSpreadsheet.Common.CustomException;
using System;

namespace SimpleSpreadsheet.BLL.UnitTests.Command
{
    [TestClass]
    public class QuitCommandTest
    {
        [TestMethod]
        [ExpectedException(typeof(QuitException))]
        public void Test_Quit_Success()
        {
            var excelService = A.Fake<IExcelService>();

            var command = new QuitCommand(excelService);
            command.Execute(new string[] { "Q"});
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidValueException))]
        public void Test_Quit_TwoParameters_InvalidValueException()
        {
            var excelService = A.Fake<IExcelService>();

            var command = new QuitCommand(excelService);
            command.Execute(new string[] { "Q", "1" });
        }
    }
}