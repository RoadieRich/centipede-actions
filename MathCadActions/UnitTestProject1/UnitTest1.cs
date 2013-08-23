using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var result = MessageBox.Show("Click \"Yes\"", "title", MessageBoxButtons.YesNoCancel);
            Assert.AreEqual(result, DialogResult.Yes);
            
        }
        [TestMethod]
        public void TestMethod2()
        {


        }



    }

}
