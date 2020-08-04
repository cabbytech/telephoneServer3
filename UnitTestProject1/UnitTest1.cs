using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TelephoneServer3;
using Traysoft.AddTapi;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {

            var e = new EventArgs() as TapiEventArgs;
          
                
            Main m = new Main();
            m.Show();
            m.OnCallDisconnected(new object(),e as TapiEventArgs);

        }
    }
}
