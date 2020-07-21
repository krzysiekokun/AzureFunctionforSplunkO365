using AzureFunctionForSplunk;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace AzureFunctionForSplunkTests
{

    [TestClass]
    public class ActivityMessagesTests
    {
        [TestMethod]
        [DeploymentItem(@"message1.txt", "optionalOutFolder")]
        public void CheckIncommingMessageCorrectStructure()
        {
            var azMonMsgs = (AzMonMessages)Activator.CreateInstance(typeof(ActivityLogMessages), TestUtils.CreateLogger(LoggerTypes.Null));
            var result = azMonMsgs.DecomposeIncomingBatch(new string[] { File.ReadAllText("message1.txt") });

        }

        
    }
}
