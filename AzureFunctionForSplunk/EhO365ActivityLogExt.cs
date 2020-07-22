using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctionForSplunk
{
    public static class EhO365ActivityLogExt
    {
        [FunctionName("EhO365ActivityLogExt")]
        public static async Task Run(
            [EventHubTrigger("%input-hub-name-activity-log%", Connection = "hubConnection", ConsumerGroup = "%consumer-group-activity-log%")] string[] messages,
            [EventHub("%output-hub-name-proxy%", Connection = "outputHubConnection")] IAsyncCollector<string> outputEvents,
            IBinder blobFaultBinder,
            IBinder incomingBatchBinder,
            Binder queueFaultBinder,
            ILogger log)
        {
            var runner = new Runner();
            await runner.Run<O365AuditMessages, ActivityLogsSplunkEventMessages>(messages, blobFaultBinder, queueFaultBinder, incomingBatchBinder, outputEvents, log);
        }
    }
}
