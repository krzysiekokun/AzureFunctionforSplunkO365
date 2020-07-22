using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace AzureFunctionForSplunk
{
    public class O365AuditMessages : AzMonMessages
    {
        public O365AuditMessages(ILogger log) : base(log) { }

        public override List<string> DecomposeIncomingBatch(string[] messages)
        {
            List<string> decomposed = new List<string>();
            Log.LogInformation($"Received  {messages.Length} messages for decomposition");

            foreach (var message in messages)
            {
                List<Dictionary<string, dynamic>> l = new List<Dictionary<string, dynamic>>();

                if (message.TrimStart().StartsWith('['))
                {
                    //we received array of objects
                    dynamic obj = JsonConvert.DeserializeObject<List<Dictionary<string, dynamic>>>(message);

                    foreach(var item in obj)
                    {
                        RemoveEmptyProperties(item);
                        decomposed.Add(JsonConvert.SerializeObject((dynamic)item).ToString());

                    }
                }
                else
                {
                    //received single object
                    dynamic obj = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(message);
                    RemoveEmptyProperties(obj);
                    decomposed.Add(JsonConvert.SerializeObject((dynamic)obj).ToString());

                }


            }
            Log.LogInformation($"Decomposed {decomposed.Count} messages");
            return decomposed;
        }

        private void RemoveEmptyProperties(dynamic item)
        {
            var elements = item as Dictionary<string, dynamic>;
            if (elements != null)
            {
                List<string> keysToRemove = new List<string>();
                foreach (var pair in elements)
                {
                    if (pair.Value != null && string.IsNullOrEmpty(pair.Value.ToString()))
                    {
                        keysToRemove.Add(pair.Key);
                    }
                }
                foreach (string key in keysToRemove)
                {
                    elements.Remove(key);
                }
                Log.LogInformation($"Message decompostion, removed {keysToRemove.Count} elements");
            }
        }
    }
}