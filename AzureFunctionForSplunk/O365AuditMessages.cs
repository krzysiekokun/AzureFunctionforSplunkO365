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

            foreach (var message in messages)
            {
                List<Dictionary<string, dynamic>> l = new List<Dictionary<string, dynamic>>();

                if (message.TrimStart().StartsWith('['))
                {
                    //we received array of objects
                    dynamic obj = JsonConvert.DeserializeObject<List<Dictionary<string, dynamic>>>(message);

                    foreach(var item in obj)
                    {
                        var elemets = item as Dictionary<string, dynamic>;
                        if (elemets != null)
                        {
                            List<string> keysToRemove = new List<string>();
                            foreach(var pair in elemets)
                            {
                                if(pair.Value != null && string.IsNullOrEmpty(pair.Value.ToString()))
                                {
                                    keysToRemove.Add(pair.Key);
                                }
                            }
                            foreach(string key in keysToRemove)
                            {
                                elemets.Remove(key);
                            }
                            decomposed.Add(JsonConvert.SerializeObject(elemets));
                        }
                    }
                }
                else
                {
                    //received single object
                    decomposed.Add(message);
                }






            }
            Log.LogInformation($"Decomposed {decomposed.Count} messages");
            return decomposed;
        }
    }
}