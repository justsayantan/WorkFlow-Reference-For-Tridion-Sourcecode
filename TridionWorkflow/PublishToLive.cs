//=======================================//
// ------- Author : Sayantan Basu -------
// ------- Date   : 08-01-2016    -------  
//=======================================//

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Tridion.ContentManager.CoreService.Client;
using Tridion.ContentManager.CoreService.Workflow;
using Tridion.Logging;
using System.Diagnostics;


namespace TridionWorkflow
{
    class PublishToLive : ExternalActivity
    {
        /// <summary>
        /// This method is used to Publish Item to live environment
        /// </summary>
        protected override void Execute()
        {
            PublishInstructionData publishInstractionData = new PublishInstructionData();
            publishInstractionData.RenderInstruction = new RenderInstructionData();
            publishInstractionData.ResolveInstruction = new ResolveInstructionData();
            publishInstractionData.ResolveInstruction.IncludeWorkflow = true;
            publishInstractionData.ResolveInstruction.IncludeChildPublications = false;

            List<string> itemToPublish = new List<string>();
            String[] targets = Utility.GetPublishingTarget("Live Publication Target");
            String[] PublishTo = Utility.GetPublishingTo();
            int length = PublishTo.Length;

            ActivityInstanceData activityInstance = ActivityInstance;

            foreach (WorkItemData wid in activityInstance.WorkItems)
            {
                if (length != 0 && PublishTo[0].ToString() != "All")
                {
                    for (int counter = 0; counter < length; counter++)
                    {
                        string comp = wid.Subject.IdRef;
                        int index = comp.IndexOf('-');
                        string sub1 = comp.Substring(index, comp.Length - index);
                        comp = @"tcm:" + PublishTo[counter].ToString() + sub1;
                        itemToPublish.Add(comp);
                    }
                }
                else
                {
                    publishInstractionData.ResolveInstruction.IncludeChildPublications = true;
                    itemToPublish.Add(wid.Subject.IdRef);
                }
            }

            PublishTransactionData[] publishTransactionData = CoreServiceClient.Publish(itemToPublish.ToArray<String>(), publishInstractionData, targets, PublishPriority.High, null);

            if (ProcessInstance.Variables.ContainsKey("PublishTransaction"))
            {
                ProcessInstance.Variables["PublishTransaction"] = publishTransactionData[0].Id;
            }
            else
            {
                ProcessInstance.Variables.Add("PublishTransaction", publishTransactionData[0].Id);
            }

            CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Publish to Live Queued: Finished Activity" }, null);
            Logger.Write(string.Format("Message: {0}", "Publish to Live Queued: Finished Activity"), "Workflow", LoggingCategory.General, TraceEventType.Information);
        }
    }
}
