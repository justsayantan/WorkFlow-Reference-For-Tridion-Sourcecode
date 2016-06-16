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
    class BundlePublishToPreview : ExternalActivity
    {
        protected override void Execute()
        {
            PublishInstructionData publishInstruction = new PublishInstructionData();
            publishInstruction.ResolveInstruction = new ResolveInstructionData();
            publishInstruction.RenderInstruction = new RenderInstructionData();

            //Needed for publishing workflow revision/version
            publishInstruction.ResolveInstruction.IncludeWorkflow = true;
            ActivityInstanceData activityInstance = ActivityInstance;
            IList<String> itemsToPublishList = new List<String>();

            //Staging publication target URI
            String[] targets = Utility.GetPublishingTarget("Preview Publication Target");

            foreach (WorkItemData wid in activityInstance.WorkItems)
            {
                int value = Convert.ToInt32(Enum.Parse(typeof(ItemType), "VirtualFolder"));
                if (wid.Subject.IdRef.EndsWith(value.ToString()))
                {
                    itemsToPublishList.Add(wid.Subject.IdRef);
                }
            }

            //PublishTransactionData requires reference to System.ServiceModel
            PublishTransactionData[] publishTransactions = CoreServiceClient.Publish(itemsToPublishList.ToArray<String>(), publishInstruction, targets, PublishPriority.Normal, null);

            //Store the publish transaction id so that we can undo if needed!
            if (ProcessInstance.Variables.ContainsKey("PublishTransaction"))
            {
                ProcessInstance.Variables["PublishTransaction"] = publishTransactions[0].Id;
            }
            else
            {
                ProcessInstance.Variables.Add("PublishTransaction", publishTransactions[0].Id);
            }

            CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Publish to Staging Queued: Finished Activity" }, null);
        }
    }
}
