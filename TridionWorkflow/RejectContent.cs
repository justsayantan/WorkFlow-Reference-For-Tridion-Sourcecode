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
using Tridion.ContentManager.Workflow;
using System.Net.Mail;


namespace TridionWorkflow
{
    class RejectContent : ExternalActivity
    {
        protected override void Execute()
        {
            PublishInstructionData publishInstruction = new PublishInstructionData();
            publishInstruction.ResolveInstruction = new ResolveInstructionData();
            publishInstruction.RenderInstruction = new RenderInstructionData();

            //Needed for publishing workflow revision/version
            publishInstruction.ResolveInstruction.IncludeWorkflow = true;
            TrusteeData lastPerformer = GetFirstManualActivityPerformer();
            ActivityInstanceData activityInstance = ActivityInstance;
            Logger.Write(string.Format("lastPerformer: {0}", lastPerformer.Title), "Workflow", LoggingCategory.General, TraceEventType.Information);

            if (Utility.IsMailSendOptionTrue(activityInstance.Title))
            {
                Logger.Write(string.Format("Mail Send Option: {0}", "True"), "Workflow", LoggingCategory.General, TraceEventType.Information);

                try
                {
                    SmtpClient client = Utility.SMTPClientConfiguration();
                    MailMessage mail = Utility.WorkflowMailMessageConfiguration(ActivityInstance.Title.ToString(), null, lastPerformer.Title);
                    client.Send(mail);
                    Logger.Write(string.Format("Mail : {0}", mail.Body.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
                    Logger.Write(string.Format("ActivityInstance.Title : {0}", "Mail Sent"), "Workflow", LoggingCategory.General, TraceEventType.Information);

                }
                catch (Exception ex)
                {
                    Logger.Write(string.Format("ActivityInstance.Title : {0}", ex.Message.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);

                }
            }

            if (Utility.IsPublishedToPreviewTrue())
            {
                // Retrieving the Publication Target and Publish Transaction
                if (ProcessInstance.Variables.ContainsKey("PublishTransaction"))
                {
                    string publishTransactionId = ProcessInstance.Variables["PublishTransaction"];
                    // Undo Publish Transaction
                    CoreServiceClient.UndoPublishTransaction(publishTransactionId, QueueMessagePriority.Normal, null);
                }
            }

            
            // Finish the Activity
                ActivityFinishData finishData = new ActivityFinishData()
                {
                    Message = "The Item " + activityInstance.WorkItems.FirstOrDefault().ToString() + " has been rejected and reassigned to " + lastPerformer.Title,
                    NextAssignee = new LinkToTrusteeData() { IdRef = lastPerformer.Id }
                };
                CoreServiceClient.FinishActivity(activityInstance.Id, finishData, null); 
            


        }


        private UserData GetFirstManualActivityPerformer()
        {
            ReadOptions readoption = new ReadOptions();
            ActivityInstanceData firstManualActivity = GetFirstManualActivity();
            Logger.Write(string.Format("FirstManualActivity: {0}", firstManualActivity.Title), "Workflow", LoggingCategory.General, TraceEventType.Information);
            return (UserData)CoreServiceClient.Read(firstManualActivity.Performers.Last().IdRef, null);
        }


        private ActivityInstanceData GetFirstManualActivity()
        {
            IEnumerable<ActivityInstanceData> activityInstances =
            ProcessInstance.Activities.OfType<ActivityInstanceData>().OrderBy(o => o.StartDate);

            return activityInstances.First(a =>
            {
                TridionActivityDefinitionData activityDefinition = (TridionActivityDefinitionData)CoreServiceClient.Read(a.ActivityDefinition.IdRef, null);
                return string.IsNullOrEmpty(activityDefinition.Script);
            });
        }
    }
}
