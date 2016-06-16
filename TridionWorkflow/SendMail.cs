//=======================================//
// ------- Author : Sayantan Basu -------
// ------- Date   : 05-01-2016    -------  
//=======================================//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Configuration;
using Tridion.ContentManager.CoreService.Workflow;
using Tridion.ContentManager.CoreService.Client;
using Tridion.Logging;
using System.Diagnostics;
using System.ComponentModel;
using Tridion.ContentManager.Workflow;


namespace TridionWorkflow
{
    class SendMail :ExternalActivity
    {
        //public string nextActivityId;
        ReadOptions readoption = new ReadOptions();

        protected override void Execute()
        {
            ActivityInstanceData activityInstance = ActivityInstance;
            List<String> items = new List<string>();
            TridionActivityDefinitionData activitydefinition = (TridionActivityDefinitionData)CoreServiceClient.Read(ActivityInstance.ActivityDefinition.IdRef, readoption);
            ProcessDefinitionData processdefinition = (ProcessDefinitionData)CoreServiceClient.Read(activitydefinition.ProcessDefinition.IdRef, readoption);
            Logger.Write(string.Format("ActivityInstance.Title : {0}", ActivityInstance.Title), "Workflow", LoggingCategory.General, TraceEventType.Information);
            
            foreach (WorkItemData wid in activityInstance.WorkItems)
            {
                items.Add(wid.Subject.Title);
            }

            try
            {
                SmtpClient client = Utility.SMTPClientConfiguration();
                MailMessage mail = Utility.WorkflowMailMessageConfiguration(ActivityInstance.Title.ToString(), items, null);
                client.Send(mail);
                Logger.Write(string.Format("Mail : {0}", mail.Body.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
                Logger.Write(string.Format("ActivityInstance.Title : {0}", "Mail Sent"), "Workflow", LoggingCategory.General, TraceEventType.Information);

            }
            catch (Exception ex)
            {
                Logger.Write(string.Format("ActivityInstance.Title : {0}", ex.Message.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);

            }
            finally
            {
                CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Mail Sent to Target Audience, Finished Activity"}, null);
                Logger.Write(string.Format("Message: {0}", "Auto Approved and Send for Next Activity , Finished Activity"), "Workflow", LoggingCategory.General, TraceEventType.Information);
            }
        }

    }
}
