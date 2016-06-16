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
using System.Net.Mail;

namespace TridionWorkflow
{
    class ExpirationScript :ExternalActivity
    {
        ReadOptions readoption = new ReadOptions();        
        /// <summary>
        /// This method is called whenthe activity is expired. It moves the content to the next activity
        /// </summary>
        protected override void Expire()
        {
            ActivityInstanceData activityInstance = ActivityInstance;
            List<String> items = new List<string>();
            TridionActivityDefinitionData activitydefinition = (TridionActivityDefinitionData)CoreServiceClient.Read(ActivityInstance.ActivityDefinition.IdRef, readoption);
            ProcessDefinitionData processdefinition = (ProcessDefinitionData)CoreServiceClient.Read(activitydefinition.ProcessDefinition.IdRef, readoption);
            Logger.Write(string.Format("ActivityInstance.Title : {0}", ActivityInstance.Title), "Workflow", LoggingCategory.General, TraceEventType.Information);
            Logger.Write(string.Format("Due Date : {0}", activityInstance.DueDate.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
            Logger.Write(string.Format("Start Date : {0}", activityInstance.StartDate.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
            Logger.Write(string.Format("Finish Date : {0}", activityInstance.FinishDate.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);

            foreach (WorkItemData wid in activityInstance.WorkItems)
            {
                items.Add(wid.Subject.Title);
            }

            try
            {
                activityInstance.DueDate = System.DateTime.Now.AddMinutes(2);
                Logger.Write(string.Format("New DueDate : {0}", activityInstance.DueDate), "Workflow", LoggingCategory.General, TraceEventType.Information);
                
                SmtpClient client = Utility.SMTPClientConfiguration();
                MailMessage mail = Utility.WorkflowMailMessageConfiguration(ActivityInstance.Title.ToString() + "ExpireDueDate", items, null);
                client.Send(mail);
                Logger.Write(string.Format("Mail : {0}", mail.Body.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
                Logger.Write(string.Format("Message : {0}", "Mail Sent"), "Workflow", LoggingCategory.General, TraceEventType.Information);

            }
            catch (Exception ex)
            {
                Logger.Write(string.Format("Exception Happend : {0}", ex.Message.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);

            }
            finally
            {
                activityInstance.DueDate = System.DateTime.Now.AddMinutes(2);
                ActivityInstance.ActivityState = ActivityState.Finished;
                CoreServiceClient.RestartActivity(ActivityInstance.Id, null);
                //activityInstance.DueDate = System.DateTime.Now.AddMinutes(2);
                //Logger.Write(string.Format("New Due Date Time : {0}", activityInstance.DueDate.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
                ////CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Mail Sent to Target Audience, Finished Activity" }, null);
                //Logger.Write(string.Format("Message: {0}", "Auto Approved and Send for Next Activity , Finished Activity"), "Workflow", LoggingCategory.General, TraceEventType.Information);
            }
           // CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Auto Approved and Send for Next Activity , Finished Activity" }, null);
           // Logger.Write(string.Format("Message: {0}", "Activity Expired and Send for Next Activity , Finished Activity"), "Workflow", LoggingCategory.General, TraceEventType.Information);
        }
    }
}
