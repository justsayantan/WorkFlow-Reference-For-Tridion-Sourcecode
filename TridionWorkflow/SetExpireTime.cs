//=======================================//
// ------- Author : Sayantan Basu -------
// ------- Date   : 09-02-2016    -------  
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
    class SetExpireTime : ExternalActivity
    {
        protected override void Execute()
        {
            CoreServiceClient.FinishActivity(ActivityInstance.Id, new ActivityFinishData { Message = "Mail Sent to Target Audience, Finished Activity", NextActivityDueDate = System.DateTime.Now.AddMinutes(5) }, null);
            Logger.Write(string.Format("Message: {0}", "Auto Approved and Send for Next Activity , Finished Activity"), "Workflow", LoggingCategory.General, TraceEventType.Information);
        } 
    }
}
