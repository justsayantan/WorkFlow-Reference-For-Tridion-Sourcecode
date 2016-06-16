//=======================================//
// ------- Author : Sayantan Basu -------
// ------- Date   : 28-01-2016    -------  
//=======================================//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Configuration;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.Reflection;
using Tridion.Logging;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace TridionWorkflow
{
    public static class Utility
    {
        #region Public Method

        /// <summary>
        /// 
        /// </summary>
        /// <param name="targetType"></param>
        /// <returns></returns>
        public static string[] GetPublishingTarget(string targetType)
        {
            AppSettingsSection section = GetAppSettings();
            string _target = section.Settings[targetType].Value;
            string[] target = new[] { _target };
            return target;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string[] GetPublishingTo()
        {
            string[] publishTo;
            AppSettingsSection section = GetAppSettings();
            string _publishTo = section.Settings["Publish To PublicationID"].Value;
            if (_publishTo.Contains(','))
            {
                publishTo = _publishTo.Split(',');
            }
            else
            {
                publishTo = new[] { _publishTo };
            }
            return publishTo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static SmtpClient SMTPClientConfiguration()
        {
            AppSettingsSection section = GetAppSettings();

            SmtpClient client = new SmtpClient();

            //Get variable from App config
            string _MailServer = section.Settings["Mail Server"].Value;
            string _AuthenticationID = section.Settings["Mail Server Authentication ID"].Value;
            string _AuthenticationPassword = section.Settings["Mail Server Authentication Password"].Value;

            //Set value to smtp client
            client.Host = _MailServer;
            client.UseDefaultCredentials = false;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Credentials = new NetworkCredential(_AuthenticationID, _AuthenticationPassword);

            return client;
        }

        public static bool IsMailSendOptionTrue(string activityInstance)
        {
            AppSettingsSection section = GetAppSettings();
            bool result = false;
            string myactivity = GetActivityInstance(activityInstance);
            string checkMailSendOption = "Mail Send Option" + myactivity;
            Logger.Write(string.Format("checkMailSendOption : {0}", checkMailSendOption), "Workflow", LoggingCategory.General, TraceEventType.Information);
                    
            string _checkMailSendOption = section.Settings[checkMailSendOption].Value;
            if (_checkMailSendOption == "true")
            {
                result = true;
            }
            return result;
        }

        public static bool IsPublishedToPreviewTrue()
        {
            AppSettingsSection section = GetAppSettings();
            bool result = false;
            string checkPublishedToPreviewTrue = "Published To Preview";

            string _checkPublishedToPreviewTrue = section.Settings[checkPublishedToPreviewTrue].Value;
            if (_checkPublishedToPreviewTrue == "true")
            {
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static MailMessage WorkflowMailMessageConfiguration(string activityInstance, List<string> items, string lastPerformer)
        {
            AppSettingsSection section = GetAppSettings();
            /// <summary>
            /// Get Mail Settings
            /// </summary>
            string _MailTo;
            string msgXML = CreateXML(items);
            string myactivity = GetActivityInstance(activityInstance);
            string Mailto = "Mail To" + myactivity;
            string MessageBodyXslt = "MessageBodyXslt" + myactivity;
            Logger.Write(string.Format("MessageBodyXslt Name: {0}", MessageBodyXslt), "Workflow", LoggingCategory.General, TraceEventType.Information);
            
            if (lastPerformer != null)
            {
                string PerformerName = lastPerformer;
                if (PerformerName.Contains("\\"))
                {
                    PerformerName = GetPerformerName(PerformerName);
                }
                string _domainName = section.Settings["Mail Domain Name"].Value;
                _MailTo = PerformerName + "@" + _domainName;
            }
            else
            {
                _MailTo = section.Settings[Mailto].Value;
            }
            string _MessageBodyXslt = section.Settings[MessageBodyXslt].Value;        

            string _MessageSubject = section.Settings["Message Subject"].Value;
            string _MailFrom = section.Settings["Mail From"].Value;
            Logger.Write(string.Format("Mail T0, Mail Form : {0} {1}", _MailTo, _MailFrom), "Workflow", LoggingCategory.General, TraceEventType.Information);

            MailMessage msg = new MailMessage(_MailFrom, _MailTo);

            msg.IsBodyHtml = true;
            msg.Body = GetHtmlBody(_MessageBodyXslt, msgXML);
            msg.Subject = _MessageSubject;

            return msg;
        }    

        

        #endregion

        #region Private Method

        private static string GetPerformerName(string lastPerformer)
        {
            int index = lastPerformer.IndexOf('\\')+1;
            int stringLength = lastPerformer.Length;
            string result = lastPerformer.Substring(index, stringLength-index);

            return result;
        }

        /// <summary>
        /// Find and Read the App Config path from the Server
        /// </summary>
        /// <returns></returns>
        private static AppSettingsSection GetAppSettings()
        {
            System.Configuration.ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            String EnviromentPath = System.Environment.GetEnvironmentVariable("TRIDION_HOME", EnvironmentVariableTarget.Machine);
            configFileMap.ExeConfigFilename = EnviromentPath + @"config\Workflow.config";
                        
            System.Configuration.Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            AppSettingsSection section = (AppSettingsSection)configuration.GetSection("appSettings");
            return section;
        }

        /// <summary>
        /// Read the Body of the mail from XSLT and populate the HTML 
        /// </summary>
        /// <param name="messageBodyXslt"></param>
        /// <returns></returns>
        private static string GetHtmlBody(string messageBodyXslt, string xmlMsg)
        {
            string xmlInput = xmlMsg;
            if (xmlMsg == null)
            {
                xmlInput = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
                                <Workflow><Items>
                                </Items></Workflow>";
            }
            

            string xslInput = System.IO.File.ReadAllText(messageBodyXslt);
            Logger.Write(string.Format("xmlInput : {0}", xmlInput), "Workflow", LoggingCategory.General, TraceEventType.Information);
            Logger.Write(string.Format("xslInput : {0}", xslInput), "Workflow", LoggingCategory.General, TraceEventType.Information);
            using (StringReader srt = new StringReader(xslInput)) // xslInput is a string that contains xsl
            using (StringReader sri = new StringReader(xmlInput)) // xmlInput is a string that contains xml
            {
                using (XmlReader xrt = XmlReader.Create(srt))
                using (XmlReader xri = XmlReader.Create(sri))
                {
                    XslCompiledTransform xslt = new XslCompiledTransform();
                    xslt.Load(xrt);
                    using (StringWriter sw = new StringWriter())
                    using (XmlWriter xwo = XmlWriter.Create(sw, xslt.OutputSettings)) // use OutputSettings of xsl, so it can be output as HTML
                    {
                        xslt.Transform(xri, xwo);
                        Logger.Write(string.Format("Mail String : {0}", sw.ToString()), "Workflow", LoggingCategory.General, TraceEventType.Information);
                        return sw.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Read the next Activity Instance and prepare a String to get the approver name 
        /// </summary>
        /// <param name="activityInstance"></param>
        /// <returns></returns>
        private static string GetActivityInstance(string activityInstance)
        {
            string input = activityInstance;
            var output = Regex.Replace(input, @"[\d.]", string.Empty);

            return output.ToString();
        }

        private static string CreateXML(List<string> items)
        {
            string xmlInput = null;
            if (items != null)
            {
                string xmlHead = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
                                <Workflow><Items>";
                string xmlEnd = @"</Items></Workflow>";
                string xmlBody = null;
                foreach (var item in items)
                {
                    xmlBody = xmlBody + @"<Item>" + item + @"</Item>";
                }

                xmlInput = xmlHead + xmlBody + xmlEnd;               
            }
            return xmlInput;
        }

        #endregion

        
    }
}
