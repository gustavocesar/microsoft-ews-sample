using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using EWS_Sample.Config;
using Microsoft.Exchange.WebServices.Data;

namespace EWS_Sample
{
    class Program
    {
        private static ExchangeService _service = StartupService();
        private static ItemView _view = new ItemView(ExchangeServiceConfig.PageSize, ExchangeServiceConfig.Offset, OffsetBasePoint.Beginning) { PropertySet = PropertySet.IdOnly };

        private static ExchangeService StartupService()
        {
            var service = new ExchangeService();
            service.UseDefaultCredentials = ExchangeServiceConfig.UseDefaultCredentials;

            if (ExchangeServiceConfig.IsGenericBox)
            {
                service.Url = new Uri(ExchangeServiceConfig.HostGenericBox);
                service.Credentials = new WebCredentials(ExchangeServiceConfig.Username, ExchangeServiceConfig.Password);
            }
            else
            {
                service.Url = new Uri(ExchangeServiceConfig.HostOffice365);
                service.Credentials = new NetworkCredential(ExchangeServiceConfig.Username, ExchangeServiceConfig.Password);
            }

            return service;
        }
        
        static void Main(string[] args)
        {
            Console.WriteLine();
            Console.WriteLine("*** Searching Mail Box ***");

            foreach (var subject in ExchangeServiceConfig.FindSubjects)
                SearchMailbox(subject);
        }

        private static void SearchMailbox(string subject)
        {
            //Search filter
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subject));
            FindItemsResults<Item> findResults = _service.FindItems(new FolderId(WellKnownFolderName.Inbox, ExchangeServiceConfig.Username), searchFilter, _view);

            if (!findResults.Items.Any())
            {
                Console.WriteLine("No messages found!");
                return;
            }

            var emails = new List<EmailMessage>();
            foreach (var item in findResults.Items)
                emails.Add((EmailMessage)item);

            var properties = (BasePropertySet.FirstClassProperties);
            _service.LoadPropertiesForItems(emails, properties);

            //Each email found 
            foreach (var email in emails)
            {
                Console.WriteLine("Subject: " + email.Subject + " (de " + email.DateTimeReceived.ToString("dd/MM/yyyy HH:mm:ss") + ")");
                
                DownloadAttachments(email);
                ExportEmail(email);
            }
        }

        private static void ExportEmail(Item item)
        {
            PropertySet props = new PropertySet(EmailMessageSchema.MimeContent);
            
            // This results in a GetItem call to EWS.
            var email = EmailMessage.Bind(_service, item.Id, props);

            var timestamp = item.DateTimeReceived.ToString("yyyy-MM-dd-HH-mm");
            var subject = Regex.Replace(item.Subject, "[^0-9a-zA-Z]+", "").ToLower();
            var name = $"{timestamp}_{subject}";

            // Save as .eml.
            var emlFileName = $"./Export/{name}.eml";
            using (FileStream fs = new FileStream(emlFileName, FileMode.Create, FileAccess.Write))
            {
                fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                Console.WriteLine($"  Exported Email: {emlFileName}");
            }
        }

        private static void DownloadAttachments(Item item)
        {
            if (!item.Attachments.Any())
                Console.WriteLine("  No attachments");

            foreach (var attachment in item.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;
                    var file = $"./Download/{fileAttachment.Name}";
                    fileAttachment.Load(file);
                    Console.WriteLine($"  Attachment: {file}");
                }
            }
        }
    }
}

