using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookTools
{
    public partial class OutlookEmail
    {
        public int RecordLimit { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public WellKnownFolderName Folder { get; set; }
        public string AttachmentPath { get; set; }
        public string QueryString { get; set; }
        public PropertySet Fields { get; set; }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public List<Item> GetMessages(long recordLimit)
        {
            ExchangeService service = new ExchangeService();
            List<Item> messages = new List<Item>();

            // Set specific credentials.
            service.Credentials = new NetworkCredential(UserName, Password);

            // Look up the user's EWS endpoint by using Autodiscover.
            service.AutodiscoverUrl(UserName, RedirectionCallback);

            // Create the item view limit based on the number of records requested from the Alteryx engine.
            ItemView itemView = new ItemView(1000);
            FindItemsResults<Item> results = null;

            do
            {
                // Query items via EWS.
                results = service.FindItems(Folder, QueryString, itemView);

                foreach (var item in results.Items)
                {
                    // Bind an email message and pull the specified set of properties.
                    Item message = Item.Bind(service, item.Id, Fields);

                    var attachments = string.Empty;

                    // Extract attachments from each message item if found.
                    if (!String.IsNullOrEmpty(AttachmentPath))
                    {
                        GetAttachmentsFromEmail(message);
                    }

                    messages.Add(message);
                }

                itemView.Offset += results.Items.Count;
            }
            while (results.MoreAvailable);

            return messages;
        }

        public void GetAttachmentsFromEmail(Item message)
        {
            // Iterate through the attachments collection and load each attachment.
            foreach (Attachment attachment in message.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = (FileAttachment)attachment;

                    MemoryStream ms = new MemoryStream();
                    
                    // Save the extracted attachment to the loation specified in the tool's configuration UI.
                    string file = String.Format("{0}\\{1}", AttachmentPath, fileAttachment.Name);
                    FileStream fs = new FileStream(file, FileMode.Create);

                    fileAttachment.Load(fs);

                    fs.Close();
                }
            }
        }
    }
}
