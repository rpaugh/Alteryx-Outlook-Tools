using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace OutlookTools
{
    public partial class OutlookEmail
    {
        private List<Item> _messages;

        public int RecordLimit { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public bool UseManualServiceURL { get; set; }
        public string ServiceURL { get; set; }
        public WellKnownFolderName Folder { get; set; }
        public string AttachmentPath { get; set; }
        public string QueryString { get; set; }
        public bool IncludeSubFolders { get; set; }
        public string SubFolderName { get; set; }
        public bool SkipRootFolder { get; set; }
        public bool UseUniqueFileName { get; set; }
        public PropertySet Fields { get; set; }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public List<Item> GetMessages(long recordLimit)
        {
            _messages = new List<Item>();

            ExchangeService service = new ExchangeService();
            List<Folder> folders = new List<Folder>();

            // Set specific credentials.
            service.Credentials = new WebCredentials(UserName, Password);

            if (UseManualServiceURL)
            {
                service.Url = new Uri(ServiceURL);
            }
            else
            {
                // Look up the user's EWS endpoint by using Autodiscover.
                service.AutodiscoverUrl(UserName, RedirectionCallback);
            }

            if (!SkipRootFolder)
            {
                // Get items from the selected root folder.
                GetItemsFromFolder(service, Folder, true);
            }

            // Search sub-folders if desired.
            if (IncludeSubFolders)
            {
                FolderView folderView = new FolderView(1000) { PropertySet = new PropertySet(BasePropertySet.IdOnly), Traversal = FolderTraversal.Deep };

                FindFoldersResults folderResults = null;

                do
                {
                    // Search for only a specific sub-folder if supplied, otherwise search all sub-folders for the selected root folder.
                    if (!String.IsNullOrWhiteSpace(SubFolderName))
                    {
                        folderResults = service.FindFolders(Folder, new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, SubFolderName), folderView);
                    }
                    else
                    {
                        folderResults = service.FindFolders(Folder, folderView);
                    }

                    foreach (var folder in folderResults)
                    {
                        folders.Add(folder);
                    }

                    folderView.Offset += folderResults.Folders.Count;
                }
                while (folderResults.MoreAvailable);


                foreach (var folder in folders)
                {
                    GetItemsFromFolder(service, folder, false);
                }
            }

            return _messages;
        }

        public void GetItemsFromFolder(ExchangeService service, object folder, bool isRoot)
        {
            // Create the item view limit based on the number of records requested from the Alteryx engine.
            ItemView itemView = new ItemView(1000) { Traversal = ItemTraversal.Shallow };
            FindItemsResults<Item> results = null;

            do
            {
                // Query items via EWS.
                results = service.FindItems(isRoot ? (WellKnownFolderName)folder : ((Folder)folder).Id, QueryString, itemView);

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

                    _messages.Add(message);
                }

                itemView.Offset += results.Items.Count;
            }
            while (results.MoreAvailable);
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

                    // Save the extracted attachment to the location specified in the tool's configuration UI.
                    string attachmentName = UseUniqueFileName ? String.Format("{0}_{1}{2}", Path.GetFileNameWithoutExtension(fileAttachment.Name), DateTime.Now.ToFileTime().ToString(), Path.GetExtension(fileAttachment.Name)) : fileAttachment.Name;
                    string file = String.Format("{0}\\{1}", AttachmentPath, attachmentName);
                    FileStream fs = new FileStream(file, FileMode.Create);

                    fileAttachment.Load(fs);

                    fs.Close();
                }
            }
        }
    }
}
