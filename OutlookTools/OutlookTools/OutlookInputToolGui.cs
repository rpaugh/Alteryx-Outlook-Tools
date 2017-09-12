using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookTools
{
    // The user interface that allows the user to configure the tool in the Alteryx GUI.
    // The class must be a UserControl and must also implement the IPluginConfiguration interface
    // found in the AlteryxGuiToolkit.dll.  This interface includes two methods:  
    // GetConfigurationControl - updates the control with the xml configuration and returns it to Alteryx
    // SaveResultsToXml - saves the configuration to xml
    public partial class OutlookInputToolGui : UserControl, AlteryxGuiToolkit.Plugins.IPluginConfiguration
    {
        public OutlookInputToolGui()
        {
            InitializeComponent();

            List<KeyValuePair<int, string>> folderList = Enum<WellKnownFolderName>.ToList();

            cboFolderToSearch.DataSource = folderList;
            cboFolderToSearch.DisplayMember = "Value";
            cboFolderToSearch.ValueMember = "Key";
            cboFolderToSearch.SelectedValue = folderList.Where(x => x.Value == "Inbox").First().Key;

            clbFields.DataSource = new BindingSource(Member<ItemSchema>.ToList(), null);
            clbFields.DisplayMember = "Value";
            clbFields.ValueMember = "Key";
            
            lblQueryStringHelpLink.Links[0].LinkData = "https://msdn.microsoft.com/en-us/library/ee693615(exchg.140).aspx";
        }

        public Control GetConfigurationControl(AlteryxGuiToolkit.Document.Properties docProperties, XmlElement eConfig, XmlElement[] eIncomingMetaInfo, int nToolId, string strToolName)
        {
            // This method is called by Alteryx to initialize this configuration control from data
            // stored in the module containing this tool.

            // Tool configuration is handled through an XML document managed by Alteryx.  When it is
            // time to initialize this control (generally when it is made visible to the user), this
            // method is called, passing the XML document in with the eConfig parameter.

            // When this tool is connected to one or more upstream data providers, the structure of
            // each data stream is passed as an entry in the eIncomingMetaInfo array.  This data can
            // then be used to dynamically construct your UI based on the type of input.

            // Since this tool is an input tool, the eIncomingMetaInfo parameter will be empty.
            // We will use the information in eConfig to populate the XmlFile text box.  This will
            // cause the file to be examined if it exists (through the XmlFile_TextChanged event
            // handler) in order to populate the sample information.  We will then use the remaining
            // information in eConfig to set the field properties if that information exists.

            // Call LoadFromConfiguration to get the xml file name and field information from eConfig.
            XmlInputConfiguration xmlConfig = XmlInputConfiguration.LoadFromConfiguration(eConfig);

            if (xmlConfig != null)
            {
                // Update the XmlFile and ElementName textboxes.
                txtUserName.Text = xmlConfig.UserName;
                txtPassword.Text = xmlConfig.Password;
                chkUseManualServiceURL.Checked = xmlConfig.UseManualServiceURL;

                if (xmlConfig.UseManualServiceURL)
                {
                    txtServiceURL.Enabled = true;
                }
                else
                {
                    txtServiceURL.Enabled = false;
                }

                txtServiceURL.Text = xmlConfig.ServiceURL;
                cboFolderToSearch.SelectedValue = xmlConfig.Folder;
                txtAttachmentPath.Text = xmlConfig.AttachmentPath;
                txtQueryString.Text = xmlConfig.QueryString;
                chkIncludeSubFolders.Checked = xmlConfig.IncludeSubFolders;

                if (xmlConfig.IncludeSubFolders)
                {
                    txtSubFolderName.Enabled = true;
                }
                else
                {
                    txtSubFolderName.Enabled = false;
                }

                txtSubFolderName.Text = xmlConfig.SubFolderName;
                chkSkipRootFolder.Checked = xmlConfig.SkipRootFolder;
                chkUseUniqueFileName.Checked = xmlConfig.UseUniqueFileName;

                foreach (XmlInputField field in xmlConfig.Fields)
                {
                    // Use a little LINQ to find the row in the field list that
                    // corresponds to the current field.
                    var item = clbFields.Items.Cast<KeyValuePair<string, string>>().Where(x => x.Value == field.Name).FirstOrDefault();
                    
                    if (clbFields.Items.IndexOf(item) != -1)
                    {
                        clbFields.SetItemChecked(clbFields.Items.IndexOf(item), true);
                    }
                }
            }

            return this;
        }

        public void SaveResultsToXml(XmlElement eConfig, out string strDefaultAnnotation)
        {
            // This method is called by Alteryx to retrieve the (possibly) updated XML document
            // representing the configuration of the tool.  This information is then stored back
            // in the Alteryx module.

            // All information which should be stored and thus made available to the Engine
            // associated with this tool must be added into the eConfig element at this time.

            // The strDefaultAnnotation represents the text to be displayed as a label next to
            // the tool in the flowchart view of the module, and should be set to some appropriate
            // string, if possible.

            // Set the default annotation to be the user's email address.
            strDefaultAnnotation = txtUserName.Text;
            
            // Assign the configuration values from the UI to their corresponding XML config elements.
            XmlElement userName = XmlHelpers.GetOrCreateChildNode(eConfig, "UserName");
            userName.InnerText = txtUserName.Text;
            
            XmlElement password = XmlHelpers.GetOrCreateChildNode(eConfig, "Password");
            password.InnerText = txtPassword.Text;

            XmlElement useManualServiceURL = XmlHelpers.GetOrCreateChildNode(eConfig, "UseManualServiceURL");
            useManualServiceURL.InnerText = chkUseManualServiceURL.Checked.ToString();

            XmlElement serviceURL = XmlHelpers.GetOrCreateChildNode(eConfig, "ServiceURL");
            serviceURL.InnerText = txtServiceURL.Text;

            XmlElement folder = XmlHelpers.GetOrCreateChildNode(eConfig, "Folder");
            folder.InnerText = cboFolderToSearch.SelectedValue.ToString();

            XmlElement attachmentPath = XmlHelpers.GetOrCreateChildNode(eConfig, "AttachmentPath");
            attachmentPath.InnerText = txtAttachmentPath.Text;

            XmlElement queryString = XmlHelpers.GetOrCreateChildNode(eConfig, "QueryString");
            queryString.InnerText = txtQueryString.Text;

            XmlElement includeSubFolders = XmlHelpers.GetOrCreateChildNode(eConfig, "IncludeSubFolders");
            includeSubFolders.InnerText = chkIncludeSubFolders.Checked.ToString();

            XmlElement subFolderName = XmlHelpers.GetOrCreateChildNode(eConfig, "SubFolderName");
            subFolderName.InnerText = txtSubFolderName.Text;

            XmlElement skipRootFolder = XmlHelpers.GetOrCreateChildNode(eConfig, "SkipRootFolder");
            skipRootFolder.InnerText = chkSkipRootFolder.Checked.ToString();

            XmlElement useUniqueFileName = XmlHelpers.GetOrCreateChildNode(eConfig, "UseUniqueFileName");
            useUniqueFileName.InnerText = chkUseUniqueFileName.Checked.ToString();

            // Create a Fields element to contain the information for each output field.
            XmlElement fieldsElement = XmlHelpers.GetOrCreateChildNode(eConfig, "Fields");
            fieldsElement.RemoveAll();

            foreach (var field in clbFields.CheckedItems)
            {
                // Create a new Field element for each field and set the Name, Type,
                // Size and Scale attributes from the field grid.
                XmlElement fieldElement = XmlHelpers.CreateElement("Field", fieldsElement);
                XmlAttribute nameAttribute = XmlHelpers.CreateAttribute("Name", fieldElement);
                nameAttribute.Value = GetValueString(((KeyValuePair<string, string>)field).Value);
                XmlAttribute typeAttribute = XmlHelpers.CreateAttribute("Type", fieldElement);
                typeAttribute.Value = GetValueString(AlteryxRecordInfoNet.FieldType.E_FT_V_WString);
                XmlAttribute sizeAttribute = XmlHelpers.CreateAttribute("Size", fieldElement);
                sizeAttribute.Value = GetValueString("1073741823");
                XmlAttribute scaleAttribute = XmlHelpers.CreateAttribute("Scale", fieldElement);
                scaleAttribute.Value = GetValueString(null);
            }

        }

        // GetValueString is a helper function that deals with potential null values
        // coming out of the source.
        string GetValueString(object value)
        {
            if (value == null) return "";
            else return value.ToString();
        }

        private void btnAttachmentPath_Click(object sender, EventArgs e)
        {
            if (fbAttachmentPath.ShowDialog() == DialogResult.OK)
            {
                txtAttachmentPath.Text = fbAttachmentPath.SelectedPath;
            }
        }

        private void lblQueryStringHelpLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        }

        private void chkIncludeSubFolders_Click(object sender, EventArgs e)
        {
            if (chkIncludeSubFolders.Checked)
            {
                txtSubFolderName.Enabled = true;
            }
            else
            {
                txtSubFolderName.Enabled = false;
            }
        }

        private void chkUseManualServiceURL_Click(object sender, EventArgs e)
        {
            if (chkUseManualServiceURL.Checked)
            {
                txtServiceURL.Enabled = true;
            }
            else
            {
                txtServiceURL.Enabled = false;
            }
        }
    }

    // The XmlInputConfiguration class is used to parse an XML file to determine what
    // data fields are contained in the elements with the specified containerElementName
    public class XmlInputConfiguration
    {
        public string UserName { get; private set; }
        public string Password { get; private set; }
        public bool UseManualServiceURL { get; private set; }
        public string ServiceURL { get; private set; }
        public int Folder { get; private set; }
        public string AttachmentPath { get; private set; }
        public string QueryString { get; private set; }
        public int RecordLimit { get; private set; }
        public bool IncludeSubFolders { get; private set; }
        public string SubFolderName { get; private set; }
        public bool SkipRootFolder { get; private set; }
        public bool UseUniqueFileName { get; private set; }
        public List<XmlInputField> Fields { get; private set; }

        // Note that the constructor is private.  Instances are created through the
        // LoadFromConfigration method.
        XmlInputConfiguration(string userName, string password, bool useManualServiceURL, string serviceURL, int folder, string attachmentPath, string queryString, bool includeSubFolders, string subFolderName, bool skipRootFolder, bool useUniqueFileName)
        {
            UserName = userName;
            Password = password;
            UseManualServiceURL = useManualServiceURL;
            ServiceURL = serviceURL;
            Folder = folder;
            AttachmentPath = attachmentPath;
            QueryString = queryString;
            IncludeSubFolders = includeSubFolders;
            SubFolderName = subFolderName;
            SkipRootFolder = skipRootFolder;
            UseUniqueFileName = useUniqueFileName;
            Fields = new List<XmlInputField>();
        }

        void AddField(string name, AlteryxRecordInfoNet.FieldType type, int size, int scale)
        {
            Fields.Add(new XmlInputField()
            {
                Name = name,
                FieldType = type,
                Size = size,
                Scale = scale
            });
        }

        // Creates a new instance of the XmlInputConfiguration class based on the information
        // in the eConfig xml element.  This is the eConfig that Alteryx provides in the
        // IPluginConfiguration.GetConfigurationControl() method.
        public static XmlInputConfiguration LoadFromConfiguration(XmlElement eConfig)
        {
            // Get the configuration values from the XML config elements to be place in their corresponding fields in the UI.
            XmlElement userName = eConfig.SelectSingleNode("UserName") as XmlElement;
            
            XmlElement password = eConfig.SelectSingleNode("Password") as XmlElement;

            XmlElement useManualServiceURL = eConfig.SelectSingleNode("UseManualServiceURL") as XmlElement;

            XmlElement serviceURL = eConfig.SelectSingleNode("ServiceURL") as XmlElement;

            XmlElement folder = eConfig.SelectSingleNode("Folder") as XmlElement;

            XmlElement attachmentPath = eConfig.SelectSingleNode("AttachmentPath") as XmlElement;

            XmlElement queryString = eConfig.SelectSingleNode("QueryString") as XmlElement;

            XmlElement includeSubFolders = eConfig.SelectSingleNode("IncludeSubFolders") as XmlElement;

            XmlElement subFolderName = eConfig.SelectSingleNode("SubFolderName") as XmlElement;

            XmlElement skipRootFolder = eConfig.SelectSingleNode("SkipRootFolder") as XmlElement;

            XmlElement useUniqueFileName = eConfig.SelectSingleNode("UseUniqueFileName") as XmlElement;

            if (userName != null && password != null)
            {
                // Create the new XmlInputConfiguration object.
                XmlInputConfiguration xmlConfig = new XmlInputConfiguration(userName.InnerText, password.InnerText, Convert.ToBoolean(useManualServiceURL.InnerText), serviceURL.InnerText, Convert.ToInt16(folder.InnerText), attachmentPath.InnerText, queryString.InnerText, Convert.ToBoolean(includeSubFolders.InnerText), subFolderName.InnerText, Convert.ToBoolean(skipRootFolder.InnerText), Convert.ToBoolean(useUniqueFileName.InnerText));

                // Find all of the Field elements in the configuration.
                XmlNodeList fields = eConfig.SelectNodes("Fields/Field");
                foreach (XmlElement fieldElement in fields)
                {
                    // For each field element, add a new field to the object with the name, type, size and scale info.
                    string name = fieldElement.GetAttribute("Name");
                    AlteryxRecordInfoNet.FieldType type = AlteryxRecordInfoNet.FieldType.E_FT_String;
                    Enum.TryParse<AlteryxRecordInfoNet.FieldType>(fieldElement.GetAttribute("Type"), out type);
                    int size = 0;
                    int.TryParse(fieldElement.GetAttribute("Size"), out size);
                    int scale = 0;
                    int.TryParse(fieldElement.GetAttribute("Scale"), out scale);
                    
                    xmlConfig.AddField(name, type, size, scale);
                }

                return xmlConfig;
            }

            return null;
        }
    }

    // The XmlInputField class encapsulates the name, type, size and scale of a field.
    public class XmlInputField
    {
        public string Name { get; set; }
        public AlteryxRecordInfoNet.FieldType FieldType { get; set; }
        public int Size { get; set; }
        public int Scale { get; set; }
    }
}
