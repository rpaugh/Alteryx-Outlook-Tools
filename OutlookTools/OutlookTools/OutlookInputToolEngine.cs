using System;
using System.Linq;
using System.Xml;
using Microsoft.Exchange.WebServices.Data;
using System.Reflection;
using System.Collections.Generic;

namespace OutlookTools
{
    // Implements the plugin's runtime engine.
    // The class must implement INetPlugin which Alteryx uses to communicate with
    // the plugin.  INetPlugin exposes six methods that the class must implement:
    // PI_Init - called by Alteryx to initlialize the plugin with configuration info.
    // PI_AddIncomingConnection - called by Alteryx to notify the plugin of a new incoming connection.
    // PI_AddOutgoingConnection - called by Alteryx to notify the plugin of a new ougoing connection.
    // PI_PushAllRecords - called by Alteryx for an Input tool to request all of the records.
    // PI_Close - called by Alteryx when the plugin is being closed.
    // ShowDebugMessages - called by Alteryx when an error has occured to determine
    //                     whether or not to show debug details.
    public class OutlookInputToolEngine : AlteryxRecordInfoNet.INetPlugin
    {
        #region Class-level data

        // Manages our output connections.
        // Message stream:
        AlteryxRecordInfoNet.PluginOutputConnectionHelper m_outputHelper;
        // Attachment stream:
        AlteryxRecordInfoNet.PluginOutputConnectionHelper m_attachmentOutputHelper;

        // Our tool's unique id.
        int m_nToolId;

        // An interface for sending messages to Alteryx.
        AlteryxRecordInfoNet.EngineInterface m_engineInterface;

        // The xml configuration for our tool - created by our GUI.
        XmlElement m_xmlProperties;

        #endregion

        #region INetPlugin methods

        // Called by Alteryx to initlialize the plugin with configuration info.
        public void PI_Init(int nToolID, AlteryxRecordInfoNet.EngineInterface engineInterface, System.Xml.XmlElement pXmlProperties)
        {
            // Save the incoming information.
            m_nToolId = nToolID;
            m_engineInterface = engineInterface;
            m_xmlProperties = pXmlProperties;

            // Create our PluginOutputConnectionHelper that will be used to manage
            // any outgoing connections.
            m_outputHelper = new AlteryxRecordInfoNet.PluginOutputConnectionHelper(m_nToolId, m_engineInterface);
            m_attachmentOutputHelper = new AlteryxRecordInfoNet.PluginOutputConnectionHelper(m_nToolId, m_engineInterface);
        }

        // Called by Alteryx to notify the plugin of a new incoming connection.
        public AlteryxRecordInfoNet.IIncomingConnectionInterface PI_AddIncomingConnection(string pIncomingConnectionType, string pIncomingConnectionName)
        {
            // Since we are an input tool, we don't accept any incoming connections.
            // In our implementation of IPlugin (in XmlInputTool.cs) we specified no
            // incoming connections, so this method should never be called.
            throw new NotImplementedException("Input tools cannot accept incoming connections.");
        }

        // Called by Alteryx to notify the plugin of a new ougoing connection.
        public bool PI_AddOutgoingConnection(string pOutgoingConnectionName, AlteryxRecordInfoNet.OutgoingConnection outgoingConnection)
        {
            // Add the outgoing connection to our PluginOutputConnectionHelper so it can manage it.
            if (pOutgoingConnectionName == "Attachment Output")
                m_attachmentOutputHelper.AddOutgoingConnection(outgoingConnection);
            else
                m_outputHelper.AddOutgoingConnection(outgoingConnection);

            // Return true to indicate that we accepted the connection.
            return true;
        }

        // Called by Alteryx for an Input tool to request all of the records.
        public bool PI_PushAllRecords(long nRecordLimit)
        {
            // The nRecordLimit parameter specifies the maximum number of records that
            // should be provided, or -1 for unlimited.  If it is -1, set it to
            // long.MaxValue to make the processing easier later on.  Sometimes Alteryx
            // will call this function with a record limit of 0 just to get the output
            // record configuration.
            if (nRecordLimit < 0) nRecordLimit = long.MaxValue;

            m_engineInterface.OutputMessage(m_nToolId, AlteryxRecordInfoNet.MessageStatus.STATUS_Info, nRecordLimit.ToString());

            if (m_engineInterface.GetInitVar("UpdateOnly") == "True")
            {
                m_engineInterface.OutputMessage(m_nToolId, AlteryxRecordInfoNet.MessageStatus.STATUS_Complete, nRecordLimit.ToString());
            }
            else
            { 
                // Get the configuration section out of the properties xml that was passed
                // into PI_Init() and use it to determine the xml file, container element,
                // and field configuration for our tool.
                
                XmlElement configXml = m_xmlProperties.SelectSingleNode("Configuration") as XmlElement;
                XmlInputConfiguration xmlConfig = XmlInputConfiguration.LoadFromConfiguration(configXml);

                if (xmlConfig == null)
                {
                    throw new ApplicationException("Invalid configuration.  Ensure that the input file is set correctly.");
                }

                if (xmlConfig.Fields.Count == 0)
                {
                    throw new ApplicationException("There are no fields.  Make sure your container element is set properly.");
                }
                
                // Create a new RecordInfo object to describe our outgoing message records and
                // configure it based on the field information in our saved configuration.
                AlteryxRecordInfoNet.RecordInfo recordInfoOut = new AlteryxRecordInfoNet.RecordInfo();
                foreach (XmlInputField field in xmlConfig.Fields)
                {
                    // For each field in our configuration, add it to our RecordInfo object.
                    recordInfoOut.AddField(field.Name, field.FieldType, field.Size, field.Scale, "", "");
                }

                // Use the new RecordInfo object to initialize the PluginOutputConnectionHelper.
                // The PluginOutputConnectionHelper can't be used until this step is performed.
                m_outputHelper.Init(recordInfoOut, "Message Output", null, m_xmlProperties);

                // Create a Record object to hold the data for each outgoing message record.
                AlteryxRecordInfoNet.Record recordOut = recordInfoOut.CreateRecord();

                // Create a new RecordInfo object to describe our outgoing attachment records and
                // configure it based on the field information in our saved configuration.
                AlteryxRecordInfoNet.RecordInfo recordInfoOut_AttachmentPaths = new AlteryxRecordInfoNet.RecordInfo();

                // For each field in our configuration, add it to our RecordInfo object.
                recordInfoOut_AttachmentPaths.AddField("Id", AlteryxRecordInfoNet.FieldType.E_FT_String, 4000, 0, "", "");
                recordInfoOut_AttachmentPaths.AddField("AttachmentPath", AlteryxRecordInfoNet.FieldType.E_FT_String, 4000, 0, "", "");

                // Use the new RecordInfo object to initialize the PluginOutputConnectionHelper.
                // The PluginOutputConnectionHelper can't be used until this step is performed.
                m_attachmentOutputHelper.Init(recordInfoOut_AttachmentPaths, "Attachment Output", null, m_xmlProperties);

                // Create a Record object to hold the data for each outgoing attachment record.
                AlteryxRecordInfoNet.Record recordOut_AttachmentPaths = recordInfoOut_AttachmentPaths.CreateRecord();
                
                // Define the necessary PropertyDefinitionBase objects for each field in the XML configuration document.
                PropertyDefinitionBase[] propertyDefinitionBase = new PropertyDefinitionBase[xmlConfig.Fields.Count];

                for (int i = 0; i < xmlConfig.Fields.Count; i++)
                {
                    propertyDefinitionBase[i] = (PropertyDefinitionBase)typeof(ItemSchema).GetField(xmlConfig.Fields[i].Name).GetValue(null);
                }

                // Assign the configuration settings and field list to the OutlookEmail object.
                OutlookEmail email = new OutlookEmail() { UserName = xmlConfig.UserName, Password = xmlConfig.Password, UseManualServiceURL = xmlConfig.UseManualServiceURL, ServiceURL = xmlConfig.ServiceURL, Folder = (WellKnownFolderName)xmlConfig.Folder, AttachmentPath = xmlConfig.AttachmentPath, QueryString = xmlConfig.QueryString, IncludeSubFolders = xmlConfig.IncludeSubFolders, SubFolderName = xmlConfig.SubFolderName, SkipRootFolder = xmlConfig.SkipRootFolder, UseUniqueFileName = xmlConfig.UseUniqueFileName };
                email.Fields = new PropertySet(propertyDefinitionBase);

                // Get the list of messages (this includes attachments if the Attachment field was selected for output).
                List<Item> messages = email.GetMessages(nRecordLimit);

                // We will need to send status updates to Alteryx at regular intervals during
                // this process, so we'll do that based on an elapsed time.
                DateTime last = DateTime.Now;

                // We need to keep track of how many records we have processed.
                long nRecords = 0;

                // Process the data in each message object.
                foreach (var message in messages)
                {
                    // If we've exceeded the record limit, stop processing.
                    if (nRecords >= nRecordLimit) break;

                    // Reset our output record so we can reuse it.  This is better than
                    // creating a new Record in each iteration as these objects can get large.
                    recordOut.Reset();

                    // For each field, load the data from the corresponding element into the Record.
                    foreach (XmlInputField field in xmlConfig.Fields)
                    {
                        // Get the FieldBase from the RecordInfo for the field.
                        AlteryxRecordInfoNet.FieldBase fieldBase = recordInfoOut.GetFieldByName(field.Name, false);
                        if (fieldBase != null)
                        {
                            // Find the element within the container element that has the same name as the field.
                            var value = Convert.ToString(message.GetType().GetProperty(field.Name).GetValue(message, null));

                            if (value == null)
                            {
                                // If the field element doesn't exist, set the output field's value to null.
                                fieldBase.SetNull(recordOut);
                            }
                            else
                            {
                                // Otherwise, set the output field's value based on the element's inner text.
                                fieldBase.SetFromString(recordOut, value);
                            }
                        }
                    }

                    // Return message attachments if applicable.
                    if (!string.IsNullOrWhiteSpace(xmlConfig.AttachmentPath))
                    {
                        foreach (var attachment in email.Attachments.Where(x => x.Id == message.Id)/*message.Attachments*/)
                        {
                            //if (attachment is FileAttachment)
                            //{
                            //    FileAttachment fileAttachment = (FileAttachment)attachment;

                                // Reset our output record so we can reuse it.  This is better than
                                // creating a new Record in each iteration as these objects can get large.
                                recordOut_AttachmentPaths.Reset();

                                // Get the FieldBase from the RecordInfo for the ID field.
                                AlteryxRecordInfoNet.FieldBase fieldBase_ID = recordInfoOut_AttachmentPaths.GetFieldByName("ID", false);
                                if (fieldBase_ID != null) { fieldBase_ID.SetFromString(recordOut_AttachmentPaths, message.Id.ToString()); }

                                // Get the FieldBase from the RecordInfo for the AttachmentPath field.
                                AlteryxRecordInfoNet.FieldBase fieldBase_AttachmentPath = recordInfoOut_AttachmentPaths.GetFieldByName("AttachmentPath", false);
                                if (fieldBase_AttachmentPath != null) { fieldBase_AttachmentPath.SetFromString(recordOut_AttachmentPaths, attachment.AttachmentPath/*String.Format("{0}\\{1}", xmlConfig.AttachmentPath, fileAttachment.Name)*/); }

                                // Send the record to the downstream tools through the PluginOutputConnectionHelper.
                                m_attachmentOutputHelper.PushRecord(recordOut_AttachmentPaths.GetRecord());
                            //}
                        }
                    }

                    // Send the record to the downstream tools through the PluginOutputConnectionHelper.
                    m_outputHelper.PushRecord(recordOut.GetRecord());

                    // If at least 1 second has passed since we started or our last update, update progress.
                    if (DateTime.Now.Subtract(last).TotalSeconds >= 1)
                    {
                        // Determine the percent complete:  (Records Processed) / Min(RecordLimit, # of Container Elements)
                        double percentComplete = (double)nRecords / Math.Min(nRecordLimit, messages.Count);

                        // Output the progress
                        if (m_engineInterface.OutputToolProgress(m_nToolId, percentComplete) != 0)
                        {
                            // If this returns anything but 0, then the user has canceled the operation.
                            throw new AlteryxRecordInfoNet.UserCanceledException();
                        }

                        // Have the PluginOutputConnectionHelper ask the downstream tools to update their progress.
                        m_outputHelper.UpdateProgress(percentComplete);

                        // Reset the timer.
                        last = DateTime.Now;
                    }

                    // Have the PluginOutputConnectionHelper update the record count display in Alteryx.
                    // The PluginOutputConnectionHelper will actually only do this if enough time and data has elapsed,
                    // so it's ok to call this in every iteration.
                    m_outputHelper.OutputRecordCount(false);

                    // Update our record count.
                    nRecords++;
                }

                // If we weren't just getting the output config, send a status message to Alterys that
                // will display the number of records that we output.
                if (nRecordLimit > 0)
                {
                    m_engineInterface.OutputMessage(m_nToolId, AlteryxRecordInfoNet.MessageStatus.STATUS_Info, nRecords.ToString() + " records read from " + xmlConfig.UserName);
                }


                // Tell Alteryx that we are done sending data so that it can close the connections to our plugin.
                m_engineInterface.OutputMessage(m_nToolId, AlteryxRecordInfoNet.MessageStatus.STATUS_Complete, "");
            }

            // Close our ouput connections.
            m_outputHelper.Close();
            m_attachmentOutputHelper.Close();

            // Return true to indicate that we successfully processed our data.
            return true;
        }

        // Called by Alteryx when the plugin is being closed.
        public void PI_Close(bool bHasErrors)
        {
            // Use this method to clean up any resources that the plugin may have used.
        }

        // Called by Alteryx when an error has occured to determine whether or not to show debug details.
        public bool ShowDebugMessages()
        {
            // Return true to help us debug our tool.
            // This should be set to false for general distribution.
            return true;
        }

        #endregion
    }
}
