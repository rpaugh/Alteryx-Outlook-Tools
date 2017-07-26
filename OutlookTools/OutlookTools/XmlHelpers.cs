using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OutlookTools
{
    // This class contains some static methods that are useful for parsing XML files.
    public class XmlHelpers
    {
        public static XmlElement GetFirstChildByName(XmlNode parentNode, System.String sName, bool bThrowError)
        {
            // This function will return the first node with the specified name.
            // If the requested node cannot be found and bThrowError is true, 
            // an exception will be thrown, otherwise the function will simply
            // return null.
            XmlNode ret = parentNode.FirstChild;
            if (ret != null && (ret.NodeType != XmlNodeType.Element || ret.Name != sName))
                ret = GetNextSiblingByName(ret, sName);

            if (bThrowError && ret == null)
                throw new System.Exception("Missing XML node: " + sName);

            return (XmlElement)ret;
        }

        public static XmlElement GetNextSiblingByName(XmlNode siblingNode, System.String sName)
        {
            // This function will search for the next sibling node with the specified name.
            // If the node is not found, it will return null.
            if (sName == null || sName.Length == 0)
                sName = siblingNode.Name;

            XmlNode ret;
            for (ret = siblingNode.NextSibling; ret != null; ret = ret.NextSibling)
            {
                if (ret.NodeType == XmlNodeType.Element && ret.Name == sName)
                    break;
            }
            return (XmlElement)ret;
        }

        public static XmlElement GetOrCreateChildNode(XmlElement parentNode, string sName)
        {
            // This function returns the specified node if it already exists or
            // creates it if it does not.
            XmlElement ret = GetFirstChildByName(parentNode, sName, false);
            if (ret == null)
            {
                ret = parentNode.OwnerDocument.CreateElement(sName);
                parentNode.AppendChild(ret);
            }
            return ret;
        }

        public static XmlElement CreateElement(string sName, XmlElement parentNode)
        {
            // This function create the specified element.
            XmlElement ret = parentNode.OwnerDocument.CreateElement(sName);
            parentNode.InsertBefore(ret, null);
            return ret;
        }

        public static XmlAttribute CreateAttribute(string sName, XmlElement parentNode)
        {
            // This function create the specified attribute.
            XmlAttribute ret = parentNode.OwnerDocument.CreateAttribute(sName);
            parentNode.Attributes.Append(ret);
            return ret;
        }

        public static AlteryxRecordInfoNet.RecordInfo GetRecordInfo(XmlDocument xml, string elementName)
        {
            // This function will locate the first element with the specified elementName in the
            // provided XML document.  It will then create a RecordInfo object that contains a field
            // for each element within that element.

            // Create the RecordInfo object to be returned.
            AlteryxRecordInfoNet.RecordInfo recordInfo = new AlteryxRecordInfoNet.RecordInfo();

            // Get all elements with the specified element name.
            XmlNodeList nodes = xml.GetElementsByTagName(elementName);

            if (nodes.Count > 0)
            {
                // For each child node in the first found element...
                foreach (XmlNode child in nodes[0].ChildNodes)
                {
                    // ...add a field to the RecordInfo using the element's name aas the field name.
                    recordInfo.AddField(child.Name, AlteryxRecordInfoNet.FieldType.E_FT_String, 255, 0, "", "");
                }
            }

            return recordInfo;
        }
    }
}
