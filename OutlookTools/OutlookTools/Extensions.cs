using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OutlookTools
{
    // XML extension methods used to handle null values from the tool's configuration file and return the appropriate value type.
    public static class Extensions
    {
        public static bool InnerBoolean(this XmlElement xml)
        {
            bool value = false;

            if (xml != null)
            {
                Boolean.TryParse(xml.InnerText, out value);

                return value;
            }

            return value;
        }

        public static DateTime InnerDateTime(this XmlElement xml)
        {
            DateTime value = DateTime.Now;

            if (xml != null)
            {
                DateTime.TryParse(xml.InnerText, out value);

                return value;
            }

            return value;
        }

        public static int InnerInt(this XmlElement xml)
        {
            Int16 value = 0;

            if (xml != null)
            {
                Int16.TryParse(xml.InnerText, out value);

                return value;
            }

            return value;
        }

        public static int InnerInt<T>(this XmlElement xml) where T : IComparable, IFormattable, IConvertible
        {
            Int16 value = 0;

            if (typeof(T) == typeof(ExchangeVersion))
                value = 5;

            if (typeof(T) == typeof(WellKnownFolderName))
                value = 4;

            if (xml != null)
            {
                Int16.TryParse(xml.InnerText, out value);

                return value;
            }

            return value;
        }

        public static string InnerString(this XmlElement xml)
        {
            return xml == null ? String.Empty : xml.InnerText;
        }
    }
}
