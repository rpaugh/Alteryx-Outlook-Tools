using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookTools
{
    public class Enum<T> where T : struct, IConvertible
    {
        public static int Count
        {
            get
            {
                if (!typeof(T).IsEnum)
                    throw new ArgumentException("T must be an enumerated type");

                return Enum.GetNames(typeof(T)).Length;
            }
        }

        public static List<KeyValuePair<int, string>> ToList()
        {
            if (!typeof(T).IsEnum)
                throw new ArgumentException("T must be an enumerated type");

            List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();

            var items = Enum.GetValues(typeof(T)).Cast<T>().Select(x => new KeyValuePair<int, string>(Convert.ToInt32(x), x.ToString()));

            foreach (var item in items)
            {
                list.Add(new KeyValuePair<int, string>(item.Key, item.Value));
            }

            return list;
        }
    }

    public class Member<T> where T : class
    {
        public static List<KeyValuePair<string, string>> ToList()
        {
            if (!typeof(T).IsClass)
                throw new ArgumentException("T must be a class type");

            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();

            var members = typeof(T).GetMembers(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.FlattenHierarchy);

            foreach (var member in members)
            {
                list.Add(new KeyValuePair<string, string>(member.Name, member.Name));
            }

            return list;
        }
        public static List<KeyValuePair<string, string>> ToList(ExchangeVersion exchangeVersion)
        {
            if (!typeof(T).IsClass)
                throw new ArgumentException("T must be a class type");

            List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>();

            var members = typeof(T).GetMembers(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.FlattenHierarchy);

            foreach (var member in members)
            {
                // Exclude fields that cause errors with value retrieval via reflection of the PropertyDefinitionBase
                if (member.Name != "EntityExtractionResult" && member.Name != "Equals" && member.Name != "ExtendedProperties" && member.Name != "IconIndex" && member.Name != "InternetMessageHeaders" && member.Name != "IsReminderSet" && member.Name != "ReferenceEquals" && member.Name != "ReminderDueBy" && member.Name != "ReminderMinutesBeforeStart")
                {
                    if (exchangeVersion >= ExchangeVersion.Exchange2013)
                    {
                        list.Add(new KeyValuePair<string, string>(member.Name, member.Name));
                    }
                    else
                    {
                        if (member.Name != "ArchiveTag" && member.Name != "Flag" && member.Name != "IconIndex" && member.Name != "InstanceKey" && member.Name != "NormalizedBody" && member.Name != "PolicyTag" && member.Name != "Preview" && member.Name != "RetentionDate" && member.Name != "TextBody")
                        {
                            list.Add(new KeyValuePair<string, string>(member.Name, member.Name));
                        }
                    }
                }
            }

            return list;
        }
    }
}
