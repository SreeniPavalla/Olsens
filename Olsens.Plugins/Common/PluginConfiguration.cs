using System;
using System.Xml;

namespace Olsens.Plugins.Common
{
     class PluginConfiguration
    {
        public static XmlNode GetNode(XmlDocument doc, string key)
        {
            if (doc != null && !string.IsNullOrEmpty(key))
            {
                XmlNode node = doc.SelectSingleNode(key);
                return node;
            }
            return null;
        }

        public static string GetInnerTextOfNode(XmlDocument doc, string key)
        {
            if (doc != null && !string.IsNullOrEmpty(key))
            {
                XmlNode node = doc.SelectSingleNode(key);
                if (node != null && !string.IsNullOrEmpty(node.InnerText))
                {
                    return node.InnerText;
                }
            }
            return string.Empty;
        }

        private static string GetValueNode(XmlDocument doc, string key)
        {
            XmlNode node = doc.SelectSingleNode(String.Format("configuration/settings/setting[@name='{0}']", key));
            if (node != null)
            {
                if (node.Attributes != null) return node.Attributes["value"].Value;
            }
            return string.Empty;
        }

        public static Guid GetConfigDataGuid(XmlDocument doc, string label)
        {
            string tempString = GetValueNode(doc, label);
            if (tempString != string.Empty)
            {
                return new Guid(tempString);
            }
            return Guid.Empty;
        }

        public static bool GetConfigDataBool(XmlDocument doc, string label)
        {
            bool retVar;
            if (bool.TryParse(GetValueNode(doc, label), out retVar))
            {
                return retVar;
            }
            else
            {
                return false;
            }
        }

        public static int GetConfigDataInt(XmlDocument doc, string label)
        {
            int retVar;
            if (int.TryParse(GetValueNode(doc, label), out retVar))
            {
                return retVar;
            }
            else
            {
                return -1;
            }
        }

        public static string GetConfigDataString(XmlDocument doc, string label)
        {
            return GetValueNode(doc, label);
        }
    }
}
