using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class PropertiesPipeBind : CmdletPipeBind<Dictionary<string, object>>
    {
        private string _xml;
        private Dictionary<string, object> _properties;

        public PropertiesPipeBind(Dictionary<string, object> instance) : base(instance)
        {
            _properties = instance;
        }
        public PropertiesPipeBind(Hashtable instance)
        {
            Hashtable hash = instance.Clone() as Hashtable;
            _properties = new Dictionary<string, object>();
            foreach (var key in hash.Keys)
            {
                string keyValue = key.ToString();
                _properties.Add(keyValue, hash[key]);
            }
        }
        public PropertiesPipeBind(XmlDocument instance)
        {
            _xml = instance.OuterXml;
        }

        public PropertiesPipeBind(string inputString)
        {
            XmlDocument xml = new XmlDocument();
            try
            {
                if (File.Exists(inputString))
                {
                    xml.Load(inputString);
                }
                else
                {
                    xml.LoadXml(inputString);
                }
            }
            catch
            {
                throw new Exception("The input string is not a valid XML file.");
            }
            _xml = xml.OuterXml;
        }



        protected override void Discover(Dictionary<string, object> instance)
        {
            _properties = instance;
        }

        public override Dictionary<string, object> Read()
        {
            if (_properties != null)
                return _properties;

            Dictionary<string, object> props = new Dictionary<string, object>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(_xml);
            if (xmlDoc.DocumentElement == null)
                return props;

            foreach (XmlElement propElement in xmlDoc.DocumentElement.ChildNodes)
            {
                props.Add(propElement.Attributes["Name"].Value, propElement.InnerText.Trim());
            }
            return props;
        }

    }

}
