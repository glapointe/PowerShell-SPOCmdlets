using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class XmlDocumentPipeBind : CmdletPipeBind<XmlDocument>
    {
        private string _xml;

        public XmlDocumentPipeBind(XmlDocument instance)
            : base(instance)
        {
            _xml = instance.OuterXml;
        }

        public XmlDocumentPipeBind(string inputString)
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
                throw new ArgumentException("The input string is not a valid XML file.");
            }
            _xml = xml.OuterXml;
        }



        protected override void Discover(XmlDocument instance)
        {
            _xml = instance.OuterXml;
        }

        public override XmlDocument Read()
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(_xml);
            return xml;
        }
    }
}
