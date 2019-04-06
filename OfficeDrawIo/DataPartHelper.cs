using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    class DrawIoDataPartHelper
    {
        private Microsoft.Office.Interop.Word.Document _doc;

        public DrawIoDataPartHelper(Microsoft.Office.Interop.Word.Document doc)
        {
            _doc = doc;
        }

        public bool UpdateDrawIoDataPart(string id, byte[] data)
        {
            var xmlPart = _doc.CustomXMLParts.SelectByID(id);
            if (xmlPart == null)
                return false;

            var base64 = System.Convert.ToBase64String(data);
            xmlPart.DocumentElement.FirstChild.NodeValue = base64;

            return true;
        }

        public bool ExistsDrawIoDataPart(string id)
        {
            var xmlPart = _doc.CustomXMLParts.SelectByID(id);
            return xmlPart != null;
        }

        public Microsoft.Office.Core.CustomXMLPart AddDrawIoDataPart(byte[] data)
        {

            var base64 = System.Convert.ToBase64String(data);
            var xmlPart = _doc.CustomXMLParts.Add($"<OfficeDrawIo v=\"1\">{base64}</OfficeDrawIo>");
            return xmlPart;
        }

        public byte[] GetDrawIoDataPart(string id)
        {
            var xmlPart = _doc.CustomXMLParts.SelectByID(id);
            if (xmlPart == null)
                return null;

            var xdoc = new System.Xml.XmlDocument();
            xdoc.LoadXml(xmlPart.XML);

            var bytes = System.Convert.FromBase64String(xdoc.InnerText);
            return bytes;
        }

        public void DeleteDrawIoDataPart(string id)
        {
            foreach (Microsoft.Office.Core.CustomXMLPart part in _doc.CustomXMLParts)
            {
                if (part.Id == id)
                {
                    part.Delete();
                    break;
                }
            }
        }
    }
}
