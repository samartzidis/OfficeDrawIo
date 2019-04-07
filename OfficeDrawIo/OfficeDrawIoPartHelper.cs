using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    class OfficeDrawIoPartHelper
    {
        private Microsoft.Office.Interop.Word.Document _doc;

        public OfficeDrawIoPartHelper(Microsoft.Office.Interop.Word.Document doc)
        {
            _doc = doc;
        }

        public void UpdateDrawIoDataPart(string id, byte[] data)
        {
            var xmlPart = _doc.CustomXMLParts.SelectByID(id);
            if (xmlPart == null)
                return;

            var base64 = Convert.ToBase64String(data);
            xmlPart.DocumentElement.FirstChild.NodeValue = base64;
        }

        public bool ExistsDrawIoDataPart(string id)
        {
            var xmlPart = _doc.CustomXMLParts.SelectByID(id);

            return xmlPart != null;
        }

        public Microsoft.Office.Core.CustomXMLPart AddDrawIoDataPart(byte[] data)
        {
            var base64 = System.Convert.ToBase64String(data);

            var partTemplate = Helpers.LoadStringResource("Resources.OfficeDrawIoPartTemplate.xml");
            var partPayload = string.Format(partTemplate, base64);

            var part = _doc.CustomXMLParts.Add(partPayload);
            return part;
        }

        public byte[] GetDrawIoDataPart(string id)
        {
            var part = _doc.CustomXMLParts.SelectByID(id);
            if (part == null)
                return null;

            byte[] data = null;
            try
            {
                var partPayloadDoc = new System.Xml.XmlDocument();
                partPayloadDoc.LoadXml(part.XML);
                data = System.Convert.FromBase64String(partPayloadDoc.InnerText);
            }
            catch (Exception m)
            {
                Trace.TraceError($"Failed to decode part with id = {id}, details: {m.ToString()}");
            }

            return data;
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

        public static bool IsOfficeDrawIoPart(Microsoft.Office.Core.CustomXMLPart part)
        {
            return part.XML != null && part.XML.TrimStart().StartsWith("<OfficeDrawIo");
        }
    }
}
