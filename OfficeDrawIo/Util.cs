using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    public static class Util
    {
        public const string OfficeDrawIoPayloadPrexix = "office-drawio";

        public static byte[] FileReadAllBytes(string path)
        {

            using (FileStream fileStream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite))
            {
                return fileStream.ReadAllBytes();
            }

        }

        public static string EncodePngFile(string path)
        {
            var pngBytes = FileReadAllBytes(path);
            var encoder = new BasE91();
            var pngEncoded = encoder.Encode(pngBytes).ToString();

            var encoded = $"{OfficeDrawIoPayloadPrexix}:{pngEncoded}";

            return encoded;
        }

        public static byte[] DecodePngFile(string encoded)
        {
            encoded = encoded.Substring($"{OfficeDrawIoPayloadPrexix}:".Length);
            var encoder = new BasE91();
            var pngBytes = encoder.Decode(encoded).ToArray();

            return pngBytes;
        }
    }
}
