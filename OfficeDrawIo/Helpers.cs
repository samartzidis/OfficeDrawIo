using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace OfficeDrawIo
{
    static class Helpers
    {
        public static string LoadStringResource(string name)
        {
            var execAsm = Assembly.GetExecutingAssembly();
            var resourceName = execAsm.GetName().Name + "." + name;
            using (var s = execAsm.GetManifestResourceStream(resourceName))
            {
                if (s == null)
                    throw new Exception("Could not load resource: " + resourceName);
                using (StreamReader reader = new StreamReader(s))
                    return reader.ReadToEnd();
            }
        }

        public static byte[] LoadBinaryResource(string name)
        {
            var execAsm = Assembly.GetExecutingAssembly();
            var resourceName = execAsm.GetName().Name + "." + name;
            using (var s = execAsm.GetManifestResourceStream(resourceName))
            {
                if (s == null)
                    throw new Exception("Could not load resource: " + resourceName);

                return s.ReadAllBytes();
            }
        }

        public static Stream GetResourceStream(string name)
        {
            var execAsm = Assembly.GetExecutingAssembly();
            var resourceName = execAsm.GetName().Name + "." + name;
            var s = execAsm.GetManifestResourceStream(resourceName);

            if (s == null)
                throw new Exception("Could not load resource: " + resourceName);

            return s;
        }

        public static string GetVersionString()
        {
            var asm = Assembly.GetExecutingAssembly();
            var ver = asm.GetName().Version; 
            
            var res = $"{ver.Major}.{ver.Minor}.{ver.Build}";
            if (ver.Revision != 0)
                res += $".{ver.Revision}";

            return res;
        }
    }

    public static class StreamExtensions
    {
        public static byte[] ReadAllBytes(this Stream stream)
        {
            using (MemoryStream ms = new MemoryStream((int)stream.Length))
            {
                byte[] buffer = new byte[4096];
                int bytesRead;
                while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) != 0)
                {
                    ms.Write(buffer, 0, bytesRead);
                }
                return ms.ToArray();
            }
        }

        public static void CopyTo(this Stream input, Stream output)
        {
            const int size = 4096;
            byte[] bytes = new byte[4096];
            int numBytes;
            while ((numBytes = input.Read(bytes, 0, size)) > 0)
                output.Write(bytes, 0, numBytes);
        }
    }
}
