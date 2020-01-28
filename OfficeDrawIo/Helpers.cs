using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace OfficeDrawIo
{
    public static class Helpers
    {
        public static string LoadStringResource(Assembly asm, string name)
        {
            var resourceName = asm.GetName().Name + "." + name;
            using (var s = asm.GetManifestResourceStream(resourceName))
            {
                if (s == null)
                    throw new Exception("Could not load resource: " + resourceName);
                using (StreamReader reader = new StreamReader(s))
                    return reader.ReadToEnd();
            }
        }

        public static byte[] LoadBinaryResource(Assembly asm, string name)
        {
            var resourceName = asm.GetName().Name + "." + name;
            using (var s = asm.GetManifestResourceStream(resourceName))
            {
                if (s == null)
                    throw new Exception("Could not load resource: " + resourceName);

                return s.ReadAllBytes();
            }
        }

        public static Stream GetResourceStream(Assembly asm, string name)
        {
            var resourceName = asm.GetName().Name + "." + name;
            var s = asm.GetManifestResourceStream(resourceName);

            if (s == null)
                throw new Exception("Could not load resource: " + resourceName);

            return s;
        }

        public static string GetVersionString(Assembly asm)
        {
            var ver = asm.GetName().Version; 
            
            var res = $"{ver.Major}.{ver.Minor}.{ver.Build}";
            if (ver.Revision != 0)
                res += $".{ver.Revision}";

            return res;
        }
    }
}
