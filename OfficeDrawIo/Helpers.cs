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
}
