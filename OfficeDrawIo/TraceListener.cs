using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    public class TraceListener : DefaultTraceListener
    {
        public override void Write(string message)
        {
            base.Write($"OfficeDrawIo: {message}");
        }

        public override void WriteLine(string message)
        {
            base.WriteLine($"OfficeDrawIo: {message}");
        }
    }
}
