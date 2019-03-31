using System;
using System.Windows.Forms;

namespace OfficeDrawIo
{
    public class ScopedCursor : IDisposable
    {
        protected Cursor oldCursor = null;

        protected ScopedCursor()
        {

        }

        public ScopedCursor(Cursor c)
        {
            oldCursor = Cursor.Current;
            Cursor.Current = c;
        }

        ~ScopedCursor()
        {
            if (oldCursor != null)
                Cursor.Current = oldCursor;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (oldCursor != null)
                Cursor.Current = oldCursor;
        }

        #endregion
    }
}
