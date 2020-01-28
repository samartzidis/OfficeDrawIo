using System;

namespace OfficeDrawIo
{
    public class ScopedLambda : IDisposable
    {
        private bool _disposed;
        private readonly Action _executeOnDispose;

        public ScopedLambda(Action executeOnConstruct, Action executeOnDispose)
        {
            executeOnConstruct?.Invoke();

            _executeOnDispose = executeOnDispose;
        }

        public ScopedLambda(Action executeOnDispose)
        {
            _executeOnDispose = executeOnDispose;
        }

        #region IDisposable Members
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (_disposed)
                return;
            if (disposing && _executeOnDispose != null)
                _executeOnDispose();

            _disposed = true;
        }
        #endregion
    }
}
