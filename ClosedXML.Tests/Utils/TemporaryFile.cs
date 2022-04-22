using System;
using System.IO;

namespace ClosedXML.Tests.Utils
{
    internal class TemporaryFile : IDisposable
    {
        private bool _disposed = false;

        internal TemporaryFile()
            : this(System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), "xlsx"))
        { }

        internal TemporaryFile(string path)
            : this(path, false)
        { }

        internal TemporaryFile(String path, bool preserve)
        {
            this.Path = path;
            this.Preserve = preserve;
        }

        public string Path { get; private set; }
        public bool Preserve { get; private set; }

        public void Dispose()
        {
            // Dispose of unmanaged resources.
            Dispose(true);
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                if (!Preserve)
                    File.Delete(Path);
            }

            _disposed = true;
        }

        public override string ToString()
        {
            return this.Path;
        }
    }
}
