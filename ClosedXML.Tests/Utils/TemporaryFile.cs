using System;
using System.IO;

namespace ClosedXML.Tests.Utils
{
    internal class TemporaryFile : IDisposable
    {
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
            if (!Preserve)
                File.Delete(Path);
        }

        public override string ToString()
        {
            return this.Path;
        }
    }
}
