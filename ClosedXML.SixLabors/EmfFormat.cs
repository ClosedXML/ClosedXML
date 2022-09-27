using SixLabors.ImageSharp.Formats;
using System.Collections.Generic;

namespace ClosedXML.Graphics
{
    internal class EmfFormat : IImageFormat<EmfMetadata>
    {
        public static EmfFormat Instance { get; } = new();

        public string Name => "EMF";

        public string DefaultMimeType => "image/emf";

        public IEnumerable<string> MimeTypes { get; } = new[] { "image/emf" };

        public IEnumerable<string> FileExtensions { get; } = new[] { "emf" };

        public EmfMetadata CreateDefaultFormatMetadata() => new();
    }
}
