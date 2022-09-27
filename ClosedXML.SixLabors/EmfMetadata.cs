using SixLabors.ImageSharp;

namespace ClosedXML.Graphics
{
    /// <summary>
    /// EMF specific metadata information.
    /// </summary>
    internal class EmfMetadata : IDeepCloneable
    {
        public EmfMetadata()
        {
        }

        private EmfMetadata(EmfMetadata other)
        {
            Frame = other.Frame;
        }

        /// <summary>
        /// A rectangular frame to draw in .01 millimeter units that surrounds the image strored in the metadafile.
        /// Borders are inclusive.
        /// </summary>
        public Rectangle Frame { get; set; }

        public IDeepCloneable DeepClone() => new EmfMetadata(this);
    }
}
