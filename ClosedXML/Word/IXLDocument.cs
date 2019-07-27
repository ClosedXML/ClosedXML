using System;

namespace ClosedXML.Word
{
    public interface IXLDocument : IDisposable
    {
        /// <summary>
        /// Saves the document
        /// </summary>
        void Save( );

        /// <summary>
        /// Saves the document to a file
        /// </summary>
        /// <param name="file"></param>
        void SaveAs( string file );
    }
}
