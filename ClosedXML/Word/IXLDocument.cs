using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

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

        /// <summary>
        /// The word document
        /// </summary>
        WordprocessingDocument Document { get; set; }
        MainDocumentPart MainDocumentPart { get; set; }
        Document DocumentPart { get; set; }
        Body BodyPart { get; set; }

        string FileName { get; set; }

        void AddTextBlock( IXLTextBlock textBlock );

        void AddTextBlock( string text );

        void AddBlock( IXLBlock block );

        string GenerateBlockIds( );

        /// <summary>
        /// Gets the block with the given id
        /// </summary>
        /// <param name="blockId"></param>
        /// <returns></returns>
        /// /// <exception cref="InvalidOperationException">Thrown when there is no Block with the given id</exception>
        IXLBlock Block(string blockId);

        /// <summary>
        /// All the blocks in the document
        /// </summary>
        /// <returns></returns>
        IXLBlocks Blocks();
    }
}
