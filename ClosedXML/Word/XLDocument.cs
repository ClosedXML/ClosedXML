﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace ClosedXML.Word
{
    public class XLDocument : IDisposable
    {
        private string file;

        public XLDocument( string file )
        {
            this.file = file;
            Load( this.file );
        }

        public XLDocument( )
        {
        }

        public void CreateNewWordDocument( )
        {
            string path = Path.Combine( Path.GetTempPath( ), "test.docx" );

            using ( FileStream fs = File.Create( path ) )
            {
                using ( WordprocessingDocument document = WordprocessingDocument.Create( fs, WordprocessingDocumentType.Document, true ) )
                {
                    MainDocumentPart main = document.AddMainDocumentPart( );
                    main.Document = new Document( );
                    Body body = new Body( );
                    main.Document.Body = body;
                    AddTextToBody( body, "ClosedXML Word Test" );
                    document.Close( );
                }
            }
        }

        private void Load( string file )
        {
            using ( WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open( file, true ) )
                LoadWordDocument( wordprocessingDocument );
        }

        private void LoadWordDocument( WordprocessingDocument document )
        {
            Body body = document.MainDocumentPart.Document.Body;
            AddTextToBody( body, "ClosedXML Word Test" );
            document.Close( );
        }

        private void AddTextToBody( Body body, string text )
        {
            Paragraph para = body.AppendChild( new Paragraph( ) );
            Run run = para.AppendChild( new Run( ) );
            run.AppendChild( new Text( text ) );
        }

        public void Dispose( )
        {
            throw new NotImplementedException( );
        }
    }
}