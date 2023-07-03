#nullable disable

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using ClosedXML.Extensions;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class CommentPartWriter
    {
        internal static void GenerateWorksheetCommentsPartContent(WorksheetCommentsPart worksheetCommentsPart,
            XLWorksheet xlWorksheet)
        {
            var settings = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = XLHelper.NoBomUTF8
            };
            var partStream = worksheetCommentsPart.GetStream(FileMode.Create);
            using var xml = XmlWriter.Create(partStream, settings);

            var commentCells = new List<XLCell>();
            var authorsDict = new Dictionary<String, Int32>();
            xml.WriteStartElement("x", "comments", Main2006SsNs);
            foreach (var c in xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
            {
                var authorName = c.GetComment().Author;

                if (!authorsDict.TryGetValue(authorName, out var authorId))
                {
                    authorId = authorsDict.Count;
                    authorsDict.Add(authorName, authorId);
                }

                commentCells.Add(c);
            }

            xml.WriteStartElement("authors", Main2006SsNs);
            foreach (var author in authorsDict)
                xml.WriteElementString("author", Main2006SsNs, author.Key);

            xml.WriteEndElement(); // authors

            var refBuffer = new char[10];
            xml.WriteStartElement("commentList", Main2006SsNs);
            foreach (var commentCell in commentCells)
            {
                var comment = commentCell.GetComment();
                xml.WriteStartElement("comment", Main2006SsNs);

                var refLen = commentCell.SheetPoint.Format(refBuffer);
                xml.WriteStartAttribute("ref");
                xml.WriteRaw(refBuffer, 0, refLen);
                xml.WriteEndAttribute(); // ref

                var authorId = authorsDict[comment.Author];
                xml.WriteAttribute("authorId", authorId);

                // Excel specifies @guid is optional if the workbook is not shared
                // Excel ignores the shapeId attribute.

                xml.WriteStartElement("text", Main2006SsNs);
                var richText = XLImmutableRichText.Create(comment);
                foreach (var run in richText.Runs)
                    TextSerializer.WriteRun(xml, richText, run);

                xml.WriteEndElement(); // text
                xml.WriteEndElement(); // comment
            }

            xml.WriteEndElement(); // commentList
            xml.WriteEndElement(); // comments

            xml.Close();
        }
    }
}
