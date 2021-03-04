using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;

namespace ClosedXML.Examples 
{
    class EditingComments : IXLExample 
    {

        public void Create(string filePath) {

            // Exercise(@"path/to/test/resources/comments");
          
        }

        public void Exercise(string basePath) 
        {

            // INCOMPLETE

            var book = new XLWorkbook(Path.Combine(basePath, "EditingComments.xlsx"));
            var sheet = book.Worksheet(1);

            // no change
            // A1

            // edit existing comment
            sheet.Cell("B3").GetComment().AddNewLine();
            sheet.Cell("B3").GetComment().AddSignature();
            sheet.Cell("B3").GetComment().AddText("more comment");

            // delete
            //sheet.Cell("C1").DeleteComment();

            // clear contents
            sheet.Cell("D3").Clear(XLClearOptions.Contents);

            // new basic
            sheet.Cell("E1").GetComment().AddText("non authored comment");

            // new with author
            sheet.Cell("F3").GetComment().AddSignature();
            sheet.Cell("F3").GetComment().AddText("comment from author");

            // TODO: merge with cells
            // TODO: resize with cells
            // TODO: visible

            book.SaveAs(Path.Combine(basePath, "EditingComments_modified.xlsx"));
        }
    }
}
