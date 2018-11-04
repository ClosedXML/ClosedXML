using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;
using MoreLinq;

namespace ClosedXML_Examples 
{
    public class AddingComments : IXLExample 
    {

        public void Create(string filePath)
        {
            var wb = new XLWorkbook {Author = "Manuel"};
            AddMiscComments(wb);
            AddVisibilityComments(wb);
            AddPosition(wb);
            AddSignatures(wb);
            AddStyleAlignment(wb);
            AddColorsAndLines(wb);
            AddMagins(wb);
            AddProperties(wb);
            AddProtection(wb);
            AddSize(wb);
            AddWeb(wb);

            wb.SaveAs(filePath);
        }

        private void AddWeb(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Web");
            ws.Cell("A1").Comment.Style.Web.AlternateText = "The alternate text in case you need it.";
        }

        private void AddSize(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Size");

            // Automatic size is a copy of the property comment.Style.Alignment.AutomaticSize
            // I created the duplicate because it makes more sense for it to be in Size
            // but Excel has it under the Alignment tab.
            ws.Cell("A2").Comment.AddText("Things are very tight around here.");
            ws.Cell("A2").Comment.Style.Size.SetAutomaticSize();

            ws.Cell("A4").Comment.AddText("Different size");
            ws.Cell("A4").Comment.Style
                .Size.SetHeight(30) // The height is set in the same units as row.Height
                .Size.SetWidth(30); // The width is set in the same units as row.Width

            // Set all comments to visible
            ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private void AddProtection(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Protection");

            ws.Cell("A1").Comment.Style
                .Protection.SetLocked(false)
                .Protection.SetLockText(false);
        }

        private void AddProperties(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Properties");

            ws.Cell("A1").Comment.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
            ws.Cell("A2").Comment.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
            ws.Cell("A3").Comment.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
        }

        private void AddMagins(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Margins");

            ws.Cell("A2").Comment
                .SetVisible()
                .AddText("Lorem ipsum dolor sit amet, adipiscing elit. ").AddNewLine()
                .AddText("Nunc elementum, sapien a ultrices, commodo nisl. ").AddNewLine()
                .AddText("Consequat erat lectus a nisi. Aliquam facilisis.");

            ws.Cell("A2").Comment.Style
                .Margins.SetAll(0.25)
                .Size.SetAutomaticSize();
        }

        private void AddColorsAndLines(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Colors and Lines");

            ws.Cell("A2").Comment
                .AddText("Now ")
                .AddText("THIS").SetBold().SetFontColor(XLColor.Red)
                .AddText(" is colorful!");
            ws.Cell("A2").Comment.Style
                .ColorsAndLines.SetFillColor(XLColor.RichCarmine)
                .ColorsAndLines.SetFillTransparency(0.25) // 25% opaque
                .ColorsAndLines.SetLineColor(XLColor.Blue)
                .ColorsAndLines.SetLineTransparency(0.75) // 75% opaque
                .ColorsAndLines.SetLineDash(XLDashStyle.LongDash)
                .ColorsAndLines.SetLineStyle(XLLineStyle.ThickBetweenThin)
                .ColorsAndLines.SetLineWeight(7.5);

            // Set all comments to visible
            ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private void AddStyleAlignment(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Alignment");

            // Automagically adjust the size of the comment to fit the contents
            ws.Cell("A1").Comment.Style.Alignment.SetAutomaticSize();
            ws.Cell("A1").Comment.AddText("Things are pretty tight around here");

            // Default values
            ws.Cell("A3").Comment
                .AddText("Default Alignments:").AddNewLine()
                .AddText("Vertical = Top").AddNewLine()
                .AddText("Horizontal = Left").AddNewLine()
                .AddText("Orientation = Left to Right");

            // Let's change the alignments
            ws.Cell("A8").Comment
                .AddText("Vertical = Bottom").AddNewLine()
                .AddText("Horizontal = Right");
            ws.Cell("A8").Comment.Style
                .Alignment.SetVertical(XLDrawingVerticalAlignment.Bottom)
                .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Right);

            // And now the orientation...
            ws.Cell("D3").Comment.AddText("Orientation = Bottom to Top");
            ws.Cell("D3").Comment.Style
                .Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop)
                .Alignment.SetAutomaticSize();

            ws.Cell("E3").Comment.AddText("Orientation = Top to Bottom");
            ws.Cell("E3").Comment.Style
                .Alignment.SetOrientation(XLDrawingTextOrientation.TopToBottom)
                .Alignment.SetAutomaticSize();

            ws.Cell("F3").Comment.AddText("Orientation = Vertical");
            ws.Cell("F3").Comment.Style
                .Alignment.SetOrientation(XLDrawingTextOrientation.Vertical)
                .Alignment.SetAutomaticSize();


            // Set all comments to visible
            ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private static void AddMiscComments(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Comments");

            ws.Cell("A1").SetValue("Hidden").Comment.AddText("Hidden");
            ws.Cell("A2").SetValue("Visible").Comment.AddText("Visible");
            ws.Cell("A3").SetValue("On Top").Comment.AddText("On Top");
            ws.Cell("A4").SetValue("Underneath").Comment.AddText("Underneath");
            ws.Cell("A4").Comment.Style.Alignment.SetVertical(XLDrawingVerticalAlignment.Bottom);
            ws.Cell("A3").Comment.SetZOrder(ws.Cell("A4").Comment.ZOrder + 1);

            ws.Cell("D9").Comment.AddText("Vertical");
            ws.Cell("D9").Comment.Style.Alignment.Orientation = XLDrawingTextOrientation.Vertical;
            ws.Cell("D9").Comment.Style.Size.SetAutomaticSize();

            ws.Cell("E9").Comment.AddText("Top to Bottom");
            ws.Cell("E9").Comment.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
            ws.Cell("E9").Comment.Style.Size.SetAutomaticSize();

            ws.Cell("F9").Comment.AddText("Bottom to Top");
            ws.Cell("F9").Comment.Style.Alignment.Orientation = XLDrawingTextOrientation.BottomToTop;
            ws.Cell("F9").Comment.Style.Size.SetAutomaticSize();

            ws.Cell("E1").Comment.Position.SetColumn(5);
            ws.Cell("E1").Comment.AddText("Start on Col E, on top border");
            ws.Cell("E1").Comment.Style.Size.SetWidth(10);
            var cE3 = ws.Cell("E3").Comment;
            cE3.AddText("Size and position");
            cE3.Position.SetColumn(5).SetRow(4).SetColumnOffset(7).SetRowOffset(10);
            cE3.Style.Size.SetHeight(25).Size.SetWidth(10);
            var cE7 = ws.Cell("E7").Comment;
            cE7.Position.SetColumn(6).SetRow(7).SetColumnOffset(0).SetRowOffset(0);
            cE7.Style.Size.SetHeight(ws.Row(7).Height).Size.SetWidth(ws.Column(6).Width);

            ws.Cell("G1").Comment.AddText("Automatic Size");
            ws.Cell("G1").Comment.Style.Alignment.SetAutomaticSize();
            var cG3 = ws.Cell("G3").Comment;
            cG3.SetAuthor("MDeLeon");
            cG3.AddSignature();
            cG3.AddText("This is a test of the emergency broadcast system.");
            cG3.AddNewLine();
            cG3.AddText("Do ");
            cG3.AddText("NOT").SetFontColor(XLColor.RadicalRed).SetUnderline().SetBold();
            cG3.AddText(" forget it.");
            cG3.Style
                .Size.SetWidth(25)
                .Size.SetHeight(100)
                .Alignment.SetDirection(XLDrawingTextDirection.LeftToRight)
                .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Distributed)
                .Alignment.SetVertical(XLDrawingVerticalAlignment.Center)
                .Alignment.SetOrientation(XLDrawingTextOrientation.LeftToRight)
                .ColorsAndLines.SetFillColor(XLColor.Cyan)
                .ColorsAndLines.SetFillTransparency(0.25)
                .ColorsAndLines.SetLineColor(XLColor.DarkBlue)
                .ColorsAndLines.SetLineTransparency(0.75)
                .ColorsAndLines.SetLineDash(XLDashStyle.DashDot)
                .ColorsAndLines.SetLineStyle(XLLineStyle.ThinThick)
                .ColorsAndLines.SetLineWeight(5)
                .Margins.SetAll(0.25)
                .Properties.SetPositioning(XLDrawingAnchor.MoveAndSizeWithCells)
                .Protection.SetLocked(false)
                .Protection.SetLockText(false)
                .Web.SetAlternateText("This won't be released to the web");

            ws.Cell("A9").Comment.SetAuthor("MDeLeon").AddSignature().AddText("Something");
            ws.Cell("A9").Comment.SetBold().SetFontColor(XLColor.DarkBlue);

            ws.CellsUsed(XLCellsUsedOptions.All, c => !c.Address.ToStringRelative().Equals("A1") && c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private static void AddVisibilityComments(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Visibility");

            // By default comments are hidden
            ws.Cell("A1").SetValue("I have a hidden comment").Comment.AddText("Hidden");
            
            // Set the comment as visible
            ws.Cell("A2").Comment.SetVisible().AddText("Visible");

            // The ZOrder on previous comments were 1 and 2 respectively
            // here we're explicit about the ZOrder
            ws.Cell("A3").Comment.SetZOrder(5).SetVisible().AddText("On Top");

            // We want this comment to appear underneath the one for A3
            // so we set the ZOrder to something lower
            ws.Cell("A4").Comment.SetZOrder(4).SetVisible().AddText("Underneath");
            ws.Cell("A4").Comment.Style.Alignment.SetVertical(XLDrawingVerticalAlignment.Bottom);
            
            // Alternatively you could set all comments to visible with the following line:
            // ws.CellsUsed(true, c => c.HasComment).ForEach(c => c.Comment.SetVisible());

            ws.Columns().AdjustToContents();
        }

        private void AddPosition(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Position");
            
            ws.Columns().Width = 10;

            ws.Cell("A1").Comment.AddText("This is an unusual place for a comment...");
            ws.Cell("A1").Comment.Position
                .SetColumn(3) // Starting from the third column
                .SetColumnOffset(5) // The comment will start in the middle of the third column
                .SetRow(5) // Starting from the fifth row
                .SetRowOffset(7.5); // The comment will start in the middle of the fifth row

            // Set all comments to visible
            ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private void AddSignatures(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Signatures");

            // By default the signature will be with the logged user
            // ws.Cell("A2").Comment.AddSignature().AddText("Hello World!");

            // You can override this by specifying the comment's author:
            ws.Cell("A2").Comment
                .SetAuthor("MDeLeon")
                .AddSignature()
                .AddText("Hello World!");
            

            // Set all comments to visible
            ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasComment).ForEach(c => c.Comment.SetVisible());
        }
    }
}
