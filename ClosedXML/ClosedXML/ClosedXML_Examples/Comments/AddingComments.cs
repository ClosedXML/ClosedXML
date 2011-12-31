using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;

namespace ClosedXML_Examples 
{
    public class AddingComments : IXLExample 
    {

        public void Create(string filePath) 
        {
            var wb = new XLWorkbook();
            AddMiscComments(wb);
            AddVisibilityComments(wb);
            AddStyleAlignment(wb);
            wb.SaveAs(filePath);
        }

        private void AddStyleAlignment(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Style Alignment");

            // Automagically adjust the size of the comment to fit the contents
            ws.Cell("A1").Comment.Style.Alignment.SetAutomaticSize();
            ws.Cell("A1").Comment.AddText("Things are pretty tight around here");


            // Set all comments to visible
            ws.CellsUsed(c => c.HasComment).ForEach(c => c.Comment.SetVisible());
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


            ws.CellsUsed(c => !c.Address.ToStringRelative().Equals("A1") && c.HasComment).ForEach(c => c.Comment.SetVisible());
        }

        private static void AddVisibilityComments(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Visibility");

            // By default comments are hidden
            ws.Cell("A1").SetValue("I have a comment").Comment.AddText("Hidden");
            
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
            // ws.CellsUsed(c => c.HasComment).ForEach(c => c.Comment.SetVisible());

            ws.Columns().AdjustToContents();
        }
    }
}
