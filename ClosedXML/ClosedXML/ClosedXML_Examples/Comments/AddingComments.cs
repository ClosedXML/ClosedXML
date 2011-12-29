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

            
            ws.CellsUsed(true, c => !c.Address.ToStringRelative().Equals("A1") && c.HasComment).ForEach(c => c.Comment.SetVisible());
            wb.SaveAs(filePath);
        }
    }
}
