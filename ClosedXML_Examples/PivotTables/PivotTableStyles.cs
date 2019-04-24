using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ClosedXML_Examples
{
    public class PivotTableStyles : IXLExample
    {
        private class Order
        {
            public string Company { get; }
            public string PaymentMethod { get; }
            public string OrderNo { get; }
            public DateTime ShipDate { get; }
            public int ItemsTotal { get; }
            public double TaxRate { get; }
            public double AmountPaid { get; }

            public Order(string company, string paymentMethod, string orderNo, DateTime shipDate, int itemsTotal, double taxRate, double amountPaid)
            {
                Company = company;
                PaymentMethod = paymentMethod;
                OrderNo = orderNo;
                ShipDate = shipDate;
                ItemsTotal = itemsTotal;
                TaxRate = taxRate;
                AmountPaid = amountPaid;
            }
        }

        public void Create(string filePath)
        {
            var wb = Create();
            using (wb)
            {
                wb.SaveAs(filePath);
            }
        }

        public static XLWorkbook Create()
        {
            var orders = new List<Order>
            {
                new Order("Davy Jones' Locker", "Check", "1004", DateTime.Parse("04.18.88"), 7885, 0, 7885),
                new Order("Davy Jones' Locker", "Credit", "1020", DateTime.Parse("06.25.88"), 9955, 0, 9955),
                new Order("Davy Jones' Locker", "Credit", "1120", DateTime.Parse("05.25.93"), 785, 0, 785),
                new Order("Davy Jones' Locker", "MC", "1095", DateTime.Parse("06.06.89"), 7532, 0, 0),
                new Order("Davy Jones' Locker", "MC", "1295", DateTime.Parse("01.06.95"), 17917, 0, 17917),
                new Order("Sight Diver", "Credit", "1003", DateTime.Parse("05.03.88"), 125, 4.5, 0),
                new Order("Sight Diver", "Credit", "1052", DateTime.Parse("01.07.89"), 16788, 0, 16788),
                new Order("Sight Diver", "Credit", "1055", DateTime.Parse("02.05.89"), 23406, 0, 23406),
                new Order("Sight Diver", "Credit", "1087", DateTime.Parse("05.21.89"), 14045, 0, 14045),
                new Order("Sight Diver", "Credit", "1152", DateTime.Parse("04.07.94"), 97699, 0, 97699),
                new Order("Sight Diver", "Credit", "1155", DateTime.Parse("05.05.94"), 13936, 0, 13936),
                new Order("Sight Diver", "Credit", "1163", DateTime.Parse("06.14.94"), 342, 0, 342),
                new Order("Sight Diver", "Credit", "1255", DateTime.Parse("12.09.94"), 64116, 0, 64116),
                new Order("Sight Diver", "MC", "1075", DateTime.Parse("04.22.89"), 8560, 0, 8560),
                new Order("Sight Diver", "MC", "1275", DateTime.Parse("12.22.94"), 16940, 0, 16940),
                new Order("Sight Diver", "Visa", "1067", DateTime.Parse("04.02.89"), 4495, 0, 4495)
            };

            var wb = new XLWorkbook();
            wb.Style.Font.SetFontSize(8);
            wb.Style.Font.FontName = "Arial";
            var ws = wb.Worksheets.Add("OrdersData");
            // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
            var table = ws.Cell(1, 1).InsertTable(orders, "OrdersData", true);
            ws.Columns().AdjustToContents();


            var ptSheet = wb.Worksheets.Add("pvtOrders");
            var pt = ptSheet.PivotTables.Add("pvtOrders", ptSheet.Cell("B3"), table.AsRange());
            pt.ClassicPivotTableLayout = true;
            pt.ShowExpandCollapseButtons = false;
            pt.MergeAndCenterWithLabels = true;
            pt.Theme = XLPivotTableTheme.None;


            // fields
            var companyField = pt.RowLabels.Add("Company").AddSubtotal(XLSubtotalFunction.Sum);
            var payMethodField = pt.RowLabels.Add("PaymentMethod", "Payment method").AddSubtotal(XLSubtotalFunction.Sum);
            var orderNoField = pt.RowLabels.Add("OrderNo", "OrderNo");

            var taxRateField = pt.ColumnLabels.Add("TaxRate", "Tax rate");

            var amountPaidField = pt.Values.Add("AmountPaid", "Amount paid").SetSummaryFormula(XLPivotSummary.Sum);
            var itemsTotalField = pt.Values.Add("ItemsTotal", "Items total").SetSummaryFormula(XLPivotSummary.Sum);

            // styles
            var purpleBlue = XLColor.FromArgb(102, 102, 153);
            var brightYellow = XLColor.FromArgb(255, 255, 153);
            var gray = XLColor.FromArgb(192, 192, 192);

            companyField.StyleFormats.Header.Style.Font.SetBold();
            companyField.StyleFormats.Header.Style.Font.SetFontSize(10);
            companyField.StyleFormats.Label.Style.Fill.BackgroundColor = brightYellow;
            companyField.StyleFormats.Subtotal.Label.Style.Fill.BackgroundColor = purpleBlue;
            companyField.StyleFormats.Subtotal.Label.Style.Font.SetFontSize(9);
            companyField.StyleFormats.Subtotal.Label.Style.Font.SetItalic(true);
            companyField.StyleFormats.Subtotal.Label.Style.Font.FontColor = XLColor.White;
            companyField.StyleFormats.Subtotal.AddValuesFormat().Style = companyField.StyleFormats.Subtotal.Label.Style;

            payMethodField.StyleFormats.Header.Style.Font.SetBold();
            payMethodField.StyleFormats.Header.Style.Font.SetFontSize(10);
            payMethodField.StyleFormats.Label.Style.Fill.BackgroundColor = purpleBlue;
            payMethodField.StyleFormats.Label.Style.Font.FontColor = XLColor.White;
            payMethodField.StyleFormats.Subtotal.Label.Style.Fill.BackgroundColor = gray;
            payMethodField.StyleFormats.Subtotal.AddValuesFormat().Style = payMethodField.StyleFormats.Subtotal.Label.Style;

            orderNoField.StyleFormats.Header.Style.Font.SetBold();
            orderNoField.StyleFormats.Header.Style.Font.SetFontSize(10);
            orderNoField.StyleFormats.Label.Style.Fill.BackgroundColor = purpleBlue;
            orderNoField.StyleFormats.Label.Style.Font.FontColor = XLColor.White;

            return wb;
        }
    }
}
