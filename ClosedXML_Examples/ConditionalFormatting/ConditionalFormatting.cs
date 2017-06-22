using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples
{
    public class CFColorScaleLowMidHigh : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().ColorScale()
                .LowestValue(XLColor.Red)
                .Midpoint(XLCFContentType.Percent, "50", XLColor.Yellow)
                .HighestValue(XLColor.Green);

            workbook.SaveAs(filePath);
        }
    }

    public class CFColorScaleLowHigh : IXLExample
    {

        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().ColorScale()
                .Minimum(XLCFContentType.Number, "2", XLColor.Red)
                .Maximum(XLCFContentType.Percentile, "90", XLColor.Green);

            workbook.SaveAs(filePath);
        }
    }

    public class CFColorScaleMinimumMaximum : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().ColorScale()
                .LowestValue(XLColor.FromHtml("#FFFF7128"))
                .HighestValue(XLColor.FromHtml("#FFFFEF9C"));

            workbook.SaveAs(filePath);
        }
    }

    public class CFStartsWith : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenStartsWith("Hell")
                .Fill.SetBackgroundColor(XLColor.Red)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue)
                .Font.SetBold();

            workbook.SaveAs(filePath);
        }
    }

    public class CFEndsWith : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenEndsWith("ll")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFIsBlank : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("")
                .CellBelow().SetValue("")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenIsBlank()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotBlank : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("")
                .CellBelow().SetValue("")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenNotBlank()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFIsError : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetFormulaA1("1/0")
                .CellBelow().SetFormulaA1("1/0")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenIsError()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotError : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetFormulaA1("1/0")
                .CellBelow().SetFormulaA1("1/0")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenNotError()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFContains : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenContains("Hell")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotContains : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenNotContains("Hell")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }


    public class CFEqualsString : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenEquals("Hell")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFEqualsNumber : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenEquals(2)
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotEqualsString : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenNotEquals("Hell")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotEqualsNumber : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenNotEquals(2)
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFGreaterThan : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenGreaterThan("2")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFEqualOrGreaterThan : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenEqualOrGreaterThan("2")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFLessThan : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenLessThan("2")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFEqualOrLessThan : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenEqualOrLessThan("2")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFBetween : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenBetween("2", "3")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFNotBetween : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenNotBetween("2", "3")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFUnique : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenIsUnique()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFDuplicate : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenIsDuplicate()
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFIsTrue : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenIsTrue("TRUE")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFTop : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenIsTop(2)
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFBottom : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().WhenIsBottom(10, XLTopBottomType.Percent)
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFDataBar : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().DataBar(XLColor.Red, true)
                .LowestValue()
                .Maximum(XLCFContentType.Percent, "100");

            workbook.SaveAs(filePath);
        }
    }

    public class CFIconSet : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().IconSet(XLIconSetStyle.ThreeTrafficLights2, true, true)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "2", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number);

            workbook.SaveAs(filePath);
        }
    }

    public class CFTwoConditions : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().IconSet(XLIconSetStyle.ThreeTrafficLights2, true, true)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "2", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number);

            ws.RangeUsed().AddConditionalFormat().WhenContains("1")
                .Fill.SetBackgroundColor(XLColor.Red);

            workbook.SaveAs(filePath);
        }
    }

    public class CFInsertRows : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.Cell(2,1).SetValue(1)
                .CellRight().SetValue(1)
                .CellRight().SetValue(2)
                .CellRight().SetValue(3);

            var range = ws.RangeUsed();
            range.AddConditionalFormat().WhenEquals("1").Font.SetBold();
            range.InsertRowsAbove(1);


            workbook.SaveAs(filePath);
        }
    }

    public class CFTest : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
              .CellBelow().SetValue(1)
              .CellBelow().SetValue(2)
              .CellBelow().SetValue(3)
              .CellBelow().SetValue(4);

            ws.RangeUsed().AddConditionalFormat().DataBar(XLColor.Red)
                .LowestValue()
                .HighestValue();

            workbook.SaveAs(filePath);
        }
    }

    public class CFMultipleConditions : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            using(var range = ws.Range("A1:A10"))
            {
                range.AddConditionalFormat().WhenEquals("3")
                    .Fill.SetBackgroundColor(XLColor.Blue);
                range.AddConditionalFormat().WhenEquals("2")
                    .Fill.SetBackgroundColor(XLColor.Green);
                range.AddConditionalFormat().WhenEquals("1")
                    .Fill.SetBackgroundColor(XLColor.Red);
            }

            workbook.SaveAs(filePath);
        }
    }
}
