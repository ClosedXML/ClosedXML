using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    internal class AutoFilterTester
    {
        private readonly Action<IXLFilterColumn> _setFilter;
        private readonly List<(XLCellValue Value, Action<IXLStyle> SetStyle, bool ExpectedVisibility)> _values = new();

        internal AutoFilterTester(Action<IXLFilterColumn> setFilter)
        {
            _setFilter = setFilter;
        }

        internal AutoFilterTester Add(XLCellValue value, bool shouldBeVisible)
        {
            return Add(value, static (IXLStyle _) => { }, shouldBeVisible);
        }

        internal AutoFilterTester Add(XLCellValue value, Action<IXLNumberFormat> setNumberFormat, bool shouldBeVisible)
        {
            _values.Add((value, s => setNumberFormat(s.NumberFormat), shouldBeVisible));
            return this;
        }

        internal AutoFilterTester Add(XLCellValue value, Action<IXLStyle> setStyle, bool shouldBeVisible)
        {
            _values.Add((value, setStyle, shouldBeVisible));
            return this;
        }

        internal AutoFilterTester AddTrue(params XLCellValue[] values)
        {
            foreach (var value in values)
                Add(value, true);

            return this;
        }

        internal AutoFilterTester AddFalse(params XLCellValue[] values)
        {
            foreach (var value in values)
                Add(value, false);

            return this;
        }

        internal void AssertVisibility()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Data";
            for (var i = 0; i < _values.Count; ++i)
            {
                var cell = ws.Cell(i + 2, 1);
                cell.Value = _values[i].Value;
                _values[i].SetStyle(cell.Style);
            }

            var autoFilter = ws.Range(1, 1, _values.Count + 1, 1).SetAutoFilter();
            _setFilter(autoFilter.Column(1));

            for (var i = 0; i < _values.Count; ++i)
            {
                var row = i + 2;
                var value = ws.Cell(row, 1).CachedValue;
                var formattedString = ((XLCell)ws.Cell(row, 1)).GetFormattedString(value);
                var actualVisible = !ws.Row(row).IsHidden;
                var expectedVisibility = _values[i].ExpectedVisibility;
                Assert.AreEqual(expectedVisibility, actualVisible, $"Visibility differs at index {i} for value {value} (formatted '{formattedString}')");
            }
        }
    }
}
