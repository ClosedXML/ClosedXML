using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel;

internal class XLHyperlinks : IXLHyperlinks, ISheetListener
{
    private readonly XLWorksheet _worksheet;
    private readonly Dictionary<XLSheetRange, XLHyperlink> _hyperlinks = new();

    private delegate (bool Success, XLSheetRange? RepositionedArea) RepositionFunc(XLSheetRange hyperlinkArea);

    internal XLHyperlinks(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    internal string WorksheetName => _worksheet.Name;

    #region ISheetListener

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange insertedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryInsertAreaAndShiftDown(insertedArea, out var newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange insertedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryInsertAreaAndShiftRight(insertedArea, out var newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryDeleteAreaAndShiftLeft(deletedArea, out var newHlArea);
            return (success, newHlArea);
        });
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedArea)
    {
        RepositionOnChange(sheet, hyperlinkArea =>
        {
            var success = hyperlinkArea.TryDeleteAreaAndShiftUp(deletedArea, out var newHlArea);
            return (success, newHlArea);
        });
    }

    private void RepositionOnChange(XLWorksheet sheet, RepositionFunc reposition)
    {
        if (sheet != _worksheet)
            return;

        var hyperlinkAreas = _hyperlinks.Keys.ToArray();
        foreach (var hyperlinkArea in hyperlinkAreas)
        {
            var (success, newHlArea) = reposition(hyperlinkArea);
            if (!success)
                continue; // Partial cover, don't move.

            if (hyperlinkArea == newHlArea)
                continue; // Nothing changed

            _hyperlinks.Remove(hyperlinkArea, out var hyperlink);
            if (newHlArea is not null)
                _hyperlinks.Add(newHlArea.Value, hyperlink);
        }
    }

    #endregion ISheetListener

    public IEnumerator<XLHyperlink> GetEnumerator()
    {
        return _hyperlinks.Values.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <inheritdoc />
    public bool Delete(XLHyperlink hyperlink)
    {
        if (!TryGet(hyperlink, out var range))
            return false;

        Clear(range.Value);
        ClearHyperlinkStyle(range.Value);
        return true;
    }

    /// <inheritdoc />
    public bool Delete(IXLAddress address)
    {
        var point = XLSheetPoint.FromAddress(address);
        if (Clear(point))
        {
            ClearHyperlinkStyle(point);
            return true;
        }

        return false;
    }

    /// <inheritdoc />
    public XLHyperlink Get(IXLAddress address)
    {
        return _hyperlinks[XLSheetPoint.FromAddress(address)];
    }

    /// <inheritdoc />
    public bool TryGet(IXLAddress address, out XLHyperlink hyperlink)
    {
        return _hyperlinks.TryGetValue(XLSheetPoint.FromAddress(address), out hyperlink);
    }

    /// <summary>
    /// Add a hyperlink. Doesn't modify style, unlike public API.
    /// </summary>
    internal void Add(XLSheetRange range, XLHyperlink hyperlink)
    {
        if (hyperlink.Container is not null && hyperlink.Container != this)
        {
            throw new InvalidOperationException("Hyperlink is attached to a different worksheet. Either remove it from the original worksheet or create a new hyperlink.");
        }

        _hyperlinks.Remove(range);
        _hyperlinks.Add(range, hyperlink);
        hyperlink.Container = this;
    }

    internal bool TryGet(XLSheetRange range, [NotNullWhen(true)] out XLHyperlink? hyperlink)
    {
        return _hyperlinks.TryGetValue(range, out hyperlink);
    }

    /// <summary>
    /// Remove a hyperlink. Doesn't modify style, unlike public API.
    /// </summary>
    internal bool Clear(XLSheetRange range)
    {
        if (_hyperlinks.Remove(range, out var hyperlink))
        {
            hyperlink.Container = null;
            return true;
        }

        return false;
    }

    internal XLCell? GetCell(XLHyperlink hyperlink)
    {
        if (!TryGet(hyperlink, out var range))
            return null;

        return new XLCell(_worksheet, range.Value.FirstPoint);
    }

    private bool TryGet(XLHyperlink hyperlink, [NotNullWhen(true)] out XLSheetRange? range)
    {
        var ranges = _hyperlinks
            .Where(x => x.Value == hyperlink)
            .Select(x => x.Key)
            .ToList();
        if (ranges.Count == 0)
        {
            range = null;
            return false;
        }

        range = ranges.Single();
        return true;
    }

    private void ClearHyperlinkStyle(XLSheetRange range)
    {
        var sheetColor = _worksheet.StyleValue.Font.FontColor;
        var sheetUnderline = _worksheet.StyleValue.Font.Underline;
        foreach (var point in range)
        {
            var cell = _worksheet.GetCell(point);
            if (cell is null)
                continue;

            if (cell.Style.Font.FontColor.Equals(XLColor.FromTheme(XLThemeColor.Hyperlink)))
                cell.Style.Font.FontColor = sheetColor;

            cell.Style.Font.Underline = sheetUnderline;
        }
    }
}
