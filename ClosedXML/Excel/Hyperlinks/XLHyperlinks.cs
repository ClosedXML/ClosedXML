using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ClosedXML.Utils;

namespace ClosedXML.Excel;

internal class XLHyperlinks : IXLHyperlinks, ISheetListener
{
    private readonly XLWorksheet _worksheet;

    /// <summary>
    /// XLHyperlink doesn't contain range, it is user created and only then it is associated with an area in a sheet.
    /// </summary>
    private readonly List<(XLHyperlink Link, XLSheetRange Area)> _hyperlinks = new();
    private readonly RTree<XLHyperlink> _areaIndex = new();
    private readonly Dictionary<XLHyperlink, XLSheetRange> _linkIndex = new();

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

        // Styles are responsibility of style slice => only shift areas
        foreach (var (link, linkArea) in _hyperlinks.ToArray())
        {
            var (success, newLinkArea) = reposition(linkArea);
            if (!success)
                continue; // Partial cover, don't move.

            if (linkArea == newLinkArea)
                continue; // Nothing changed

            Remove(link);
            if (newLinkArea is not null)
                Add(newLinkArea.Value, link);
        }
    }

    #endregion ISheetListener

    public IEnumerator<XLHyperlink> GetEnumerator()
    {
        // Enumerate in same order it was loaded and will be saved
        return _hyperlinks.Select(static x => x.Link).GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <inheritdoc />
    public bool Delete(XLHyperlink hyperlink)
    {
        if (!Remove(hyperlink, out var linkArea))
            return false;

        ClearHyperlinkStyle(linkArea);
        return true;
    }

    /// <inheritdoc />
    public bool Delete(IXLAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
            return false;

        var cellPoint = XLSheetPoint.FromAddress(address);
        if (!TryGet(cellPoint, out var cellLink))
            return false;

        Remove(cellLink);
        ClearHyperlinkStyle(cellPoint);
        return true;
    }

    /// <inheritdoc />
    public XLHyperlink Get(IXLAddress address)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
            throw new KeyNotFoundException("Address is for a different sheet.");

        var point = XLSheetPoint.FromAddress(address);
        if (!TryGet(point, out var link))
            throw new KeyNotFoundException($"No hyperlink is defined for cell {point}.");

        return link;
    }

    /// <inheritdoc />
    public bool TryGet(IXLAddress address, [NotNullWhen(true)] out XLHyperlink? hyperlink)
    {
        if (address.Worksheet is not null && address.Worksheet != _worksheet)
        {
            hyperlink = null;
            return false;
        }

        var point = XLSheetPoint.FromAddress(address);
        return TryGet(point, out hyperlink);
    }

    internal bool HasHyperlink(XLSheetPoint point)
    {
        var areaNodes = new List<RTree<XLHyperlink>.Node>();
        return _areaIndex.GetNodes(point, areaNodes).Count > 0;
    }

    /// <summary>
    /// Set a hyperlink of a single cell. Doesn't modify style, ignores hyperlinks with areas that
    /// cover the cell.
    /// </summary>
    internal void SetCellHyperlink(XLSheetPoint point, XLHyperlink? link)
    {
        // We only care about links defined for individual cell, not any link that covers the cell
        var pointNodes = new List<RTree<XLHyperlink>.Node>();
        _areaIndex.GetNodes(point, pointNodes);
        foreach (var existingLink in pointNodes)
            Remove(existingLink.Data);
        
        if (link is null)
            return;

        Add(point, link);
    }

    internal bool TryGet(XLSheetPoint point, [NotNullWhen(true)] out XLHyperlink? hyperlink)
    {
        var cellArea = new XLSheetRange(point);
        var areaNodes = new List<RTree<XLHyperlink>.Node>();
        _areaIndex.GetNodes(cellArea, areaNodes);

        if (areaNodes.Count == 0)
        {
            hyperlink = null;
            return false;
        }

        if (areaNodes.Count == 1)
        {
            hyperlink = areaNodes[0].Data;
            return true;
        }

        // There are multiple areas for the point. When hyperlink areas overlap, Excel opens
        // the last one. So it is likely the correct one. But take a random one (areaNodes are
        // not guaranteed to be in correct order), because this API is just beyond any hope and
        // will be completely scrapped ASAP.
        hyperlink = areaNodes[^1].Data;
        return true;
    }

    internal XLCell? GetCell(XLHyperlink hyperlink)
    {
        if (!_linkIndex.TryGetValue(hyperlink, out var area))
            return null;

        return new XLCell(_worksheet, area.FirstPoint);
    }

    private void Add(XLSheetRange linkArea, XLHyperlink link)
    {
        if (link.Container is not null && link.Container != this)
            throw new InvalidOperationException("Hyperlink is attached to a different worksheet. Either remove it from the original worksheet or create a new hyperlink.");

        if (_linkIndex.ContainsKey(link))
            return;

        _linkIndex.Add(link, linkArea);
        _areaIndex.Insert(new RTree<XLHyperlink>.Node(linkArea, link));
        _hyperlinks.Add((link, linkArea));
        link.Container = this;
        Debug.Assert(_hyperlinks.Count == _linkIndex.Count);
        Debug.Assert(_hyperlinks.Count == _areaIndex.Count);
    }

    private void Remove(XLHyperlink link)
    {
        Remove(link, out _);
    }

    private bool Remove(XLHyperlink link, out XLSheetRange area)
    {
        if (!_linkIndex.Remove(link, out area))
            return false;

        _areaIndex.Delete(new RTree<XLHyperlink>.Node(area, link));
        _hyperlinks.RemoveAll(x => x.Link == link);
        link.Container = null;
        Debug.Assert(_hyperlinks.Count == _linkIndex.Count);
        Debug.Assert(_hyperlinks.Count == _areaIndex.Count);
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
