using System;

namespace ClosedXML.Excel;

/// <summary>
/// A field that describes calculation of value to display in the <see cref="XLPivotAreaType.Data"/>
/// area of pivot table.
/// </summary>
internal class XLPivotDataField : IXLPivotValue
{
    private readonly XLPivotTable _pivotTable;
    private int _baseField = -1;
    private XLPivotCalculation _showDataAsFormat = XLPivotCalculation.Normal;
    private XLPivotSummary _subtotal = XLPivotSummary.Sum;

    internal XLPivotDataField(XLPivotTable pivotTable, int field)
    {
        _pivotTable = pivotTable;
        Field = field;
    }

    /// <summary>
    /// Custom name of the data field (e.g. <em>Sum of Sold</em>). Can be left empty to keep same
    /// as source name. Use <see cref="CustomName"/> to get value with fallback.
    /// </summary>
    /// <remarks>
    /// For data fields, the name is duplicated at <see cref="XLPivotTableField.Name"/> and here.
    /// This property has a preference.
    /// </remarks>
    internal string? DataFieldName { get; set; }

    /// <summary>
    /// Field index to <see cref="XLPivotTable.PivotFields"/>.
    /// </summary>
    /// <remarks>
    /// Unlike axis, this field index can't be <c>-2</c> for data fields. That field can't be in
    /// the data area.
    /// </remarks>
    internal int Field { get; }

    /// <summary>
    /// An aggregation function that calculates the value to display in the data cells of pivot area.
    /// </summary>
    public XLPivotSummary Subtotal
    {
        get => _subtotal;
        init => _subtotal = value;
    }

    /// <summary>
    /// A calculation takes value calculated by <see cref="Subtotal"/> aggregation and transforms
    /// it into the final value to display to the user. The calculation might need
    /// <see cref="BaseField"/> and/or <see cref="BaseItem"/>.
    /// </summary>
    public XLPivotCalculation ShowDataAsFormat
    {
        get => _showDataAsFormat;
        init => _showDataAsFormat = value;
    }

    /// <summary>
    /// Index to the base field (<see cref="XLPivotTable.PivotFields"/>) when
    /// <see cref="ShowDataAsFormat"/> needs a field for its calculation.
    /// </summary>
    public int BaseField
    {
        get => _baseField;
        init => _baseField = value;
    }

    /// <summary>
    /// Index to the base item of <see cref="BaseField"/> when <see cref="ShowDataAsFormat"/> needs
    /// an item for its calculation.
    /// </summary>
    public uint BaseItem { get; init; } = 1048832;

    /// <summary>
    /// Formatting to apply to the data field. If <see cref="XLPivotFormat"/> disagree, this has precedence.
    /// </summary>
    public int? NumberFormatId { get; set; }

    /// <summary>
    /// A custom number format string for the value of the field. If empty, use <see cref="NumberFormatId"/>.
    /// Because pivot field uses only reference to the styles table, this is transformed before saving.
    /// </summary>
    internal string NumberFormatCode { get; set; } = string.Empty;

    internal bool HasCustomNumberFormat => !string.IsNullOrEmpty(NumberFormatCode);

    #region IXLPivotValue

    public string? BaseFieldName
    {
        get
        {
            var sourceNames = _pivotTable.PivotCache.FieldNames;
            if (_baseField < 0 || _baseField >= sourceNames.Count)
                return null;

            return sourceNames[_baseField];
        }
        set
        {
            if (value is null)
            {
                _baseField = -1;
                return;
            }

            if (!_pivotTable.TryGetSourceNameFieldIndex(value, out var index))
            {
                throw new ArgumentOutOfRangeException($"Source name '{value}' not found.");
            }

            _baseField = index;
        }
    }

    public XLCellValue BaseItemValue
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    public XLPivotCalculation Calculation
    {
        get => ShowDataAsFormat;
        set => _showDataAsFormat = value;
    }

    // 1048828 prev
    // 1048829 next
    public XLPivotCalculationItem CalculationItem
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    public string CustomName
    {
        get => DataFieldName ?? _pivotTable.PivotCache.FieldNames[Field];
        set => DataFieldName = value;
    }

    public IXLPivotValueFormat NumberFormat => new XLPivotValueFormat(this);

    public string SourceName => _pivotTable.PivotFields[Field].Name ?? _pivotTable.PivotCache.FieldNames[Field];

    public XLPivotSummary SummaryFormula
    {
        get => Subtotal;
        set => _subtotal = value;
    }

    public IXLPivotValue SetBaseFieldName(string value)
    {
        BaseFieldName = value;
        return this;
    }

    public IXLPivotValue SetBaseItemValue(XLCellValue value)
    {
        BaseItemValue = value; return this;
    }

    public IXLPivotValue SetCalculation(XLPivotCalculation value)
    {
        Calculation = value;
        return this;
    }

    public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value)
    {
        CalculationItem = value; return this;
    }

    public IXLPivotValue SetSummaryFormula(XLPivotSummary value)
    {
        SummaryFormula = value;
        return this;
    }

    public IXLPivotValueCombination ShowAsDifferenceFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XLPivotCalculation.DifferenceFrom);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValue ShowAsIndex()
    {
        return SetCalculation(XLPivotCalculation.Index);
    }

    public IXLPivotValue ShowAsNormal()
    {
        return SetCalculation(XLPivotCalculation.Normal);
    }

    public IXLPivotValueCombination ShowAsPercentageDifferenceFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XLPivotCalculation.PercentageDifferenceFrom);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValueCombination ShowAsPercentageFrom(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        SetCalculation(XLPivotCalculation.PercentageOf);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValue ShowAsPercentageOfColumn()
    {
        return SetCalculation(XLPivotCalculation.PercentageOfColumn);
    }

    public IXLPivotValue ShowAsPercentageOfRow()
    {
        return SetCalculation(XLPivotCalculation.PercentageOfRow);
    }

    public IXLPivotValue ShowAsPercentageOfTotal()
    {
        return SetCalculation(XLPivotCalculation.PercentageOfTotal);
    }

    public IXLPivotValue ShowAsRunningTotalIn(string fieldSourceName)
    {
        BaseFieldName = fieldSourceName;
        return SetCalculation(XLPivotCalculation.RunningTotal);
    }

    #endregion IXPivotValue
}
