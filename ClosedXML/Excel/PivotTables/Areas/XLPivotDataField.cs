using System;
using System.Diagnostics;

namespace ClosedXML.Excel;

/// <summary>
/// A field that describes calculation of value to display in the <see cref="XLPivotAreaType.Data"/>
/// area of pivot table.
/// </summary>
internal class XLPivotDataField : IXLPivotValue
{
    private const int BaseFieldDefaultValue = -1;
    private const int BaseItemPreviousValue = 1048828;
    private const int BaseItemNextValue = 1048829;
    private const int BaseItemDefaultValue = 1048832;

    private readonly XLPivotTable _pivotTable;

    private int _baseField = BaseFieldDefaultValue;
    private uint _baseItem = BaseItemDefaultValue;
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
        set => _subtotal = value;
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
    public uint BaseItem
    {
        get => _baseItem;
        init => _baseItem = value;
    }

    /// <summary>
    /// Formatting to apply to the data field. If <see cref="XLPivotFormat"/> disagree, this has precedence.
    /// </summary>
    internal XLNumberFormatValue? NumberFormatValue { get; set; }

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
                _baseField = BaseFieldDefaultValue;
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
        get
        {
            var baseFieldSpecified = _baseField != BaseFieldDefaultValue;
            if (!baseFieldSpecified)
                return Blank.Value;

            var baseItemSpecified = _baseItem != BaseItemDefaultValue;
            if (!baseItemSpecified)
                return Blank.Value;

            if (_baseItem == BaseItemPreviousValue)
                return Blank.Value;

            if (_baseItem == BaseItemNextValue)
                return Blank.Value;

            var baseField = _pivotTable.PivotFields[_baseField];
            var fieldItem = baseField.Items[checked((int)BaseItem)];
            return fieldItem.GetValue() ?? Blank.Value;
        }
        set
        {
            if (_baseField == BaseItemDefaultValue)
                throw new InvalidOperationException("Base field not specified for the field.");

            var field = _pivotTable.PivotFields[_baseField];
            var fieldItem = field.GetOrAddItem(value);
            var itemIndex = fieldItem.ItemIndex ?? BaseFieldDefaultValue;
            _baseItem = checked((uint)itemIndex);
        }
    }

    public XLPivotCalculation Calculation
    {
        get => ShowDataAsFormat;
        set => _showDataAsFormat = value;
    }

    public XLPivotCalculationItem CalculationItem
    {
        get
        {
            return _baseItem switch
            {
                BaseItemPreviousValue => XLPivotCalculationItem.Previous,
                BaseItemNextValue => XLPivotCalculationItem.Next,
                _ => XLPivotCalculationItem.Value,
            };
        }
        set
        {
            switch (value)
            {
                case XLPivotCalculationItem.Previous:
                    _baseItem = BaseItemPreviousValue;
                    break;
                case XLPivotCalculationItem.Next:
                    _baseItem = BaseItemNextValue;
                    break;
                case XLPivotCalculationItem.Value:
                    // Calculation value should be set in tandem with the base item value.
                    // Base item other than prev/next special constants is implicitly a value.
                    if (BaseItem is BaseItemPreviousValue or BaseItemNextValue)
                    {
                        // If value is not yet set, just use unspecified value. User should
                        // set value by calling `BaseItemValue` after calling this, but Excel
                        // accepts valid base field with unspecified item without need to repair.
                        _baseItem = BaseItemDefaultValue;
                    }

                    // When base item is not a valid reference to the field.Items, Excel
                    // tries to repair the workbook, so user should always set base value.
                    break;
                default:
                    throw new UnreachableException();
            }
        }
    }

    public string CustomName
    {
        get => DataFieldName ?? _pivotTable.PivotFields[Field].Name ?? _pivotTable.PivotCache.FieldNames[Field];
        set => DataFieldName = value;
    }

    public IXLPivotValueFormat NumberFormat => new XLPivotValueFormat(this);

    public string SourceName => _pivotTable.PivotCache.FieldNames[Field];

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
        BaseItemValue = value;
        return this;
    }

    public IXLPivotValue SetCalculation(XLPivotCalculation value)
    {
        Calculation = value;
        return this;
    }

    public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value)
    {
        CalculationItem = value;
        return this;
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
