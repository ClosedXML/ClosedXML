#nullable disable

namespace ClosedXML.Excel;

internal class XLPivotTableAxisFieldStyleFormats : IXLPivotFieldStyleFormats
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotTableAxisField _axisField;

    public XLPivotTableAxisFieldStyleFormats(XLPivotTable pivotTable, XLPivotTableAxisField axisField)
    {
        _pivotTable = pivotTable;
        _axisField = axisField;
    }

    #region IXLPivotFieldStyleFormats

    public IXLPivotValueStyleFormat DataValuesFormat { get; }

    public IXLPivotStyleFormat Header
    {
        get
        {
            /*
             * <x:pivotArea field="4"
             *  type="button"
             *  axis="axisCol"
             *  fieldPosition="0"/>
             *
             * The area must have field position and axis, otherwise the style is not correctly
             * displayed.
             */
            // If table is not compact, each field has it's own header and thus pivot area must
            // contain correct position of the field on the axis. If table is compact, there is
            // only one header and its position is first in axis, because it's the only one.
            var fieldPosition = _pivotTable.Compact ? 0 : _axisField.Position;
            var fieldAxis = _axisField.Axis;
            var headerArea = new XLPivotArea
            {
                Field = _axisField.Offset,
                Type = XLPivotAreaType.Button,
                Axis = fieldAxis,
                FieldPosition = (uint)fieldPosition,
            };

            return new XLPivotStyleFormat(
                _pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, headerArea),
                () => headerArea);
        }
    }

    public IXLPivotStyleFormat Label
    {
        get
        {
            /* <x:pivotArea type="normal"
			 *              dataOnly="0"
			 *              labelOnly="1">
			 *     <x:references count="1">
			 *         <x:reference field="4"/>
			 *	   </x:references>
			 * </x:pivotArea>
             */
            var labelArea = new XLPivotArea
            {
                DataOnly = false,
                LabelOnly = true,
            };
            labelArea.AddReference(new XLPivotReference
            {
                Field = (uint)_axisField.Offset,
            });

            return new XLPivotStyleFormat(
                _pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, labelArea),
                () => labelArea);
        }
    }

    public IXLPivotStyleFormat Subtotal
    {
        get
        {
            /* <pivotArea outline="0">
             *   <references count="1">
             *     <reference field="0"
             *                count="0"
             *                defaultSubtotal="1"/>
             *   </references>
             * </pivotArea>
             */
            // Subtotal fields in reference can't mix default and custom subtotals. It always must
            // reference only one type. Excel doesn't select correct area if they are mixed.
            // The outline flag has weird behavior, but is required for subtotals of last field in
            // an axis with multiple fields (i.e. subtotals are displayed at the bottom).
            var subtotals = _axisField.Subtotals.ToHashSet();
            var subtotalArea = new XLPivotArea
            {
                Outline = false
            };
            subtotalArea.AddReference(new XLPivotReference
            {
                Field = unchecked((uint)_axisField.Offset),
                Subtotals = subtotals,
            });

            return new XLPivotStyleFormat(
                _pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, subtotalArea),
                () => subtotalArea);
        }
    }

    #endregion
}
