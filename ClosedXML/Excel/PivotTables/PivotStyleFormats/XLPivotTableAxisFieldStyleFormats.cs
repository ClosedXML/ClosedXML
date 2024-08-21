#nullable disable
using System.Linq;

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
            return new XLPivotStyleFormat(
                _pivotTable,
                AreaIsHeader,
                CreateHeaderArea);

            bool AreaIsHeader(XLPivotArea area)
            {
                return
                    area.References.Count == 0 &&
                    area.Field == _axisField.Offset &&
                    area.Type == XLPivotAreaType.Button &&
                    area.DataOnly &&
                    !area.LabelOnly &&
                    !area.GrandRow &&
                    !area.GrandCol &&
                    area.CacheIndex == false &&
                    area.Offset is null &&
                    !area.CollapsedLevelsAreSubtotals &&
                    area.Axis == fieldAxis &&
                    area.FieldPosition == fieldPosition;
            }

            XLPivotArea CreateHeaderArea()
            {
                return new XLPivotArea
                {
                    Field = _axisField.Offset,
                    Type = XLPivotAreaType.Button,
                    Axis = fieldAxis,
                    FieldPosition = (uint)fieldPosition,
                };
            }
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
            return new XLPivotStyleFormat(
                _pivotTable,
                area => AreaBelongsToField(area, XLPivotStyleFormatElement.Label),
                () => CreateFieldArea(XLPivotStyleFormatElement.Label));
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
            var subtotals = _axisField.Subtotals;
            var subtotalArea = new XLPivotArea
            {
                Outline = false
            };
            subtotalArea.AddReference(new XLPivotReference
            {
                Field = unchecked((uint)_axisField.Offset),
                DefaultSubtotal = subtotals.Contains(XLSubtotalFunction.Automatic),
                SumSubtotal = subtotals.Contains(XLSubtotalFunction.Sum),
                CountASubtotal = subtotals.Contains(XLSubtotalFunction.Count),
                AvgSubtotal = subtotals.Contains(XLSubtotalFunction.Average),
                MaxSubtotal = subtotals.Contains(XLSubtotalFunction.Maximum),
                MinSubtotal = subtotals.Contains(XLSubtotalFunction.Minimum),
                ProductSubtotal = subtotals.Contains(XLSubtotalFunction.Product),
                CountSubtotal = subtotals.Contains(XLSubtotalFunction.CountNumbers),
                StdDevSubtotal = subtotals.Contains(XLSubtotalFunction.StandardDeviation),
                StdDevPSubtotal = subtotals.Contains(XLSubtotalFunction.PopulationStandardDeviation),
                VarSubtotal = subtotals.Contains(XLSubtotalFunction.Variance),
                VarPSubtotal = subtotals.Contains(XLSubtotalFunction.PopulationVariance),
            });

            return new XLPivotStyleFormat(
                _pivotTable,
                area => XLPivotAreaComparer.Instance.Equals(area, subtotalArea),
                () => subtotalArea);
        }
    }

    #endregion

    private bool AreaBelongsToField(XLPivotArea area, XLPivotStyleFormatElement element)
    {
        if (area.References.Count != 1)
            return false;

        var field = area.References[0].Field;
        if (field is null || (int)field != _axisField.Offset)
            return false;

        return
            area.Field is null &&
            area.Type == XLPivotAreaType.Normal &&
            area.DataOnly == (element == XLPivotStyleFormatElement.Data) &&
            area.LabelOnly == (element == XLPivotStyleFormatElement.Label) &&
            !area.GrandRow &&
            !area.GrandCol &&
            area.CacheIndex == false &&
            area.Offset is null &&
            !area.CollapsedLevelsAreSubtotals &&
            area.Axis is null &&
            area.FieldPosition is null;
    }

    private XLPivotArea CreateFieldArea(XLPivotStyleFormatElement element)
    {
        var area = new XLPivotArea
        {
            DataOnly = (element == XLPivotStyleFormatElement.Data),
            LabelOnly = (element == XLPivotStyleFormatElement.Label),
        };
        area.AddReference(new XLPivotReference
        {
            Field = (uint)_axisField.Offset,
        });
        return area;
    }
}
