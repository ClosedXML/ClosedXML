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

    public IXLPivotStyleFormat Subtotal { get; }

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
