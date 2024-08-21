using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

internal class XLPivotAreaComparer : IEqualityComparer<XLPivotArea>
{
    private readonly XLPivotReferenceComparer _referenceComparer = new();

    public static readonly XLPivotAreaComparer Instance = new();

    public bool Equals(XLPivotArea? x, XLPivotArea? y)
    {
        if (ReferenceEquals(x, y))
            return true;

        if (x is null)
            return false;

        if (y is null)
            return false;

        return x.References.SequenceEqual(y.References, _referenceComparer) &&
               Nullable.Equals(x.Field, y.Field) &&
               x.Type == y.Type &&
               x.DataOnly == y.DataOnly &&
               x.LabelOnly == y.LabelOnly &&
               x.GrandRow == y.GrandRow &&
               x.GrandCol == y.GrandCol &&
               x.CacheIndex == y.CacheIndex &&
               x.Outline == y.Outline &&
               Nullable.Equals(x.Offset, y.Offset) &&
               x.CollapsedLevelsAreSubtotals == y.CollapsedLevelsAreSubtotals &&
               x.Axis == y.Axis &&
               x.FieldPosition == y.FieldPosition;
    }

    public int GetHashCode(XLPivotArea obj)
    {
        var hashCode = new HashCode();
        foreach (var reference in obj.References)
            hashCode.Add(reference, _referenceComparer);

        hashCode.Add(obj.Field);
        hashCode.Add(obj.Type);
        hashCode.Add(obj.DataOnly);
        hashCode.Add(obj.LabelOnly);
        hashCode.Add(obj.GrandRow);
        hashCode.Add(obj.GrandCol);
        hashCode.Add(obj.CacheIndex);
        hashCode.Add(obj.Outline);
        hashCode.Add(obj.Offset);
        hashCode.Add(obj.CollapsedLevelsAreSubtotals);
        hashCode.Add(obj.Axis);
        hashCode.Add(obj.FieldPosition);
        return hashCode.ToHashCode();
    }

    private class XLPivotReferenceComparer : IEqualityComparer<XLPivotReference>
    {
        public bool Equals(XLPivotReference? x, XLPivotReference? y)
        {
            if (ReferenceEquals(x, y))
                return true;

            if (x is null)
                return false;

            if (y is null)
                return false;

            return x.FieldItems.SequenceEqual(y.FieldItems) &&
                   x.Field == y.Field &&
                   x.Selected == y.Selected &&
                   x.ByPosition == y.ByPosition &&
                   x.Relative == y.Relative &&
                   x.DefaultSubtotal == y.DefaultSubtotal &&
                   x.SumSubtotal == y.SumSubtotal &&
                   x.CountASubtotal == y.CountASubtotal &&
                   x.AvgSubtotal == y.AvgSubtotal &&
                   x.MaxSubtotal == y.MaxSubtotal &&
                   x.MinSubtotal == y.MinSubtotal &&
                   x.ProductSubtotal == y.ProductSubtotal &&
                   x.CountSubtotal == y.CountSubtotal &&
                   x.StdDevSubtotal == y.StdDevSubtotal &&
                   x.StdDevPSubtotal == y.StdDevPSubtotal &&
                   x.VarSubtotal == y.VarSubtotal &&
                   x.VarPSubtotal == y.VarPSubtotal;
        }

        public int GetHashCode(XLPivotReference obj)
        {
            var hashCode = new HashCode();
            foreach (var item in obj.FieldItems)
                hashCode.Add(item);

            hashCode.Add(obj.Field);
            hashCode.Add(obj.Selected);
            hashCode.Add(obj.ByPosition);
            hashCode.Add(obj.Relative);
            hashCode.Add(obj.DefaultSubtotal);
            hashCode.Add(obj.SumSubtotal);
            hashCode.Add(obj.CountASubtotal);
            hashCode.Add(obj.AvgSubtotal);
            hashCode.Add(obj.MaxSubtotal);
            hashCode.Add(obj.MinSubtotal);
            hashCode.Add(obj.ProductSubtotal);
            hashCode.Add(obj.CountSubtotal);
            hashCode.Add(obj.StdDevSubtotal);
            hashCode.Add(obj.StdDevPSubtotal);
            hashCode.Add(obj.VarSubtotal);
            hashCode.Add(obj.VarPSubtotal);
            return hashCode.ToHashCode();
        }
    }
}
