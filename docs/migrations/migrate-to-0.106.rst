#############################
Migration from 0.105 to 0.106
#############################

**************
Number formats
**************

Format is always there
----------------------

Number format ```IXLNumberFormatBase.Format``` previously returned an empty
string for predefined formats. It now  returns actually used predefined format
code instead.

Number format id
----------------

The number format setters (```IXLNumberFormatBase.NumberFormatId```,
```IXLNumberFormat.SetNumberFormatId(int)``` and ```IXLPivotValueFormat.SetNumberFormatId(int)```)
now throw an ```ArgumentOutOfRangeException``` when supplied number format id is
not a predefined format id from ```XLPredefinedFormat```.
