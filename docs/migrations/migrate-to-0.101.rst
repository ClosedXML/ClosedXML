#############################
Migration from 0.100 to 0.101
#############################

******************
Dependency updates
******************

Library ``SixLabors.Fonts`` has been updated from *1.0.0-beta18* to
*1.0.0-beta19* to due to bug fixing of font measurement.

*********************
Enum underlaying type
*********************

Enums `XLAlignmentReadingOrderValues`, `XLAlignmentHorizontalValues` and
`XLAlignmentVerticalValues` now use `byte` as an underlaying type instead of
`int`. Smaller underlaying type saves memory due to alignment.

**************
Graphic engine
**************

Graphic engine has a new method: `IXLGraphicEngine.GetGlyphBox`. See XML doc
for more details.