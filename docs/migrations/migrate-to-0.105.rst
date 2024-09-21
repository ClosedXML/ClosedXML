#############################
Migration from 0.104 to 0.105
#############################

***********************
IXLRangeBase.Hyperlinks
***********************

Property ```IXLRangeBase.Hyperlinks``` has been moved to ```IXLWorksheet.Hyperlinks```.
The original place could only list hyperlinks and didn't provide correct
functionality (e.g. removal of hyperlinks).

*************
IXLHyperlinks
*************

IXLHyperlinks.Add was removed
-----------------------------

The method was removed, because there is no way to actually use it. It has a
parameter ```XLHyperlink```, but doesn't have an range where should hyperlink
be created. ```XLHyperlink``` can't have a ```XLHyperlink.Cell``` property set
to a valid value, because that would mean hyperlink is already attached to
a worksheet.

IXLHyperlinks.TryDelete was removed
-----------------------------------

It worked same way as ```IXLHyperlinks.Delete```.

IXLHyperlinks.Delete returns bool
---------------------------------

Originally, it returned ```void```. The ```bool``` value indicates if a
hyperlink was present and thus removed.
