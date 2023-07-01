#############################
Migration from 0.102 to 0.103
#############################

***********
IXLPhonetic
***********

``IXLPhonetic`` no longer has a setter for its ``IXLPhonetic.Text``,
``IXLPhonetic.Start`` and ``IXLPhonetic.End`` properties.

Use ``IXLPhonetics.ClearText()`` and ``IXLPhonetics.Add(String text, Int32 start, Int32 end)``
to redo the phonetic hints.

