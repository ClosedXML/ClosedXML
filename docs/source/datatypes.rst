*********
Datatypes
*********
Internally, Excel makes use of Datatypes. It uses these types when running calculations, as well as for other things.

Datatypes in ClosedXML
----------------------
.. doxygenenum:: ClosedXML::Excel::XLDataType

How to use Datatypes
--------------------
A DataType can be set for multiple objects: an `IXLCell`, an `IXLRange`, an `IXLColumn` and an `IXLRangeColumn`. They can also be set for a collection of the objects mentioned before. Since the method to set the DataType is equal for all these objects, we will only show you the method for a cell. The set the DataType of a cell, you can either use the set method, like this:

.. code-block:: csharp
   :caption: C#

   worksheet.Cell( 1, 1 ).SetDataType(XLDataType.Boolean);

Thats really all there is to it. Another way to set the DataType, is to use an expression like this:

.. code-block:: csharp
   :caption: C#

   worksheet.Cell( 1, 1 ).DataType = XLDataType.Number;