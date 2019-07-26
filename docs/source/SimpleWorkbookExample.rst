****************************
Creating your first workbook
****************************
In this example we will show you how to create a simple workbook with one worksheet, along with a few cells containing different datatypes. At the end of this tutorial you can see the full code, or you can get the source code from our github page `here <https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML_Examples/HelloWorld.cs>`_.

The first thing we are going to do, is create a HelloWorld class with a method Create, which takes a filepath as a parameter. That should look something like this:

.. literalinclude:: ../../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorld.orig.cs
   :language: csharp
   :lines: 4-9, 16-18
   :linenos:

The next step is to add a using statement for the ClosedXML package at the top of our file, like this:

.. literalinclude:: ../../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorld.orig.cs
   :language: csharp
   :lines: 1-9, 16-
   :emphasize-lines: 2
   :linenos:

As you can probably see, we are in the ClosedXML_Examples namespace, for now that is not important.

-------------------
Creating a workbook
-------------------
After you have added your using statement it is now time to create our workbook. In ClosedXML this can be done very easily, like so:

.. code-block:: csharp
   :caption: C#

   IXLWorkbook workbook = new XLWorkbook();

------------------
Adding a worksheet
------------------
Now that we have a workbook, we can add a worksheet to our workbook.

.. code-block:: csharp
   :caption: C#

   IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");

-------------
Adding a cell
-------------
Our last step is to add a cell to the worksheet. To do this we have two options, ``.Value = value`` and ``.SetValue(value)``. We won't go into detail about the difference here, but we will discuss it later on. You also have two options to select a cell. You can either use integers, like so:

.. code-block:: csharp
   :caption: C#

   worksheet.Cell(1,1).Value = "Hello World";

Or you can use the address, like so:

.. code-block:: csharp
   :caption: C#

   worksheet.Cell("A1").Value = "Hello World";

For small sheets the difference is neglegible, but for larger sheets it is better to use the integers.

Saving the workbook
-------------------
The last step left for us to do, is to save the workbook. ClosedXML has a method to do so, which you can pass a filepath to tell it where to store the file, and with what name.

.. code-block:: csharp
   :caption: C#

   workbook.SaveAs(filePath);

End result
----------
If we combine all this code, this is what we have:

.. literalinclude:: ../../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorld.orig.cs
   :language: csharp
   :linenos: