***********************
Introduction to styling
***********************
In our last tutorial we created a simple workbook, with a cell which says "Hello World". In this tutorial we are going to expand on what we did previously, by adding styling to our worksheet. However, before we start on that, lets add a few more cells, so we have a little bit more data to experiment with. With those changes, our file now looks like this:

.. literalinclude:: ../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorld.bfs.cs
   :diff: ../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorld.orig.cs
   :language: csharp
   :linenos:

-------------------
Background coloring
-------------------
To make it a little bit easier for ourselves, we our going to create a range before we add any styling. A range is basically a collection of cells which we can manipulate at once. For example, one of the advantages of using a range is that we can set a background color for the entire range, instead of having to set the background color for each individual cell. The first thing we are now going to do, is create a range, just like this:

.. code-block:: csharp
   :caption: C#
   
   IXLRange range = worksheet.Range(worksheet.Cell(4,2).Address, worksheet.Cell(6,4).Address);

As you can see, we only have to specify the top left and bottom left cell of the cells we want to include in the range. However, this is just one way of creating a range, we will take a look at the other ways to create a range when we discuss ranges themselves. 

-----------
Cell border
-----------
Now that we have defined a range, we can use that range to set a border around the entire range, like so:

.. code-block:: csharp
   :caption: C#

   range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

There are plenty more styling options available in ClosedXML, but we will discuss those later on. With the modifications we made in the tutorial, our file now looks like this:

.. literalinclude:: ../ClosedXML_Examples/Tutorials/HelloWorld/HelloWorldStyling.as.cs
   :language: csharp
   :linenos: