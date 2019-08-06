*********************
Organizing Worksheets
*********************
Adding a worksheet
------------------
Some of the methods we will discuss in this tutorial will come in handy when modifying a preexisting file, although you can also use them on a new Workbook. Just so we can have a clear overview in this tutorial, we will quickly reiterate over how to add a worksheet. We can achieve this by specifying a name, but as you'll see later on in this tutorial, it can also be useful to specify a position. So for now we'll add a worksheet named "Export", and give it position 3.

.. code-block:: csharp
   :caption: C#
   
   wb.Worksheets().Add("Export", 3);


Removing a worksheet
--------------------
In previous tutorials we have already looked at how to add a worksheet do your workwook, but it is also possible to remove a worksheet from your workbook. This is done in the same manner as adding a worksheet, and you have the choice whether to access the sheet by name or by position. In this case we are going to look at how to remove the worksheet at position 2.

.. code-block:: csharp
   :caption: C#
   
   wb.Worksheet(2).Delete();

This is really all there is to removing a worksheet from your workbook.

Moving worksheets
-----------------
It is also possible to rearrange the order of the worksheets in the workbook. We can do this by giving the worksheet a new position, for which we once again have the possibility of accessing the worksheet either by name or by position. Since we used the position in our previous example, we will now use a name. Let's take the worksheet we added in the first part of this tutorial, which we want to move to position 2. We can do that with the following code:

.. code-block:: csharp
   :caption: C#
   
   wb.Worksheet("Export").Position = 2;
   
There you have it. Now you have all the tools you need to organize the worksheets in your workbook.
