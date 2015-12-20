PrtPrev: How to Add Print Preview to Visual Basic Applications

This sample GENPRINT.MAK project contains the files GENPRINT.FRM and
GENPRINT.BAS.

This sample describes how to create printing routines that can
print to the printer or to a picture box. This enables you to add
print preview capabilities to your Visual Basic applications.

Generic Printing
----------------

It would be ideal to have a generic print routine that could print
to the printer or to the screen depending on what you pass it. The
Visual Basic printer object and picture box control have many of the
same methods and properties. For example, both of these are valid:

   Printer.Print AString
   Picture1.Print AString

It would be nice if you could pass a generic object to a subroutine
and the subroutine would use the Print method off of the generic object
as in this example:

   Call PrintJob(Printer)
   Call PrintJob(Picture1)

   Sub PrintJob(GenericObject As Object)
      GenericObject.Print AString
   End Sub

Unfortunately, this is not possible. The Visual Basic Printer object
is a system object, so it can't be passed as a parameter.

This leaves you with two choices in Visual Basic. You could create two
routines -- one for printing to the printer and one for print preview.
However, the code would not be reusable in your future projects. The
second approach is to write your own set of routines that can print to
the printer or a picture box based on the value of a flag. This is the
method used in the example code given below. Once you create the
routines, you can re-use them in future programs.

The example creates routines that closely mimic Visual Basic's built in
methods and properties. However, you could use this approach to create
high-level routines that greatly simplify your printing needs.

The routines work by checking the variable PrinterFlag. PrinterFlag is
True when printing is going to the printer and False when printing to
the picture box.

Here's the print routine from the example. Notice how it is just a
shell function that determines what to print to and then does it.

   Sub PrintPrint (PrintVar)
      If PrinterFlag Then
         Printer.Print PrintVar
      Else
         objPrint.Print PrintVar
      End If
   End Sub

With just a few simple routines like this, you can start to do
generic printing.

Scaling
-------

To accomplish print preview, the program must scale the output to the
picture box to match the output on the printer.

In the example, the PrintStartDoc routine initializes the printer or
picture box and sets up the scaling. The width and height of the paper
are passed to the PrintStartDoc routine. These dimensions are used to
determine the non-printable area of the printer object, find the ratio
of the picture box to the printer, re-size the picture box, and scale
the picture box. The picture box is scaled with the Scale method. After
setting the scale of the picture box, graphic methods use the new
coordinates. For an 8.5 x 11 inch piece of paper the picture box is
scaled with this command:

   Picture1.Scale (0, 0)-(8.5, 11)

The Scale method does not scale fonts. To scale the fonts, use the
ratio of the picture box height divided by the printer's height in
inches. Then multiply by this ratio to determine the correct font
size within the picture box. Here is the PrintFontSize routine that
sets the appropriate font sizes in the example:

   Sub PrintFontSize (pSize)
      If PrinterFlag Then
         Printer.FontSize = pSize
      Else
         'Sized by ratio since Scale method does not effect FontSize
         ObjPrint.FontSize = pSize * Ratio
      End If
   End Sub

The ratio used to calculate the font size can be applied to anything
you need to scale in the picture box that is not automatically scaled
by the Scale method. The ratio is also used in the PrintPicture routine
to scale pictures.

To Run the Sample:
------------------

Click the command button with the check box checked to preview the
page. Click the command button with the check box cleared to print
the page.

