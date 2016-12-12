# BitmapTracer
This is a command line tool that traces a shape, then generates VBA macro that draws the shape in an Excel workbook.

Assumes Python 2.7 is available from the command line.  Currently this only handles black and white shapes.  (See MapOfTexas.bmp)

Demo:

Here's a bitmap image Texas.

![Texas](MapOfTexas.bmp?raw=true "Texas")

At the command line, pass the .bmp file to trace.py:

    C:\Location> python trace.py MapOfTexas.bmp > macro.bas

The newly created .bas file will contain a macro that can be imported into an Excel workbook as a VBA module.  The macro draws the outline of the bitmap image as an Excel shape.

Here's what the result looks like in Excel:

![Screenshot](ExcelSnip.png?raw=true "Excel")

Future plans:
    Introduce a classifier to convert raw images into black and white bitmap files. 
