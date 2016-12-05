# BitmapTracer
This is a command line tool that traces a shape, then generates VBA macro that draws the shape in an Excel workbook.

Assumes Python 2.7 is available from the command line.  Currently this only handles black and white shapes.  (See MapOfTexas.bmp)

Example, from the command line:

    C:\Location> python trace.py MapOfTexas.bmp > macro.bas

The .bas file will contain macro that can be imported into an Excel workbook as a VBA module.  The macro draws the outline of the bitmap image as an Excel shape.

