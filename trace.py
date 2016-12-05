# Note - This code requires a black and white .bmp file as input

# Ensure you have the python imaging library (PIL) is available.  This is configured for my machine.
import sys
if sys.path.count(r'c:\users\mike\anaconda\lib\site-packages')==0:
    sys.path.append(r'c:\users\mike\anaconda\lib\site-packages')

# PIL is in the location appended to the PYTHONPATH variable.
from PIL import Image

# Image.open() mode is read-only by default...use .Convert() to 'RGB' to enable write access.
imgPath = sys.argv[1]
img = Image.open(imgPath).convert('RGB')
pixels = img.load()
RgbOutline =(0,0,0)
RgbBackground = (255,255,255)

# Some color assumed absent from the image.
RgbGreenScreen = (0,255,0)

# Trace the border: if a pixel has 1 black neighbor and 1 white neighbor, in either dimension, color it bright green:
def ProduceOutline():
    for i in range(1,img.size[0]-1):
        for j in range(1,img.size[1]-1):
		    # 1st IF is for top/left borders. 2nd if for bottom/right borders.
            if (img.getpixel((i-1, j)) == RgbBackground and img.getpixel((i+1, j)) == RgbOutline) or \
               (img.getpixel((i, j-1)) == RgbBackground and img.getpixel((i, j+1)) == RgbOutline):
                pixels[i,j] = RgbGreenScreen
            if (img.getpixel((i+1, j)) == RgbBackground and img.getpixel((i-1, j)) == RgbOutline) or \
               (img.getpixel((i, j+1)) == RgbBackground and img.getpixel((i, j-1)) == RgbOutline):
                pixels[i,j] = RgbGreenScreen

    # Remove the black fill from the image and color the border pixels black:
    for i in range(img.size[0]):
        for j in range(img.size[1]):
            if img.getpixel((i, j)) != RgbGreenScreen:
                pixels[i,j] = RgbBackground


unorderedCoordinates = []
def collectOutlinePoints():
    for i in range(img.size[0]):
        for j in range(img.size[1]):
            if pixels[i,j] == RgbGreenScreen:
             unorderedCoordinates.append([i, j])
    return(unorderedCoordinates)


def dist(pair1, pair2):
    x1, x2 = pair1[0], pair2[0]
    y1, y2 = pair1[1], pair2[1]
    dist = ( (x1 - x2)**2 + (y1 - y2)**2)
    return(dist)


def orderOutlinePoints(unorderedCoordinatesInput):
    adjacentPairs = [unorderedCoordinatesInput.pop(0)]
    while(len(unorderedCoordinatesInput) > 0):
        currentPoint = adjacentPairs[-1]
        distances = [dist(currentPoint, xy) for xy in unorderedCoordinatesInput ]
        minVal = min(distances)
        minElement = distances.index(minVal)
        adjacentPairs.append(unorderedCoordinatesInput.pop(minElement))
        distances.pop(minElement)
        while(distances.count(minVal) > 0):
            minElement = distances.index(minVal)
            unorderedCoordinatesInput.pop(minElement)
            distances.pop(minElement)
        
        # Finally, ignore any straggling points that may have been 
        # skipped.  Prevents criss-crossing effect in Excel workbook.
        # Rule - only investigate this if 99% of the points have already
        # been accounted for.
        if(len(adjacentPairs) > 99 * len(unorderedCoordinatesInput)):
            gapsToClose = [ dist(adjacentPairs[0], xy) for xy in unorderedCoordinatesInput ]
            if min(gapsToClose) < min(distances):
                break
        
    return(adjacentPairs)


def GenerateVBACode(orderedPoints):
    """ Prints VBA code for drawing the .bmp shape. """
    vbaShapeBuilderSub =  """
Public Sub BuildShape()
    ''' This code was generated programmatically via Python 2.7'''
    Application.Screenupdating = False
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim s As Object
    Dim lngCoordinateIterator As Long

    'Uncomment this loop to remove all preexisting shapes on the sheet:
    ''''For Each s In ws.Shapes
    ''''    s.Delete
    ''''Next s

    ' As needed, adjust offset to avoid collision with top/left borders of sheet.
    Dim origin As Integer: origin = 100

    'Use the first 2 coordinate pairs to build a line segment:\n"""
    vbaShapeBuilderSub += "    Set s = ws.Drawings.Add("
    vbaShapeBuilderSub += "    origin + {0}, ".format(orderedPoints[0][0], 'utf-8')
    vbaShapeBuilderSub += "    origin + {0}, ".format(orderedPoints[0][1], 'utf-8')
    vbaShapeBuilderSub += "    origin + {0}, ".format(orderedPoints[1][0], 'utf-8')
    vbaShapeBuilderSub += "    origin + {0}, False)".format(orderedPoints[1][1], 'utf-8') + "\n"

    firstPoint = orderedPoints.pop(0)

    #Then iteratively add to the line segment by appending vertices to the drawing object:
    for xy in orderedPoints:
        vbaShapeBuilderSub += "    s.addvertex origin +  {0}, origin + {1}\n".format(xy[0], xy[1])

    # Close off the shape:
    vbaShapeBuilderSub += "    s.AddVertex origin + {0}, origin + {1}\n".format(firstPoint[0], firstPoint[1])
    vbaShapeBuilderSub += "    Application.Screenupdating = True\n"
    vbaShapeBuilderSub += "End Sub"
    print(vbaShapeBuilderSub)


""" Execution """
ProduceOutline()
C = collectOutlinePoints()
outline = orderOutlinePoints(C)
GenerateVBACode(outline)