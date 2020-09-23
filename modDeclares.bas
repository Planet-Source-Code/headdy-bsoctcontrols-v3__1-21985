Attribute VB_Name = "modDeclares"
' *************************************
' bsOctControls Declares Module
' ©2000 BadSoft, all rights reserved.
' *************************************

' Here are all the API calls in the program, with their public
' constants, explained as best as I can.

Option Explicit

' Rectangle
' An API call for drawing rectangles. The style of the rectangle is decided by the form's
' (or picture object's, or any object with an hDC) drawing styles, such as FillColor,
' ForeColor and FillStyle. One of the most useful API calls, but why didn't they just use
' a RECT variable?
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' DrawEdge
' An API call for drawing those nasty 3D borders, very flat borders, or in some cases
' those attractive one-pixel borders. This is also the nastiest API call I know to include
' constants for.
' Remember that comment about "Calculates space left over"? That was for a flag
' called BF_ADJUST that goes with DrawEdge, not DT_CALCRECT for DrawText! Many
' apologies to Ariad Software, but that brief description is still useless.
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, _
    ByVal edge As Long, ByVal grfFlags As Long) As Long

' DrawState
' An API call for drawing pictures in a certain state. It can be used to draw disabled
' icons, selected icons or icons in a single colour.
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
    (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, _
    ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, _
    ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

' Image types
Public Const DST_ICON = &H3        'You're using an icon
Public Const DST_BITMAP = &H4      'You're using a bitmap

' States
Public Const DSS_NORMAL = &H0   'Draw the picture normally
Public Const DSS_UNION = &H10    'Draw the picture as "selected"
Public Const DSS_DISABLED = &H20    'Draw the picture as disabled
Public Const DSS_MONO = &H80    'Draw the picture in one colour - there are ways to
                                                    'change the colour.


' GetPixel
' Works exactly like the Point method of a Form. It returns the colour of a point on the
' object.
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long) As Long
    
    
' DrawText
' An API call for drawing text. It uses the ForeColor property of the object to
' determine the colour of the text. To draw disabled text you will need to use the
' DrawStateString API call.
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
    (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, _
    lpRect As RECT, ByVal uFormat As Long) As Long

Public Const DT_WORD_ELLIPSIS = &H40000 'gets rid of extra words at the end and
                                                         'truncates the last word.

Public Const DT_MODIFYSTRING = &H10000 'The function is allowed to modify the
                                                        'passed string to fit the text inside the rectangle
                                                        'It has no effect if DT_WORD_ELLIPSIS is not
                                                        'also specified, and also only seems to work on
                                                        'single lines of text (using the DT_SINGLELINE
                                                        'flag).

Public Const DT_CALCRECT = &H400   'Calculates the exact space the text takes up
                                                        'inside the RECT variable. (Works like the
                                                        'Autosize property of a Label control.)
                               
Public Const DT_SINGLELINE = &H20   'All the text will be placed on a single line.

Public Const DT_VCENTER = &H4        'Text is centered vertically in the RECT
                                                        'variable.
                 
Public Const DT_CENTER = &H1          'Centers the text horizontally within the
                                                        'rectangle.
Public Const DT_LEFT = &H0               'Left aligns the text (by default).
Public Const DT_RIGHT = &H2            'Right aligns the text.
Public Const DT_WORDBREAK = &H10    'Allows for more than one line of text if
                                                        'necessary.
                 
' ExtFloodFill
' An API call for filling an area with colour. crColor determines the colour to fill, I think,
' and wFillType is one of the following public constants.
' SPECIAL NOTE: Avoid using the FloodFill API as it is useful only for clearing the whole
' area of the object.
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
    
Public Const FLOODFILLBORDER = 0   'Erases everything (I think?) graphic-wise on
                                                        'the object.

Public Const FLOODFILLSURFACE = 1  'Used for filling shapes.

' SetRect
' An API call for quickly setting the coordinates of a RECT variable.
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


' OleTranslateColor
' I'm new to this API call too. Apparently, if you select a system colour, this API call will
' help you to turn it into a long (recognisable) value and avoids causing an error. You'll
' see.
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
    ByVal lHPalette As Long, lColorRef As Long) As Long

'The RECT type is very common amongst API calls, and is used to define a rectangle.
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'A special type for the shape of the button.
Type Octagon
    TopLeft As RECT
    LeftMiddle As RECT
    TopRight As RECT
    BottomLeft As RECT
    RightMiddle As RECT
    BottomRight As RECT
    Body As RECT
    Text As RECT
    Icon As RECT
End Type

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Const Author = "Andrew (aka The Bad One)"

'DrawEdge Constants
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

