VERSION 5.00
Begin VB.UserControl bsOctPanel 
   Alignable       =   -1  'True
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   PropertyPages   =   "bsHexPanel.ctx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   5280
   ToolboxBitmap   =   "bsHexPanel.ctx":003E
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   960
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picFace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2520
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "bsOctPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' *************************************
' bsOctPanel control

' started 22/3/2001
' finished 22/3/2001
'
' ©2000-2001 BadSoft, all rights reserved.
' *************************************

' Developed especially for those people who want to create games in the style of
' "Who Wants to Be a Millionaire". This is a panel to go with the bsOctButton control,
' and allows the text inside it to be left aligned, centered or right aligned. It's basically
' a cut down version of the bsHexButton.

' If you like this control, don't be afraid to drop me a line at badhart@hotpop.com. As
' all BadSoft's creations are, this is Freeware and can be used for just about any
' purpose. Acknowledgement is always appreciated.

' *************************************
' HISTORY

' See bsOctButton.
' *************************************

Option Explicit

Dim MyPanel As Octagon

'DEFAULT PROPERTY VALUES
'The button's text defaults to a normal command button's text colour:
'Const m_def_TextColour = vbButtonText
Const m_def_Colour = vbButtonFace
Const m_def_Alignment = vbCenter

'PROPERTY VARIABLES
'Dim m_Text As String
Dim m_Colour As OLE_COLOR
'Dim m_TextColour As OLE_COLOR
Dim m_Alignment As Byte

'EVENTS
Event Click(ByVal Button As Integer)
'Default Property Values:
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_CaptionColour = 0
Const m_def_FlatBorderColour = 0
Const m_def_BorderStyle = 5
Const m_def_BackType = 0
'Property Variables:
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_CaptionColour As OLE_COLOR
Dim m_Caption As String
Dim m_FlatBorderColour As OLE_COLOR
Dim m_BorderStyle As bsBorderStyle
Dim m_BackType As bsBackType
Dim m_BackPicture As Picture

Private Sub DrawEdges()
   'This draw the edges of the control in the selected style
   'and colours.
   
   With MyPanel
      If m_BorderStyle = bsEtched Then
         ForeColor = m_ShadowDKColour
      Else
         ForeColor = IIf(m_BorderStyle = bsFlat, _
            m_FlatBorderColour, m_HighlightColour)
      End If
      Line (.TopLeft.Left, .TopLeft.Bottom - 1)-(.TopLeft.Right - 1, .TopLeft.Top)
      Line (.LeftMiddle.Left, .LeftMiddle.Top)-(.LeftMiddle.Left, .LeftMiddle.Bottom)
      Line (.BottomLeft.Left, .BottomLeft.Top)-(.BottomLeft.Right, .BottomLeft.Bottom)
      Line (.TopLeft.Right - 1, .TopLeft.Top)-(.TopRight.Left, .TopRight.Top)
      
      If m_BorderStyle = bsEtched Then
         ForeColor = m_HighlightColour
      Else
         ForeColor = IIf(m_BorderStyle = bsFlat, _
            m_FlatBorderColour, m_ShadowDKColour)
      End If
      Line (.TopRight.Left, .TopRight.Top)-(.TopRight.Right, .TopRight.Bottom)
      Line (.BottomRight.Left - 1, .BottomRight.Bottom)-(.BottomRight.Right, .BottomRight.Top - 1)
      Line (.BottomLeft.Right, .BottomLeft.Bottom - 1)-(.BottomRight.Left, .BottomRight.Bottom - 1)
      Line (.BottomRight.Right - 1, .BottomRight.Top)-(.TopRight.Right - 1, .TopRight.Bottom - 1)
      
      If m_BorderStyle = bsFlat Or _
         m_BorderStyle = bsRaisedThin Then
         Exit Sub
      End If
      
      ForeColor = m_HighlightDKColour
      Line (.TopLeft.Left + 1, .TopLeft.Bottom - 1)-(.TopLeft.Right - 1, .TopLeft.Top + 1)
      Line (.LeftMiddle.Left + 1, .LeftMiddle.Top)-(.LeftMiddle.Left + 1, .LeftMiddle.Bottom)
      Line (.BottomLeft.Left + 1, .BottomLeft.Top)-(.BottomLeft.Right, .BottomLeft.Bottom - 1)
      Line (.TopLeft.Right - 1, .TopLeft.Top + 1)-(.TopRight.Left, .TopRight.Top + 1)
      ForeColor = m_ShadowColour
      Line (.TopRight.Left, .TopRight.Top + 1)-(.TopRight.Right - 1, .TopRight.Bottom)
      Line (.BottomRight.Left, .BottomRight.Bottom - 2)-(.BottomRight.Right - 1, .BottomRight.Top - 1)
      Line (.BottomLeft.Right, .BottomLeft.Bottom - 2)-(.BottomRight.Left, .BottomRight.Bottom - 2)
      Line (.BottomRight.Right - 2, .BottomRight.Top)-(.TopRight.Right - 2, .TopRight.Bottom - 1)
   End With
End Sub

' DrawPanel
' -----------------

Private Sub DrawPanel()
   Dim Gap As Single
   Dim x As Long, y As Long, newHeight As Long
   Dim oldString As String
   Dim AlignStyle As Integer, PicX As Integer, PicY As Integer
   Dim A As Integer, B As Integer
      
   'The following are set before any drawing takes place.
   'MickeySoft thinks that FillStyle should be transparent as
   'the default - idiots.
   ScaleMode = vbPixels
   AutoRedraw = True
   FillStyle = vbFSSolid
   FillColor = m_Colour
   ForeColor = m_Colour
   Picture = Nothing
   Cls
   Gap = UserControl.ScaleHeight / 3
   
   'All of the RECT variables need to be set, but fortunately
   'we can use a very useful API call. We go from top left to
   'bottom right.
   With MyPanel
   Call SetRect(.TopLeft, 0, 0, Gap, Gap)
   Call SetRect(.LeftMiddle, 0, Gap, Gap, Gap * 2)
   Call SetRect(.BottomLeft, 0, Gap * 2, Gap, _
       UserControl.ScaleHeight)
   Call SetRect(.TopRight, UserControl.ScaleWidth - Gap, _
       0, UserControl.ScaleWidth, Gap)
   Call SetRect(.RightMiddle, _
       UserControl.ScaleWidth - Gap, Gap, _
       UserControl.ScaleWidth, Gap * 2)
   Call SetRect(.BottomRight, _
       UserControl.ScaleWidth - Gap, Gap * 2, _
       UserControl.ScaleWidth, UserControl.ScaleHeight)
   'The following one is for the rest of the area, where
   'we need a top and bottom border only.
   Call SetRect(.Body, Gap, 0, UserControl.ScaleWidth - Gap, _
       UserControl.ScaleHeight)
   'The text area is initially set to the same as the body.
   .Text = .Body
   End With
   
   'Then to do the background of the panel. Remember it now
   'supports a bitmap background! This has been moved so that,
   'regardless of what the background is, the borders are drawn
   'on top.
   Select Case m_BackType
      Case btColour
         UserControl.BackColor = m_Colour
      Case btPictureSingle
         Picture = m_BackPicture
      Case btPictureStretched
         'Bloody took me ages!
         picBG.Picture = m_BackPicture
         StretchBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, _
            picBG.hdc, 0, 0, picBG.ScaleWidth, picBG.ScaleHeight, vbSrcCopy
      Case btPictureTiled
         picBG.Picture = m_BackPicture
         For B = 0 To ScaleHeight Step picBG.ScaleHeight
            For A = 0 To ScaleWidth Step picBG.ScaleWidth
               BitBlt UserControl.hdc, A, B, ScaleWidth, ScaleHeight, _
                  picBG.hdc, 0, 0, vbSrcCopy
            Next
         Next
   End Select
   picFace.Move 0, 0, ScaleWidth, ScaleHeight

   
   'The edges of the panel are drawn here.
   DrawEdges

      
   'Drawing the text is not so easy; we need to allow for more
   'than one line.
   'We can't truncate text in this control at the moment; if
   'there is too much text, the user will have to resize the
   'control. But we can center more than one line of text
   'vertically; we need to fetch the returned value from the
   'DrawText API call.

   'Doing this test doesn't draw the text onto the control.
   oldString = m_Caption
   newHeight = DrawText(UserControl.hdc, oldString, -1, _
      MyPanel.Text, DT_CALCRECT Or DT_WORDBREAK)

   'Change the height of the text rectangle to position it
   'vertically on the button.
   Gap = (UserControl.ScaleHeight - newHeight) / 2
   With MyPanel.Text
      .Top = Gap
      .Bottom = Gap + newHeight
      'To center the text correctly, the width of the rectangle must be reset to the
      'width of the body (because the DrawText test modifies the width of the
      'rectangle as well as the height). We could also have centered the rectangle
      'horizontally.
      .Right = MyPanel.Body.Right
   
      'Alignment style:
      Select Case m_Alignment
          Case vbLeftJustify
              AlignStyle = DT_LEFT
          Case vbRightJustify
              AlignStyle = DT_RIGHT
          Case vbCenter
              AlignStyle = DT_CENTER
      End Select
   
      'Now to draw the text.
      UserControl.ForeColor = m_CaptionColour
      Call DrawText(UserControl.hdc, oldString, -1, MyPanel.Text, _
          DT_WORDBREAK Or AlignStyle)
   End With

   'Draw the masked control:
   With MyPanel
       picFace.Cls
       picFace.ForeColor = vbBlack
       picFace.Line (.TopLeft.Left, .TopLeft.Bottom - 1)-(.TopLeft.Right - 1, .TopLeft.Top)
       picFace.Line (.LeftMiddle.Left, .LeftMiddle.Top)-(.LeftMiddle.Left, .LeftMiddle.Bottom)
       picFace.Line (.BottomLeft.Left, .BottomLeft.Top)-(.BottomLeft.Right, .BottomLeft.Bottom)
       picFace.Line (.TopLeft.Right - 1, .TopLeft.Top)-(.TopRight.Left, .TopRight.Top)
       picFace.Line (.TopRight.Left, .TopRight.Top)-(.TopRight.Right, .TopRight.Bottom)
       picFace.Line (.BottomRight.Left - 1, .BottomRight.Bottom)-(.BottomRight.Right, .BottomRight.Top - 1)
   End With
   Call ExtFloodFill(picFace.hdc, ScaleWidth \ 2, ScaleHeight \ 2, picFace.BackColor, FLOODFILLSURFACE)
   
   Set MaskPicture = picFace.Image
   MaskColor = &H808000
   BackStyle = 0
   
   'Remember to include this line!
   UserControl.Refresh
   Parent.Refresh
   

End Sub

' UserControl_MouseUp
' -----------------
'You know what this does.

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent Click(Button)
End Sub


' UserControl_Resize
' -----------------
' The control shouldn't really be too small, but some prick is going to try and do that
' anyway so I have added a little code to restrict the size.
' Also, I noticed something very disturbing while debugging the control. If the button is
' taller than it is wide, you get a really nasty shaped control. So, to avoid this, the height of
' the control is never more than the width.

Private Sub UserControl_Resize()
    If Height < 255 Then Height = 255
    If Width < 375 Then Width = 375
    If Height > Width Then Height = Width
    DrawPanel
End Sub


' WithinRect
' -----------------
' Any user control that responds to a mouse click, such as a command button, really
' needs a subroutine like this. This finds out if (X,Y) is inside - or on the border of - a
' rectangular area, defined by the whatRect variable.

Private Function WithinRect(ByVal x As Long, ByVal y As Long, whatRect As RECT) _
    As Boolean
    
    With whatRect
        WithinRect = (x >= .Left And x <= .Right) And (y >= .Top And y <= .Bottom)
    End With
End Function

' WithinSides
' -----------------
' The PSC newbie way of making the button would be to just check the whole rectanglular
' area for a mouse click. This is BadSoft's control, thank you. So this nice function will
' check our oddly shaped areas (namely the sides of the hexagon) to see if the mouse
' cursor is in there. Bear in mind, this took some figuring out.

Private Function WithinSides(ByVal x As Single, ByVal y As Single) As Boolean

    With MyPanel
        'Check the top left corner. If you don't know, we're dividing the odd-shaped sides
        'into triangles with a right angle in the corner, which is the reason for the
        'Hexagon type having five RECT variables - FOOLPROOF.
        With .TopLeft
            If WithinRect(x, y, MyPanel.TopLeft) And x >= .Bottom - y Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The bottom left corner - FOOLPROOF.
        With .BottomLeft
            If x >= y - .Top And WithinRect(x, y, MyPanel.BottomLeft) Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The top right corner - FOOLPROOF.
        With .TopRight
            If (x - .Left) <= y And WithinRect(x, y, MyPanel.TopRight) Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The bottom right corner! FOOLPROOF!
        With .BottomRight
            If WithinRect(x, y, MyPanel.BottomRight) And x - .Left <= .Bottom - y Then
                WithinSides = True
                Exit Function
            End If
        End With
        
    End With
End Function


' Colour
' -----------------
' The colour fo the face of the button. Unfortunately, regardless of which colour you
' use, the borders will be the same. We could try to change that in a next release(?)
' By making Colour an OLE_COLOR, we can choose the colour either from the
' property pages or from the property window.

Public Property Get Colour() As OLE_COLOR
Attribute Colour.VB_Description = "The colour of the control's face is BackType is set to btColour."
Attribute Colour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Colour = m_Colour
End Property

Public Property Let Colour(ByVal New_Colour As OLE_COLOR)
    m_Colour = New_Colour
    PropertyChanged "Colour"
    'The following line goes into all of the properties so that the control is updated.
    DrawPanel
End Property


' UserControl_InitProperties
' -----------------
' Here's where default properties for the control are set.

Private Sub UserControl_InitProperties()
   m_Colour = vbButtonFace
   m_Alignment = m_def_Alignment
   m_BackType = m_def_BackType
   m_BorderStyle = m_def_BorderStyle
   m_FlatBorderColour = m_def_FlatBorderColour
   m_Caption = UserControl.Name
   m_CaptionColour = m_def_CaptionColour
   Set Font = Ambient.Font
   Set m_BackPicture = LoadPicture("")
   m_HighlightColour = m_def_HighlightColour
   m_HighlightDKColour = m_def_HighlightDKColour
   m_ShadowColour = m_def_ShadowColour
   m_ShadowDKColour = m_def_ShadowDKColour
End Sub


' UserControl_ReadProperties
' -----------------
' Where properties for the control are read. If the properties cannot be read, the
' defaults are used.

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Colour = PropBag.ReadProperty("Colour", m_def_Colour)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
   m_BackType = PropBag.ReadProperty("BackType", m_def_BackType)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Name)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   Set Font = PropBag.ReadProperty("Font", Ambient.Font)
   Set m_BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDkColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDkColour", m_def_ShadowDKColour)
End Sub


' UserControl_Show
' -----------------
' Except when Height > Width, the control's appearance will screw up when the form is
' closed and then reopened, or on first adding it to the form. This procedure was added
' to prevent that from happening.

Private Sub UserControl_Show()
    DrawPanel
End Sub


' UserControl_WriteProperties
' -----------------
' Where the control's properties are saved. This only occurs when the form containing
' the control has been saved in the IDE. Again, if the properties cannot be found, the
' defaults are used.

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Colour", m_Colour, m_def_Colour)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Font", Font, Ambient.Font)
   Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
   Call PropBag.WriteProperty("BackType", m_BackType, m_def_BackType)
   Call PropBag.WriteProperty("BackPicture", m_BackPicture, Nothing)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Name)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDkColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDkColour", m_ShadowDKColour, m_def_ShadowDKColour)
End Sub


' Limit
' -----------------
' The Bad One's trademark function that I kind of borrowed from the Amiga version of
' Blitz Basic. It was as buggy as hell, but let's hope that the PC version is better. You
' should try the demo version, it's impressive! (Try blitzbasic.com.)
' Anyway, this function returns Low if Value < Low, and High if Value > High.

Private Function Limit(Value, Low, High)
    If Value < Low Then Value = Low
    If Value > High Then Value = High
    Limit = Value
End Function
'
'
'' TextColour
'' -----------------
'' If you can't guess what this is about, I'd call you blonde. Only kidding (?). This sets
'' the colour of the text when the control is enabled.
'
'Public Property Get TextColour() As OLE_COLOR
'    TextColour = m_TextColour
'End Property
'
'Public Property Let TextColour(ByVal New_TextColour As OLE_COLOR)
'    m_TextColour = New_TextColour
'    PropertyChanged "TextColour"
'    DrawPanel
'End Property
''
''
''' Text
''' -----------------
''' Every good button has text.
''
''Public Property Get Text() As String
''    Text = m_Text
''End Property
''
''Public Property Let Text(ByVal New_Text As String)
''    m_Text = New_Text
''    PropertyChanged "Text"
''    DrawPanel
''End Property


' Font
' -----------------
' If you like, you can change the font to something a little more classy than MS Sans
' Serif. But once again you are asked...

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font

Public Property Get Font() As Font
Attribute Font.VB_Description = "The font of the text used on the panel."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawPanel
End Property


' Alignment
' -----------------
' The main reason for creating the bsHexPanel was to align the text in a larger
' hexagon, which you can't (currently?) do with the button. You can choose from left,
' right or centre, but the text will always be centred vertically within the panel.

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "How the text is aligned on the control."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Text"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    DrawPanel
End Property
'
'
'' BackColour
'' -----------------
'' Since it's not yet possible to make the background colour of the control transparent,
'' We have to change the BackColor property of the UserControl.
'
'Public Property Get BackColour() As OLE_COLOR
'    BackColour = m_BackColour
'End Property
'
'Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
'    m_BackColour = New_BackColour
'    PropertyChanged "BackColour"
'    DrawPanel
'End Property


' ShowAbout
' -----------------
' Shows the About form.

Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows information about the control."
Attribute ShowAbout.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub
Public Property Get BackType() As bsBackType
Attribute BackType.VB_Description = "The type of background the control has (a picture or a plain colour)."
Attribute BackType.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackType = m_BackType
End Property

Public Property Let BackType(ByVal New_BackType As bsBackType)
    m_BackType = New_BackType
    PropertyChanged "BackType"
    DrawPanel
End Property

Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "The picture to be used as the background of the control."
Attribute BackPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set BackPicture = m_BackPicture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
    Set m_BackPicture = New_BackPicture
    PropertyChanged "BackPicture"
    DrawPanel
End Property

Public Property Get BorderStyle() As bsBorderStyle
Attribute BorderStyle.VB_Description = "The style of the border of the control."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsBorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    DrawPanel
End Property

Public Property Get FlatBorderColour() As OLE_COLOR
Attribute FlatBorderColour.VB_Description = "The colour of the border if Borderstyle is set to bsFlat."
Attribute FlatBorderColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
    m_FlatBorderColour = New_FlatBorderColour
    PropertyChanged "FlatBorderColour"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text displayed on the panel."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Basics"
Attribute Caption.VB_MemberFlags = "200"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawPanel
End Property

Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the text on the panel."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawPanel
End Property

Public Property Get HighlightColour() As OLE_COLOR
Attribute HighlightColour.VB_Description = "The colour of the lightest part of the border."
Attribute HighlightColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawPanel
End Property

Public Property Get HighlightDKColour() As OLE_COLOR
Attribute HighlightDKColour.VB_Description = "The colour of the second lightest area of the border."
Attribute HighlightDKColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDkColour"
   DrawPanel
End Property

Public Property Get ShadowColour() As OLE_COLOR
Attribute ShadowColour.VB_Description = "The colour of the second darkest area of the border."
Attribute ShadowColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawPanel
End Property

Public Property Get ShadowDKColour() As OLE_COLOR
Attribute ShadowDKColour.VB_Description = "The colour of the darkest area of the border."
Attribute ShadowDKColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDkColour"
   DrawPanel
End Property

