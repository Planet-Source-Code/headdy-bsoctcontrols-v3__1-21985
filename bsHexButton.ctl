VERSION 5.00
Begin VB.UserControl bsOctButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   MaskColor       =   &H00808000&
   PropertyPages   =   "bsHexButton.ctx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   4440
   ToolboxBitmap   =   "bsHexButton.ctx":0048
   Begin VB.PictureBox picFace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1560
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "bsOctButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' *************************************
' bsOctButton control

' started 22/3/2001
' finished 22/3/2001
'
' ©2000-2001 BadSoft, all rights reserved.
' *************************************

' Someone wanted an 8-sided version of the HexControls and
' lacked the experience to produce one themselves, so they
' called forth upon the creator (ie. ME!) to help them out.
' I'd do the same thing in that person's position, except
' they'd NEVER help me out. The world's greatest Visual Basic
' coder, at the end of the day, is nothing but scum if they
' don't help other people out.
' Anyways it's always another control to add to my growing
' collection, and as said before somewhere in the code it
' should be easy enough to make a six sided shape eight sided.
' Just don't ask me about odd numbered sides!

' *************************************
' HISTORY

' 22/3/2001 - A cry for help is received after posting my code
' in PSC (which came in at a dismal 10th place in the Code of
' the Day listings, but I ain't worried because the only people
' who come in first are those who post updates to already
' submitted code and then brag about how they won awards eight
' times in a row). They want me to develop an 8-sided control
' for use in their program, with acknowledgement if I am
' successful.
' Thus, the bsOctControls are born.

' 21/3/2001 - bsHexControls v2 is released.

' 21/10/2000 - Second update is ready.

' FIXED: The background of the bsHexButton can be any colour except for transparent.
'            Can anybody make the button have a transparent background? WITHOUT
'            using an Image or Picture control?

' 20/10/2000 - Half term time! Haven't checked e-mail so I don't know how well the
' bsHexButton has done. Top 10 in the Code of the Day rankings?

' FIXED: bsHexButton's icon previously could only be in ICO format, but a nicked
'            set of functions from Ariad's Toolbar control allowed me to handle bitmaps as
'            well. No support for other formats (metafiles and enhanced metafiles) is
'            intended.
'            Unfortunately, try as I might, I couldn't get the bitmap pictures to be drawn
'            disabled properly.

' FIXED: bsHexButton no longer responds if you click outside of the hexagon. I had to
'            make a little program to calculate the areas where clicking can take place.

' ADDED: bsHexButton now has the popup thing; if the mouse moves off while the
'              mouse button is held down, it pops back up again.

' 19/10/2000 - BadSoft comes back to PSC with its first submission in months! Yes I'm
' talking about the bsHexButton. Within 10 minutes 10 people had seen the code, and I
' was proud to see my submission displayed in the ticker. I was a bit worried about my
' comments though...
' *************************************

Option Explicit

Dim MyButton As Octagon
Dim PE As New ascPaintEffects   'from a class borrowed from Ariad. Hope they don't
                                               'mind! But they will, so here it is!

'DEFAULT PROPERTY VALUES
Const m_def_Margin = 2
'The button's text defaults to a normal command button's text colour:
'Const m_def_TextColour = vbButtonText
'Disabled text defaults to the system's disabled text colour:
Const m_def_DisabledTextColour = vbGrayText
'While the default button colour is the same as the others (buttons).
Const m_def_Colour = vbButtonFace
'Const m_def_BackColour = vbButtonFace
Const m_def_MaskColour = 0

'PROPERTY VARIABLES
Dim m_Margin As Integer
'Dim m_Text As String
Dim m_Icon As Picture
Dim m_Colour As OLE_COLOR
'Dim m_TextColour As OLE_COLOR
Dim m_DisabledTextColour As OLE_COLOR
'Dim m_BackColour As OLE_COLOR
Dim m_MaskColour As OLE_COLOR

'EVENTS
Event Click(ByVal Button As Integer)
'Default Property Values:
Const m_def_FlatBorderColour = 0
Const m_def_IconAlign = 0
Const m_def_CaptionColour = vbButtonText
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_Alignment = 0
Const m_def_BorderStyle = 2
Const m_def_BackType = 0
Const m_def_IsTruncated = 0
'Property Variables:
Dim m_FlatBorderColour As OLE_COLOR
Dim m_IconAlign As bsIconAlign
Dim m_Caption As String
Dim m_CaptionColour As OLE_COLOR
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_Alignment As bsTextAlign
Dim m_BorderStyle As bsBorderStyle
Dim m_BackType As bsBackType
Dim m_BackPicture As Picture
Dim m_IsTruncated As Boolean

Public Enum bsBackType
    btColour
    btPictureSingle
    btPictureTiled
    btPictureStretched
End Enum

Public Enum bsBorderStyle
    bsFlat
    bsRaisedThin
    bsRaised3D
    bsEtched
End Enum

Public Enum bsIconAlign
   bsILeft
   bsITop
   bsIRight
   bsIBottom
End Enum

Public Enum bsTextAlign
   bsAlignLeft
   bsAlignCentre
   bsAlignRight
End Enum

' DrawButton
' -----------------
' This is going to be a little more difficult than drawing
' your common/garden, PSC-style rectangular button. Don't be
' scared at the size of this procedure; The Bad One explains
' it all.

Private Sub DrawButton(Pressed As Boolean)

   'All the variables needed are dimensioned here.
   Dim Style As Long, x As Long, y As Long, hIcon As Long
   Dim oldString As String
   Dim IconGap As Integer, A As Integer, B As Integer
   Dim Gap As Single, realX As Single, realY As Single
   
   'First we set the most crucial properties and variables to
   'this control. For those who don't know I prefer to work in
   'pixels rather than MickeySoft's very useless and stupid
   'twips.
   
   Style = IIf(Pressed, BDR_SUNKEN, BDR_RAISED)
   ScaleMode = vbPixels
   AutoRedraw = True
   FillStyle = vbFSSolid
   FillColor = m_Colour
   ForeColor = m_Colour
   Picture = Nothing
   Cls
   Gap = UserControl.ScaleHeight / 3
    
    
   'SetRect is called again but 8 times instead of 6. Start
   'from the top left, then go down and right. (A little like
   'with a woman...)
    
   With MyButton
      'Top Left
      Call SetRect(.TopLeft, 0, 0, Gap, Gap)
      'Middle Left (I really should have renamed all of the others)
      Call SetRect(.LeftMiddle, 0, Gap, Gap, Gap * 2)
      'Bottom Left
      Call SetRect(.BottomLeft, 0, Gap * 2, Gap, UserControl.ScaleHeight)
      'Top Right
      Call SetRect(.TopRight, UserControl.ScaleWidth - Gap, 0, _
          UserControl.ScaleWidth, Gap)
      'Middle Right
      Call SetRect(.RightMiddle, UserControl.ScaleWidth - Gap, Gap, _
          UserControl.ScaleWidth, Gap * 2)
      'Bottom Right
      Call SetRect(.BottomRight, UserControl.ScaleWidth - Gap, Gap * 2, _
          UserControl.ScaleWidth, UserControl.ScaleHeight)
      'Finally one for the rest.
      Call SetRect(.Body, Gap, 0, UserControl.ScaleWidth - Gap, _
          UserControl.ScaleHeight)
      .Text = .Body 'The text area is initially set to the same as the body.
        
Icon:
      'This is the error-less way of discovering whether or not
      'a Picture object contains a picture (so obvious too):
             
      If Not m_Icon Is Nothing Then
      
         'Here's something you don't know. When loading a picture
         'into a Picture control or StdPicture, Visual Basic uses
         'HiMetric measurements to measure the height and width.
         'So we need to convert these to The Bad One's preferred
         'unit of measurement, the pixel.
         
         realX = UserControl.ScaleX(m_Icon.Width, vbHimetric, _
            vbPixels)
         realY = UserControl.ScaleY(m_Icon.Height, vbHimetric, _
            vbPixels)
         
         If m_IconAlign = bsILeft Or m_IconAlign = bsIRight Then
            IconGap = (UserControl.ScaleHeight - realY) / 2
         Else
            IconGap = (UserControl.ScaleWidth - realX) / 2
         End If
         
         Select Case m_IconAlign
         Case bsILeft
            .Icon.Left = Gap + IIf(Pressed, 1, 0)
            .Icon.Top = IconGap + IIf(Pressed, 1, 0)
            .Icon.Right = Gap + realX + IIf(Pressed, 1, 0)
            .Icon.Bottom = IconGap + realY + IIf(Pressed, 1, 0)
            .Text.Left = .Icon.Right + m_Margin
         Case bsIRight
            .Icon.Left = ScaleWidth - Gap - realX + IIf(Pressed, 1, 0)
            .Icon.Top = IconGap + IIf(Pressed, 1, 0)
            .Icon.Right = .Icon.Left + realX + IIf(Pressed, 1, 0)
            .Icon.Bottom = IconGap + realY + IIf(Pressed, 1, 0)
            .Text.Right = .Icon.Left - m_Margin
         Case bsITop
            'Tricky...
            .Icon.Left = IconGap + IIf(Pressed, 1, 0)
            .Icon.Top = Gap / 2 + IIf(Pressed, 1, 0)
            .Icon.Right = IconGap + realX + IIf(Pressed, 1, 0)
            .Icon.Bottom = .Icon.Top + realY
            .Text.Top = .Text.Top + m_Margin
         Case bsIBottom
            .Icon.Left = IconGap + IIf(Pressed, 1, 0)
            .Icon.Bottom = ScaleHeight - Gap / 2 + IIf(Pressed, 1, 0)
            .Icon.Top = .Icon.Bottom - realY
            .Icon.Right = IconGap + realX + IIf(Pressed, 1, 0)
            .Text.Top = .Text.Top - m_Margin
         End Select
      End If

      'To make sure graphics are drawn correctly, the
      'background of the control is drawn first.

      Select Case m_BackType
         Case btColour
            UserControl.BackColor = m_Colour
         Case btPictureSingle
            Picture = m_BackPicture
         Case btPictureStretched
            'Bloody took me ages!
            picBG.Picture = m_BackPicture
            StretchBlt UserControl.hdc, 0, 0, ScaleWidth, _
               ScaleHeight, picBG.hdc, 0, 0, picBG.ScaleWidth, _
               picBG.ScaleHeight, vbSrcCopy
         Case btPictureTiled
            picBG.Picture = m_BackPicture
            For B = 0 To ScaleHeight Step picBG.ScaleHeight
               For A = 0 To ScaleWidth Step picBG.ScaleWidth
                  BitBlt UserControl.hdc, A, B, ScaleWidth, _
                     ScaleHeight, picBG.hdc, 0, 0, vbSrcCopy
               Next
            Next
      End Select
      picFace.Move 0, 0, ScaleWidth, ScaleHeight

      'This is the hardest part; we've got to take our picture
      'and draw it on the button. We do this before everything
      'else so that the border is drawn over the picture in the
      'case that the picture is bigger than the button. There's
      'still the problem of making bitmap images disabled.

      If Not m_Icon Is Nothing Then
         Select Case m_Icon.Type
            Case vbPicTypeBitmap
               'Bitmaps are slightly harder to handle because they are drawn with no
               'transparency with the DrawState function. This is where Ariad's class
               'comes into play; all we have to do is specify a transparent colour, which is
               'the MaskColour.
               If Enabled Then
                   PE.PaintTransparentPicture UserControl.hdc, m_Icon, .Icon.Left, _
                       .Icon.Top, realX, realY, 0, 0, m_MaskColour
               Else
                  ' --- This is the closest I've got
                  PE.PaintDisabledPictureEx UserControl.hdc, .Icon.Left, .Icon.Top, _
                      realX, realY, m_Icon, 0, 0, m_MaskColour
               End If
   
            Case vbPicTypeIcon
               'Life is considerably easier if the icon submitted is in ICO format; all we
               'have to do is use the DrawState function for everything.
               hIcon = m_Icon.handle
               Call DrawState(UserControl.hdc, 0, 0, hIcon, 0, _
                  .Icon.Left, .Icon.Top, realX, realY, _
                  DST_ICON Or IIf(Enabled, DSS_NORMAL, _
                  DSS_DISABLED))
         End Select
      End If
        
      'The edges of the panel are drawn here.
      If Pressed Then
         DrawPressedEdges
      Else
         DrawEdges
      End If

      'That's the shape of the button drawn, but now for the
      'hard part - drawing the text!
      'Thanks goes out to a good Visual Basic site - I can't
      'remember the name (I think it's VB Explorer) - for
      'showing me how to use this.

      'This checks the text to see if it is too long for the
      'button, and modifies the string by truncating words
      'and placing an ellipsis (...) at the end if it is.
      'Now why can't all these "OH SO COOL, BUTTON WITH _
      'ICON!" PSC buttons be like this?
      'Using the DT_CALCRECT flag doesn't draw the text onto
      'the control, it just "autosizes" the rect to fit the
      'text exactly.
        
      ForeColor = IIf(Enabled, m_CaptionColour, m_DisabledTextColour)
      oldString = m_Caption

      'Shift the text a little bit if the button is being
      'pressed.
      With .Text
         Gap = (ScaleHeight - (.Bottom - .Top)) \ 2
         .Top = .Top + Gap + IIf(Pressed, 1, 0)
         .Left = .Left + IIf(Pressed, 1, 0)
         .Bottom = .Bottom + Gap + IIf(Pressed, 1, 0)
         .Right = .Right + IIf(Pressed, 1, 0)
         Call DrawText(UserControl.hdc, oldString, -1, MyButton.Text, _
          DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS _
          Or DT_VCENTER Or m_Alignment)
      End With
         
      'Is the modified string the same as the Text property?
         m_IsTruncated = (oldString <> m_Caption)
      End With

      'Draw the masked control (the hard way)
      With MyButton
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
    
    
      'Microsoft forgot to tell you (and I should sue them for
      'causing hours of stress) that you need to set the
      'MaskPicture property from a Picture control using the
      'IMAGE property (when AutoRedraw is True), not the picture.
   
      Set MaskPicture = picFace.Image
      MaskColor = &H808000
      UserControl.BackStyle = 0
   
      'Remember to include this line!
      UserControl.Refresh

End Sub


Private Sub DrawPressedEdges()
   'This draw the edges of the control in the selected style
   'and colours WHEN THE BUTTON HAS BEEN PRESSED.
   
   With MyButton
      If m_BorderStyle = bsEtched Then
         ForeColor = m_HighlightColour
      Else
         ForeColor = IIf(m_BorderStyle = bsFlat, _
            m_FlatBorderColour, m_ShadowDKColour)
      End If
      Line (.TopLeft.Left, .TopLeft.Bottom - 1)-(.TopLeft.Right - 1, .TopLeft.Top)
      Line (.LeftMiddle.Left, .LeftMiddle.Top)-(.LeftMiddle.Left, .LeftMiddle.Bottom)
      Line (.BottomLeft.Left, .BottomLeft.Top)-(.BottomLeft.Right, .BottomLeft.Bottom)
      Line (.TopLeft.Right - 1, .TopLeft.Top)-(.TopRight.Left, .TopRight.Top)
      
      If m_BorderStyle = bsEtched Then
         ForeColor = m_ShadowDKColour
      Else
         ForeColor = IIf(m_BorderStyle = bsFlat, _
            m_FlatBorderColour, m_HighlightColour)
      End If
      Line (.TopRight.Left, .TopRight.Top)-(.TopRight.Right, .TopRight.Bottom)
      Line (.BottomRight.Left - 1, .BottomRight.Bottom)-(.BottomRight.Right, .BottomRight.Top - 1)
      Line (.BottomLeft.Right, .BottomLeft.Bottom - 1)-(.BottomRight.Left, .BottomRight.Bottom - 1)
      Line (.BottomRight.Right - 1, .BottomRight.Top)-(.TopRight.Right - 1, .TopRight.Bottom - 1)
      
      If m_BorderStyle = bsFlat Or _
         m_BorderStyle = bsRaisedThin Then
         Exit Sub
      End If
      
      ForeColor = m_ShadowColour
      Line (.TopLeft.Left + 1, .TopLeft.Bottom - 1)-(.TopLeft.Right - 1, .TopLeft.Top + 1)
      Line (.LeftMiddle.Left + 1, .LeftMiddle.Top)-(.LeftMiddle.Left + 1, .LeftMiddle.Bottom)
      Line (.BottomLeft.Left + 1, .BottomLeft.Top)-(.BottomLeft.Right, .BottomLeft.Bottom - 1)
      Line (.TopLeft.Right - 1, .TopLeft.Top + 1)-(.TopRight.Left, .TopRight.Top + 1)
      ForeColor = m_HighlightDKColour
      Line (.TopRight.Left, .TopRight.Top + 1)-(.TopRight.Right - 1, .TopRight.Bottom)
      Line (.BottomRight.Left, .BottomRight.Bottom - 2)-(.BottomRight.Right - 1, .BottomRight.Top - 1)
      Line (.BottomLeft.Right, .BottomLeft.Bottom - 2)-(.BottomRight.Left, .BottomRight.Bottom - 2)
      Line (.BottomRight.Right - 2, .BottomRight.Top)-(.TopRight.Right - 2, .TopRight.Bottom - 1)
   End With
End Sub


Private Sub DrawEdges()
   'This draw the edges of the control in the selected style
   'and colours.
   
   With MyButton
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


' UserControl_MouseDown
' -----------------
' Something really scary happened. If I didn't add the check for the current mode
' (Ambient.Usermode is True if the control is being manipulated in the IDE, False if the
' program is running), the button would click! This exlains how the Tabbed Dialog control
' is able to work.

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Ambient.UserMode Then
      DrawButton True
   End If
End Sub


' UserControl_MouseMove
' -----------------
' A private sub to seperate the BadSoft crew from the MickeySoft crew. Sorry about
' these plugs...
' Anyway, it's not fair that as soon as the user clicks on the button they activate the
' Click event anyway. If the user moves the mouse off the button while the mouse
' button is being held down, our bsOctButton should pop back up again. Agreed?

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, _
   x As Single, y As Single)
   If Ambient.UserMode Then
      If Button And (WithinRect(x, y, MyButton.Body) Or WithinSides(x, y)) Then
         DrawButton True
      Else
         DrawButton False
      End If
   End If
End Sub

' UserControl_MouseUp
' -----------------
' When the mouse button is released, we want to trigger the Click event. What use is a
' button if you can't tell when it's been clicked?

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   DrawButton False 'pop the button back up.
   If Ambient.UserMode Then
      If Button And (WithinRect(x, y, MyButton.Body) Or _
         WithinSides(x, y)) Then
         RaiseEvent Click(Button)
      End If
   End If
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
    DrawButton False
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

    With MyButton
        'Check the top left corner. If you don't know, we're dividing the odd-shaped sides
        'into triangles with a right angle in the corner, which is the reason for the
        'Hexagon type having five RECT variables - FOOLPROOF.
        With .TopLeft
            If WithinRect(x, y, MyButton.TopLeft) And x >= .Bottom - y Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The bottom left corner - FOOLPROOF.
        With .BottomLeft
            If x >= y - .Top And WithinRect(x, y, MyButton.BottomLeft) Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The top right corner - FOOLPROOF.
        With .TopRight
            If (x - .Left) <= y And WithinRect(x, y, MyButton.TopRight) Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'The bottom right corner! FOOLPROOF!
        With .BottomRight
            If WithinRect(x, y, MyButton.BottomRight) And x - .Left <= .Bottom - y Then
                WithinSides = True
                Exit Function
            End If
        End With
        
        'Middles! (You know the words.)
        If WithinRect(x, y, .RightMiddle) Then
            WithinSides = True
            Exit Function
        End If
        If WithinRect(x, y, .LeftMiddle) Then
            WithinSides = True
            Exit Function
        End If
    End With
End Function


' Colour
' -----------------
' The colour fo the face of the button. Unfortunately, regardless of which colour you
' use, the borders will be the same. We could try to change that in a next release(?)
' By making Colour an OLE_COLOR, we can choose the colour either from the
' property pages or from the property window.

Public Property Get Colour() As OLE_COLOR
Attribute Colour.VB_Description = "The colour of the face of the button if BackStyle is set to bsColour."
Attribute Colour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Colour = m_Colour
End Property

Public Property Let Colour(ByVal New_Colour As OLE_COLOR)
    OleTranslateColor New_Colour, 0, m_Colour
    PropertyChanged "Colour"
    'The following line goes into all of the properties so that the control is updated.
    DrawButton False
End Property


' UserControl_InitProperties
' -----------------
' Here's where default properties for the control are set.

Private Sub UserControl_InitProperties()
   m_Colour = vbButtonFace
   m_DisabledTextColour = m_def_DisabledTextColour
   m_Margin = m_def_Margin
   m_MaskColour = m_def_MaskColour
   Set m_Icon = LoadPicture("")    'no icon
   Set Font = Ambient.Font         'possibly the font of the form
   m_IsTruncated = m_def_IsTruncated
   m_Caption = UserControl.Name
   m_CaptionColour = m_def_CaptionColour
   m_HighlightColour = m_def_HighlightColour
   m_HighlightDKColour = m_def_HighlightDKColour
   m_ShadowColour = m_def_ShadowColour
   m_ShadowDKColour = m_def_ShadowDKColour
   m_Alignment = m_def_Alignment
   m_BorderStyle = m_def_BorderStyle
   m_BackType = m_def_BackType
   Set m_BackPicture = LoadPicture("")
   m_IconAlign = m_def_IconAlign
   m_FlatBorderColour = m_def_FlatBorderColour
End Sub


' UserControl_ReadProperties
' -----------------
' Where properties for the control are read. If the properties cannot be read, the
' defaults are used.

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Colour = PropBag.ReadProperty("Colour", m_def_Colour)
'    m_TextColour = PropBag.ReadProperty("TextColour", m_def_TextColour)
    m_DisabledTextColour = PropBag.ReadProperty("DisabledTextColour", m_def_DisabledTextColour)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Margin = PropBag.ReadProperty("Margin", m_def_Margin)
    m_MaskColour = PropBag.ReadProperty("MaskColour", m_def_MaskColour)
    m_IsTruncated = PropBag.ReadProperty("IsTruncated", m_def_IsTruncated)
'    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Name)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDKColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDKColour", m_def_ShadowDKColour)
   m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_BackType = PropBag.ReadProperty("BackType", m_def_BackType)
   Set m_BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
   m_IconAlign = PropBag.ReadProperty("IconAlign", m_def_IconAlign)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
End Sub


' UserControl_Show
' -----------------
' Except when Height > Width, the control's appearance will screw up when the form is
' closed and then reopened, or on first adding it to the form. This procedure was added
' to prevent that from happening.

Private Sub UserControl_Show()
    DrawButton False
End Sub


' UserControl_WriteProperties
' -----------------
' Where the control's properties are saved. This only occurs when the form containing
' the control has been saved in the IDE. Again, if the properties cannot be found, the
' defaults are used.

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Colour", m_Colour, m_def_Colour)
   Call PropBag.WriteProperty("DisabledTextColour", m_DisabledTextColour, m_def_DisabledTextColour)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
   Call PropBag.WriteProperty("Font", Font, Ambient.Font)
   Call PropBag.WriteProperty("Margin", m_Margin, m_def_Margin)
   Call PropBag.WriteProperty("MaskColour", m_MaskColour, m_def_MaskColour)
   Call PropBag.WriteProperty("IsTruncated", m_IsTruncated, m_def_IsTruncated)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Name)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDKColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDKColour", m_ShadowDKColour, m_def_ShadowDKColour)
   Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("BackType", m_BackType, m_def_BackType)
   Call PropBag.WriteProperty("BackPicture", m_BackPicture, Nothing)
   Call PropBag.WriteProperty("IconAlign", m_IconAlign, m_def_IconAlign)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
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
'    OleTranslateColor New_TextColour, 0, m_TextColour
'    PropertyChanged "TextColour"
'    DrawButton False
'End Property


' DisabledTextColour
' -----------------
' The same as TextColour, except for when the control is disabled.

Public Property Get DisabledTextColour() As OLE_COLOR
Attribute DisabledTextColour.VB_Description = "The colour of the text when the button is disabled."
Attribute DisabledTextColour.VB_ProcData.VB_Invoke_Property = ";Colour"
    DisabledTextColour = m_DisabledTextColour
End Property

Public Property Let DisabledTextColour(ByVal New_DisabledTextColour As OLE_COLOR)
    OleTranslateColor New_DisabledTextColour, 0, m_DisabledTextColour
    PropertyChanged "DisabledTextColour"
    DrawButton False
End Property

' Enabled
' -----------------
' The enabled status of the control. Since it's mapped to the UserControl's Enabled
' property, maybe you should do what MickeySoft asks below:

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "The enabled state of the control."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behaviour"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawButton False
End Property


' Icon
' -----------------
' You can include an icon to go with the text on the button. Now I know some prick is
' going to try and put their 800 x 600 pixel photo in there, so I say go ahead. Be my
' guest. Just don't expect it to work.
' At present the icon can only be left aligned with the text. The icon can also be in ICO
' or BMP format (for bitmaps you will have to set the MaskColour property to the
' transparent colour of the picture). No support for metafiles; if you try, it will be
' rejected.

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "An optional icon to appear on the control."
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Picture"
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)

    Set m_Icon = New_Icon
    
    If New_Icon.Type = vbPicTypeMetafile Or New_Icon.Type = vbPicTypeEMetafile _
        Then
        MsgBox "Metafiles are not supported."
        Set m_Icon = LoadPicture 'clear the image
    End If
    
    PropertyChanged "Icon"
    DrawButton False
End Property


' Font
' -----------------
' If you like, you can change the font to something a little more classy than MS Sans
' Serif. But once again you are asked...

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font

Public Property Get Font() As Font
Attribute Font.VB_Description = "Sets the font of the text for the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawButton False
End Property


' Margin
' -----------------
' This sets the space between the text and the icon, if there is an icon. Some icons
' look better with different margin sizes.

Public Property Get Margin() As Integer
Attribute Margin.VB_Description = "The space (in pixels) between the icon and the text of the button."
Attribute Margin.VB_ProcData.VB_Invoke_Property = ";Picture"
    Margin = m_Margin
End Property

Public Property Let Margin(ByVal New_Margin As Integer)
    m_Margin = Val(New_Margin)
    PropertyChanged "Margin"
    DrawButton False
End Property


' MaskColour
' -----------------
' To be able to use bitmaps as the icon, you need to specify which colour is
' transparent in the bitmap. Unfortunately I can't even guess for you right now. This is
' needed to make your bitmap show nicely in the button; it needs to work whatever
' colour the user's 3D controls are. But you can probably get away with it by using the
' Colour property!

Public Property Get MaskColour() As OLE_COLOR
    MaskColour = m_MaskColour
End Property

Public Property Let MaskColour(ByVal New_MaskColour As OLE_COLOR)
Attribute MaskColour.VB_Description = "If the image is a bitmap, this is the colour that appears as transparent."
Attribute MaskColour.VB_ProcData.VB_Invoke_PropertyPut = ";Colour"
    OleTranslateColor New_MaskColour, 0, m_MaskColour
    PropertyChanged "MaskColour"
    DrawButton False
End Property


' IsTruncated
' -----------------
' Sometimes the user may need to know whether or not the Text property fits the button.

Public Property Get IsTruncated() As Boolean
Attribute IsTruncated.VB_Description = "Specifies whether the Caption is too long for the control."
Attribute IsTruncated.VB_MemberFlags = "400"
    IsTruncated = m_IsTruncated
End Property
'Public Property Get Caption() As Variant
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As Variant)
'    m_Caption = New_Caption
'    PropertyChanged "Caption"
'    DrawButton False
'End Property


' ShowAbout
' -----------------
' Shows the About form.

Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows the About screen."
Attribute ShowAbout.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text displayed on the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_MemberFlags = "200"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawButton False
End Property

Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the text displyed on the button."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawButton False
End Property

Public Property Get HighlightColour() As OLE_COLOR
Attribute HighlightColour.VB_Description = "The colour of the lightest area of the control's border."
Attribute HighlightColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawButton False
End Property

Public Property Get HighlightDKColour() As OLE_COLOR
Attribute HighlightDKColour.VB_Description = "The colour of the second lightest area of the control's border."
Attribute HighlightDKColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDKColour"
   DrawButton False
End Property

Public Property Get ShadowColour() As OLE_COLOR
Attribute ShadowColour.VB_Description = "The colour of the second darkest area of the control's border."
Attribute ShadowColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawButton False
End Property

Public Property Get ShadowDKColour() As OLE_COLOR
Attribute ShadowDKColour.VB_Description = "The colour of the darkest area of the control's border."
Attribute ShadowDKColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDKColour"
   DrawButton False
End Property

Public Property Get Alignment() As bsTextAlign
Attribute Alignment.VB_Description = "How the text inside the button is aligned."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
   Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As bsTextAlign)
   m_Alignment = New_Alignment
   PropertyChanged "Alignment"
   DrawButton False
End Property

Public Property Get BorderStyle() As bsBorderStyle
Attribute BorderStyle.VB_Description = "The style of the control's border."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsBorderStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawButton False
End Property

Public Property Get BackType() As bsBackType
Attribute BackType.VB_Description = "Specifies how the BackPicture is used as the background."
Attribute BackType.VB_ProcData.VB_Invoke_Property = ";Picture"
   BackType = m_BackType
End Property

Public Property Let BackType(ByVal New_BackType As bsBackType)
   m_BackType = New_BackType
   PropertyChanged "BackType"
   DrawButton False
End Property

Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "The picture that serves as the background of the control if BadkStyle is not bsColour."
Attribute BackPicture.VB_ProcData.VB_Invoke_Property = ";Picture"
   Set BackPicture = m_BackPicture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
   Set m_BackPicture = New_BackPicture
   PropertyChanged "BackPicture"
   DrawButton False
End Property

Public Property Get IconAlign() As bsIconAlign
Attribute IconAlign.VB_Description = "How the icon is aligned inside the control."
Attribute IconAlign.VB_ProcData.VB_Invoke_Property = ";Picture"
   IconAlign = m_IconAlign
End Property

Public Property Let IconAlign(ByVal New_IconAlign As bsIconAlign)
   m_IconAlign = New_IconAlign
   PropertyChanged "IconAlign"
   DrawButton False
End Property

Public Property Get FlatBorderColour() As OLE_COLOR
Attribute FlatBorderColour.VB_Description = "If BorderStyle is set to bsFlat, this is the colour of the border."
Attribute FlatBorderColour.VB_ProcData.VB_Invoke_Property = ";Colour"
   FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
   m_FlatBorderColour = New_FlatBorderColour
   PropertyChanged "FlatBorderColour"
End Property

