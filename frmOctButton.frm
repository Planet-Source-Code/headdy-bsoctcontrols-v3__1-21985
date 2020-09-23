VERSION 5.00
Object = "*\AbsOctButton.vbp"
Begin VB.Form frmOctButton 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   27
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      Caption         =   "Icons against colour"
      Height          =   2775
      Left            =   2400
      TabIndex        =   22
      Top             =   2400
      Width           =   2295
      Begin prjbsOctControls.bsOctButton bsOctButton12 
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Colour          =   12640511
         Icon            =   "frmOctButton.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Margin          =   8
         MaskColour      =   16777215
         Caption         =   "Bitmap at bottom"
         Alignment       =   1
         BackPicture     =   "frmOctButton.frx":008B
         IconAlign       =   3
      End
      Begin prjbsOctControls.bsOctButton bsOctButton13 
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Colour          =   12648384
         Icon            =   "frmOctButton.frx":0A5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Margin          =   6
         Caption         =   "Icon on top"
         Alignment       =   1
         IconAlign       =   1
      End
      Begin prjbsOctControls.bsOctButton bsOctButton14 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   16761024
         Icon            =   "frmOctButton.frx":0BB6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColour      =   16777215
         Caption         =   "Bitmap on the right"
         IconAlign       =   2
      End
      Begin prjbsOctControls.bsOctButton bsOctButton15 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   12648447
         Icon            =   "frmOctButton.frx":0C41
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Margin          =   4
         Caption         =   "Icon on the left"
         Alignment       =   2
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Truncated captions"
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   2175
      Begin prjbsOctControls.bsOctButton bsOctButton7 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Right aligned"
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IsTruncated     =   -1  'True
         Caption         =   "bsOctButton is the most innovative control"
         Alignment       =   2
      End
      Begin prjbsOctControls.bsOctButton bsOctButton6 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Middle aligned"
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IsTruncated     =   -1  'True
         Caption         =   "bsOctButton is the most innovative control"
         Alignment       =   1
      End
      Begin prjbsOctControls.bsOctButton bsOctButton5 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Guess"
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IsTruncated     =   -1  'True
         Caption         =   "bsOctButton is the most innovative control"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text Alignment"
      Height          =   1815
      Left            =   4800
      TabIndex        =   14
      Top             =   0
      Width           =   2055
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Colour          =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Left align"
      End
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Middle align"
         Alignment       =   1
         BackPicture     =   "frmOctButton.frx":0D9B
      End
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Right align"
         Alignment       =   2
         BackPicture     =   "frmOctButton.frx":176C
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Backgrounds"
      Height          =   1815
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Single bitmap"
         CaptionColour   =   16777215
         BackType        =   1
         BackPicture     =   "frmOctButton.frx":1CBE
      End
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stretched bitmap"
         CaptionColour   =   12648447
         BackType        =   3
         BackPicture     =   "frmOctButton.frx":D225
      End
      Begin prjbsOctControls.bsOctButton opBack 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tiled bitmap"
         CaptionColour   =   16777152
         BackType        =   2
         BackPicture     =   "frmOctButton.frx":DBF6
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colours"
      Height          =   2295
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   2295
      Begin prjbsOctControls.bsOctButton opColour 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   16744703
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Flat bsOctButton"
         CaptionColour   =   0
         HighlightColour =   16761024
         HighlightDKColour=   16759929
         ShadowColour    =   14248960
         ShadowDKColour  =   8930304
         Alignment       =   1
         BorderStyle     =   0
         FlatBorderColour=   8388863
      End
      Begin prjbsOctControls.bsOctButton opColour 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Raised Thin bsOctButton"
         HighlightColour =   8454143
         HighlightDKColour=   16759929
         ShadowColour    =   14248960
         ShadowDKColour  =   4227072
         Alignment       =   1
         BorderStyle     =   1
      End
      Begin prjbsOctControls.bsOctButton opColour 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   16744448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Raised 3D bsOctButton"
         HighlightColour =   16761024
         HighlightDKColour=   16759929
         ShadowColour    =   14248960
         ShadowDKColour  =   8930304
         Alignment       =   1
      End
      Begin prjbsOctControls.bsOctButton opColour 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Colour          =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Etched bsOctButton"
         HighlightColour =   16777215
         HighlightDKColour=   14737632
         ShadowColour    =   12895428
         ShadowDKColour  =   8421504
         Alignment       =   1
         BorderStyle     =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Borders"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin prjbsOctControls.bsOctButton bsOctButton4 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Etched"
         BorderStyle     =   3
      End
      Begin prjbsOctControls.bsOctButton bsOctButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Raised 3D"
      End
      Begin prjbsOctControls.bsOctButton bsOctButton2 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Raised Thin"
         BorderStyle     =   1
      End
      Begin prjbsOctControls.bsOctButton bsOctButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Flat"
         BorderStyle     =   0
      End
   End
End
Attribute VB_Name = "frmOctButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   frmMain.Show
   Me.Hide
End Sub
