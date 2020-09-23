VERSION 5.00
Object = "*\AbsOctButton.vbp"
Begin VB.Form frmOctPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bsOctButton and bsOctPanel demo"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Alignment"
      Height          =   1815
      Left            =   4800
      TabIndex        =   17
      Top             =   3360
      Width           =   2775
      Begin prjbsOctControls.bsOctPanel opAlignment 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         BorderStyle     =   1
      End
      Begin prjbsOctControls.bsOctPanel opAlignment 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin prjbsOctControls.bsOctPanel opAlignment 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BorderStyle     =   1
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Borders"
      Height          =   2295
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      Begin prjbsOctControls.bsOctPanel opBorder 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "Flat border"
      End
      Begin prjbsOctControls.bsOctPanel opBorder 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "Thin 3D border"
      End
      Begin prjbsOctControls.bsOctPanel opBorder 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         Caption         =   "3D border"
      End
      Begin prjbsOctControls.bsOctPanel opBorder 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Colour          =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "Etched border"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colours"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
      Begin prjbsOctControls.bsOctPanel bsOctPanel8 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Colour          =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "bsOctPanel with etched border + custom colours"
         CaptionColour   =   16777215
         HighlightColour =   16761087
         HighlightDkColour=   16744703
         ShadowColour    =   12583104
         ShadowDkColour  =   8388736
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel7 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Colour          =   16755370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         Caption         =   "bsOctPanel with 3D border + custom colours"
         HighlightColour =   16777215
         HighlightDkColour=   16763594
         ShadowColour    =   16744576
         ShadowDkColour  =   16727614
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel6 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Colour          =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "bsOctPanel with thin border + custom colours"
         HighlightColour =   12648384
         ShadowColour    =   16711935
         ShadowDkColour  =   16711935
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel2 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Colour          =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         FlatBorderColour=   16711680
         Caption         =   "bsOctPanel with flat border"
         HighlightColour =   16768477
         HighlightDkColour=   16756912
         ShadowColour    =   16744576
      End
   End
   Begin prjbsOctControls.bsOctPanel bsOctPanel 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      Caption         =   $"Form1.frx":030A
   End
   Begin VB.Frame Frame2 
      Caption         =   "Backgrounds"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4575
      Begin prjbsOctControls.bsOctPanel bsOctPanel5 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
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
         BackType        =   2
         BackPicture     =   "Form1.frx":03F8
         BorderStyle     =   2
         Caption         =   "bsOctPanel with a tiled background"
         CaptionColour   =   16777152
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel4 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
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
         BackType        =   1
         BackPicture     =   "Form1.frx":094A
         BorderStyle     =   2
         Caption         =   "bsOctPanel with a bitmap background"
         CaptionColour   =   16777215
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel3 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
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
         BorderStyle     =   2
         Caption         =   "bsOctPanel as it comes"
      End
      Begin prjbsOctControls.bsOctPanel bsOctPanel1 
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   714
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
         BackType        =   3
         BackPicture     =   "Form1.frx":BEB1
         BorderStyle     =   2
         Caption         =   "bsOctPanel with a stretched bitmap background"
         CaptionColour   =   16777215
      End
   End
End
Attribute VB_Name = "frmOctPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   frmMain.Show
   Me.Hide
End Sub
