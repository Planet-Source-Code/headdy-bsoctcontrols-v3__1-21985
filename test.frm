VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   3360
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   720
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Form_Load()
   Picture2.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Picture1_Click()
   On Error GoTo handle
   CommonDialog1.ShowOpen
   Picture1.Picture = LoadPicture(CommonDialog1.filename)
   StretchBlt Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, _
   Picture1.hdc, 0, 0, 88, 31, _
   vbSrcCopy
   'Form1.Refresh
handle:
End Sub
