VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pbProgress 
      DrawStyle       =   6  'Inside Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      ScaleHeight     =   315
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   1020
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   420
      Width           =   3555
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Max As Long

Private Sub Form_Load()
   pbProgress.BackColor = vbYellow
   pbProgress.AutoRedraw = True
   pbColor.Move 0, 0, pbProgress.Width, pbProgress.Height
   Label1.Caption = "Processing image. Please wait..."
End Sub

Public Property Let Max(vData As Long)
   m_Max = vData
End Property

Public Property Let ProgressValue(vData As Long)
    Dim iPercent As Single
    Dim s As String
    iPercent = vData / m_Max
    pbProgress.Cls
    s = Int(iPercent * 100) & "%"
    pbProgress.CurrentX = pbProgress.Width / 2 - pbProgress.TextWidth(s) / 2
    pbProgress.CurrentY = pbProgress.Height / 2 - pbProgress.TextHeight(s) / 1.5
    pbProgress.Print s
    BitBlt pbProgress.hDC, 0, 0, iPercent * pbProgress.Width / Screen.TwipsPerPixelX, pbProgress.Height / Screen.TwipsPerPixelY, pbColor.hDC, 0, 0, vbSrcInvert
End Property

