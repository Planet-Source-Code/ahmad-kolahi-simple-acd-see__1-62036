VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDlgExtra 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmDlgExtra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      Height          =   2355
      Left            =   4620
      ScaleHeight     =   2295
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picSaveOptions 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame1 
         Caption         =   "JPEG"
         Height          =   1335
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4275
         Begin VB.CheckBox Check1 
            Caption         =   "Save in gray scale"
            Height          =   255
            Left            =   600
            TabIndex        =   3
            Top             =   960
            Width           =   2055
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   495
            Left            =   480
            TabIndex        =   4
            Top             =   480
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   873
            _Version        =   393216
            Min             =   10
            Max             =   100
            SelStart        =   50
            TickStyle       =   1
            TickFrequency   =   10
            Value           =   50
         End
         Begin VB.Label Label1 
            Caption         =   "Save quality"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Lowest"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   7
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Best"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3540
            TabIndex        =   6
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lblQuality 
            Alignment       =   1  'Right Justify
            Caption         =   "50"
            Height          =   315
            Left            =   3900
            TabIndex        =   5
            Top             =   600
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "BMP"
         Enabled         =   0   'False
         Height          =   1395
         Left            =   0
         TabIndex        =   1
         Top             =   1500
         Width           =   4275
         Begin VB.Label Label5 
            Caption         =   "Not implemented yet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   360
            TabIndex        =   10
            Top             =   420
            Width           =   3555
         End
      End
   End
   Begin VB.Label lblPreview 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Top             =   2940
      Width           =   1935
   End
End
Attribute VB_Name = "frmDlgExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Slider1.Value = 75
   lblValue = Slider1.Value
   picPreview.Move picSaveOptions.Left, picSaveOptions.Top, picSaveOptions.Width, picSaveOptions.Height - lblPreview.Height
   lblPreview.Left = picPreview.Left
   lblPreview.Width = picPreview.Width
End Sub

Private Sub Slider1_Change()
   lblQuality = Slider1.Value
End Sub

