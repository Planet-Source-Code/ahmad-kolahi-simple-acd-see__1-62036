VERSION 5.00
Begin VB.Form frmResize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resize image"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmResize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2700
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   420
      TabIndex        =   1
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose resize options"
      Height          =   2415
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   5
         Left            =   2340
         TabIndex        =   14
         Top             =   2040
         Width           =   1755
      End
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   4
         Left            =   2340
         TabIndex        =   13
         Top             =   1680
         Width           =   1755
      End
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   3
         Left            =   2340
         TabIndex        =   12
         Top             =   1320
         Width           =   1755
      End
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   11
         Top             =   960
         Width           =   1755
      End
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   10
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optResize 
         Caption         =   "optResize"
         Height          =   255
         Index           =   0
         Left            =   2340
         TabIndex        =   9
         Top             =   240
         Width           =   1755
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   435
         Left            =   300
         TabIndex        =   8
         Top             =   1800
         Width           =   1875
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   1440
         Width           =   1875
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   195
         Left            =   1380
         TabIndex        =   6
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const ES_NUMBERSONLY = &H2000&
Private Const GWL_STYLE = (-16)

Public lWidth As Long, lHeight As Long
Public bUseResampling As Boolean
Dim bPreserveRatio As Boolean
Dim NewWidth As Long, NewHeight As Long
Dim bFromCode As Boolean

Private Sub Check1_Click()
  bPreserveRatio = Check1.Value
  If bPreserveRatio Then CheckAspect
End Sub

Private Sub Check2_Click()
  bUseResampling = Check2.Value
End Sub

Private Sub Command1_Click()
   lWidth = txtWidth
   lHeight = txtHeight
   bUseResampling = Check2.Value
   Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
  Label1 = "Current size: " & lWidth & " x " & lHeight
  txtWidth = lWidth
  txtHeight = lHeight
  Check2.Value = Abs(bUseResampling)
End Sub

Private Sub Form_Load()
  Dim sOptArray As Variant
  NewWidth = lWidth
  NewHeight = lHeight
  sOptArray = Array("&600 x 480 Pixels", "&800 x 600 Pixels", "&1024 x 768 Pixels", "Best &fit to desktop", "&Half size", "&Double size")
  Command1.Caption = "&OK"
  Command2.Caption = "&Cancel"
  SetWindowLong txtWidth.hWnd, GWL_STYLE, GetWindowLong(txtWidth.hWnd, GWL_STYLE) Or ES_NUMBERSONLY
  SetWindowLong txtHeight.hWnd, GWL_STYLE, GetWindowLong(txtHeight.hWnd, GWL_STYLE) Or ES_NUMBERSONLY
  Label2 = "x"
  Label3 = "New size:"
  For i = 0 To 5
     optResize(i).Caption = sOptArray(i)
  Next i
  Check1.Caption = "&Preserve aspect ratio"
  Check1.Value = 1
  Check2.Caption = "&Use resampling"
End Sub

Private Sub optResize_Click(Index As Integer)
   Select Case Index
       Case 0: NewWidth = 640: NewHeight = 480
       Case 1: NewWidth = 800: NewHeight = 600
       Case 2: NewWidth = 1024: NewHeight = 768
       Case 3: NewWidth = Screen.Width / Screen.TwipsPerPixelX: NewHeight = Screen.Height / Screen.TwipsPerPixelY
       Case 4: NewWidth = lWidth / 2: NewHeight = lHeight / 2
       Case 5: NewWidth = lWidth * 2: NewHeight = lHeight * 2
   End Select
   If bPreserveRatio Then CheckAspect
   txtWidth = NewWidth
   txtHeight = NewHeight
End Sub

Private Sub CheckAspect()
   Dim OldRatio As Double, NewRatio As Double
   On Error Resume Next
   OldRatio = lWidth / lHeight
   NewRatio = NewWidth / NewHeight
   If NewRatio = OldRatio Then Exit Sub
   If NewRatio < OldRatio Then
      NewHeight = NewWidth / OldRatio
   Else
      NewWidth = NewHeight * OldRatio
   End If
End Sub

Private Sub txtHeight_Change()
   If bFromCode Then Exit Sub
   bFromCode = True
   If txtHeight = "" Then txtHeight = "0"
   If txtHeight = "0" Then txtHeight = "1"
   NewHeight = CLng(txtHeight)
   If bPreserveRatio Then
      NewWidth = NewHeight * lWidth / lHeight
      If NewWidth = 0 Then NewWidth = 1
      txtWidth = NewWidth
   End If
   bFromCode = False
End Sub

Private Sub txtWidth_Change()
   If bFromCode Then Exit Sub
   bFromCode = True
   If txtWidth = "" Then txtWidth = "0"
   If txtWidth = "0" Then txtWidth = "1"
   NewWidth = CLng(txtWidth)
   If bPreserveRatio Then
      NewHeight = NewWidth * lHeight / lWidth
      If NewHeight = 0 Then NewHeight = 1
      txtHeight = NewHeight
   End If
   bFromCode = False
End Sub
