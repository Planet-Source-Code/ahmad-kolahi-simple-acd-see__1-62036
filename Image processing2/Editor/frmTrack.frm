VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmTrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sldBalance 
      Height          =   435
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   767
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   1740
      Width           =   975
   End
   Begin MSComctlLib.Slider sldBalance 
      Height          =   435
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   540
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   767
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      TickFrequency   =   10
      Value           =   50
   End
   Begin MSComctlLib.Slider sldBalance 
      Height          =   435
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   1020
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   767
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Apply To:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gamma"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Contrast"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Brightness:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   795
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   375
      Index           =   1
      Left            =   4140
      TabIndex        =   7
      Top             =   660
      Width           =   315
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   375
      Index           =   0
      Left            =   4140
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bFromCode As Boolean

Private Sub Command1_Click()
  Me.Hide
  DoEvents
  RestoreValues
  frmTest.UpdateChanges True
End Sub

Private Sub Command2_Click()
  Me.Hide
  DoEvents
  RestoreValues
  frmTest.UpdateChanges False
End Sub

Private Sub Form_Load()
   Combo1.AddItem "All Colors"
   Combo1.AddItem "Red"
   Combo1.AddItem "Green"
   Combo1.AddItem "Blue"
   RestoreValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Me.Hide
      DoEvents
      RestoreValues
      frmTest.UpdateChanges False
   End If
End Sub

Private Sub sldBalance_Change(Index As Integer)
   If Index = 2 Then
      lblBalance(Index).Caption = Left(Format(10 ^ ((sldBalance(Index).Value - 50) / 50), "0.00"), 4)
   Else
      lblBalance(Index).Caption = sldBalance(Index).Value
   End If
   If bFromCode Then Exit Sub
   frmTest.DoBalance Index, Combo1.ListIndex, sldBalance(Index).Value
End Sub

Private Sub RestoreValues()
   Combo1.ListIndex = 0
   bFromCode = True
   For i = 0 To 2
       sldBalance(i).Value = 50
       lblBalance(i) = 50
   Next i
   lblBalance(2) = "1.0"
   bFromCode = False
End Sub

Private Sub sldBalance_Scroll(Index As Integer)
    If Index = 2 Then sldBalance(Index).Text = Left(Format(10 ^ ((sldBalance(Index).Value - 50) / 50), "0.00"), 4)
End Sub
