VERSION 5.00
Begin VB.Form frmFltOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter options"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   315
      Left            =   4740
      TabIndex        =   26
      Top             =   4500
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   315
      Left            =   3720
      TabIndex        =   25
      Top             =   4500
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   2700
      TabIndex        =   24
      Top             =   4500
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   60
      TabIndex        =   23
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enhanced (Laplacian) filters"
      Height          =   1635
      Left            =   2940
      TabIndex        =   16
      Top             =   2700
      Width           =   2715
      Begin VB.ComboBox cmbEnhSize 
         Height          =   315
         Index           =   1
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox cmbEnhSize 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox cmbEnhSize 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Edges"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Details"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Focus"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1140
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rank filters"
      Height          =   1635
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   2715
      Begin VB.ComboBox cmbRankSize 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1140
         Width           =   1275
      End
      Begin VB.ComboBox cmbRankSize 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox cmbRankSize 
         Height          =   315
         Index           =   1
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maximum"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minimum"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblRank 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Median"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kernel filters"
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cmbKrnlType 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   540
         Width           =   1395
      End
      Begin VB.ComboBox cmbKrnlSize 
         Height          =   315
         Index           =   0
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   1275
      End
      Begin VB.ComboBox cmbKrnlPower 
         Height          =   315
         Index           =   0
         Left            =   4260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFltName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmFltOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varKrnlNames As Variant
Dim varFltNames As Variant
Dim varEdgeNames As Variant
Dim varDirectionNames As Variant
Dim varDirectionNames_1 As Variant

Private Sub cmbKrnlType_Click(Index As Integer)
   Dim i As Long
   If Index = 3 Then
      cmbKrnlSize(3).Clear
      If cmbKrnlType(3).ListIndex > 1 Then
         For i = 0 To 1
            cmbKrnlSize(3).AddItem varDirectionNames_1(i)
         Next i
      Else
         For i = 0 To 7
            cmbKrnlSize(3).AddItem varDirectionNames(i)
         Next i
      End If
      cmbKrnlSize(3).ListIndex = 0
   End If
End Sub

Private Sub Command1_Click()
  frmCustomFilter.Show vbModal
End Sub

Private Sub Command2_Click()
   Command3_Click
   SaveFilterSettings
   Unload Me
End Sub

Private Sub Command3_Click()
   Dim i As Long
   For i = 0 To 3
       KernelFilterType(i) = cmbKrnlType(i).ListIndex
       KernelFilterSize(i) = cmbKrnlSize(i).ListIndex
       KernelFilterPower(i) = cmbKrnlPower(i).ListIndex + 1
       If i < 3 Then RankFilterSize(i) = cmbRankSize(i).ListIndex
       If i < 3 Then EnhancedFilterSize(i) = cmbEnhSize(i).ListIndex
   Next i
End Sub

Private Sub Command4_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Long, j As Long
   Label1 = "Filter type"
   Label2 = "Kernel type"
   Label3 = "Kernel size"
   Label4 = "Filter power"
   varKrnlNames = Array("Rectangle", "Cylinder", "Gaussian", "Cone", "Pyramid", "Jinc", "Sinc", "Peak", "ExpDecay", "Laplacian")
   varFltNames = Array("Soften", "Blur", "Sharpen", "Edge detect")
   varEdgeNames = Array("Gradient", "Emboss", "Sobel", "Prewitt")
   varDirectionNames = Array("NORTH", "NE", "EAST", "SE", "SOUTH", "SW", "WEST", "NW")
   varDirectionNames_1 = Array("HORZ", "VERT")
   For i = 1 To 3
       Load lblFltName(i)
       lblFltName(i).Move lblFltName(i - 1).Left, lblFltName(i - 1).Top + lblFltName(i - 1).Height * 1.5
       lblFltName(i).Visible = True
       Load cmbKrnlType(i)
       cmbKrnlType(i).Move cmbKrnlType(i - 1).Left, cmbKrnlType(i - 1).Top + cmbKrnlType(i - 1).Height * 1.5
       cmbKrnlType(i).Visible = True
       Load cmbKrnlSize(i)
       cmbKrnlSize(i).Move cmbKrnlSize(i - 1).Left, cmbKrnlSize(i - 1).Top + cmbKrnlSize(i - 1).Height * 1.5
       cmbKrnlSize(i).Visible = True
       Load cmbKrnlPower(i)
       cmbKrnlPower(i).Move cmbKrnlPower(i - 1).Left, cmbKrnlPower(i - 1).Top + cmbKrnlPower(i - 1).Height * 1.5
       cmbKrnlPower(i).Visible = True
   Next i
   For i = 0 To 3
       For j = 1 To 10
           If i < 3 Then
              cmbKrnlType(i).AddItem varKrnlNames(j - 1)
              cmbKrnlSize(i).AddItem 1 + j * 2 & " x " & 1 + j * 2
              cmbRankSize(i).AddItem 1 + j * 2 & " x " & 1 + j * 2
              cmbEnhSize(i).AddItem 1 + j * 2 & " x " & 1 + j * 2
           ElseIf j < 9 Then
              cmbKrnlSize(i).AddItem varDirectionNames(j - 1)
           End If
           cmbKrnlPower(i).AddItem j
       Next j
       lblFltName(i) = varFltNames(i)
       cmbKrnlType(3).AddItem varEdgeNames(i)
       cmbKrnlType(i).ListIndex = KernelFilterType(i)
       cmbKrnlSize(i).ListIndex = KernelFilterSize(i)
       cmbKrnlPower(i).ListIndex = KernelFilterPower(i) - 1
       If i < 3 Then cmbRankSize(i).ListIndex = RankFilterSize(i)
       If i < 3 Then cmbEnhSize(i).ListIndex = EnhancedFilterSize(i)
   Next i
   Command1.Caption = "C&ustom filter"
   Command2.Caption = "&Save"
   Command3.Caption = "&Apply"
   Command4.Caption = "&Cancel"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   For i = 1 To 3
       Unload lblFltName(i)
       Unload cmbKrnlType(i)
       Unload cmbKrnlSize(i)
       Unload cmbKrnlPower(i)
   Next i
End Sub
