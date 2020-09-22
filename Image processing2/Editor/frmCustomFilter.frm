VERSION 5.00
Begin VB.Form frmCustomFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom filter settings"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Reverse"
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2700
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2700
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2115
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2955
      Begin VB.TextBox txtFilter 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2700
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   3060
      Picture         =   "frmCustomFilter.frx":0000
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Prototype"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   2340
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Size"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   2340
      Width           =   735
   End
End
Attribute VB_Name = "frmCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bFromCode As Boolean

Private Sub Check1_Click()
  Combo2_Click
End Sub

Private Sub Combo1_Click()
   Dim i As Long, j As Long, nSize As Long
   nSize = Combo1.ListIndex
   If bFromCode Then
      CustomFilterSize = nSize
      ShowFilter
      ResizeForm
      Exit Sub
   End If
   If nSize <> CustomFilterSize Then
      CustomFilterSize = nSize
      ResizeForm
      ReDim CustomFilter(-(3 + nSize * 2) \ 2 To (3 + nSize * 2) \ 2, -(3 + nSize * 2) \ 2 To (3 + nSize * 2) \ 2)
      If Combo2.ListIndex = 0 Then
         CustomFilter(0, 0) = 1
         ShowFilter
      Else
         Combo2_Click
      End If
   End If
End Sub

Private Sub Combo2_Click()
   If Combo2.ListIndex = 0 Then
      Check1.Visible = False
      Exit Sub
   End If
   Dim Flt As cFilter
   Check1.Visible = True
   Set Flt = New cFilter
   Flt.KernelSize = (3 + Combo1.ListIndex * 2)
   Select Case Combo2.ListIndex
       Case 1: Flt.PrepareKernel KernelFilterTypes.eRectangle, Check1.Value, CustomFilter
       Case 2: Flt.PrepareKernel KernelFilterTypes.eCylinder, Check1.Value, CustomFilter
       Case 3: Flt.PrepareKernel KernelFilterTypes.eGaussian, Check1.Value, CustomFilter
       Case 4: Flt.PrepareKernel KernelFilterTypes.eCone, Check1.Value, CustomFilter
       Case 5: Flt.PrepareKernel KernelFilterTypes.ePyramid, Check1.Value, CustomFilter
       Case 6: Flt.PrepareKernel KernelFilterTypes.eJinc, Check1.Value, CustomFilter
       Case 7: Flt.PrepareKernel KernelFilterTypes.eSinc, Check1.Value, CustomFilter
       Case 8: Flt.PrepareKernel KernelFilterTypes.ePeak, Check1.Value, CustomFilter
       Case 9: Flt.PrepareKernel KernelFilterTypes.eExpDecay, Check1.Value, CustomFilter
       Case 10: Flt.PrepareKernel KernelFilterTypes.eLaplacian, Check1.Value, CustomFilter
   End Select
   bFromCode = True
   Combo1.ListIndex = Int(Flt.KernelSize / 2) - 1
   bFromCode = False
   Set Flt = Nothing
   ShowFilter
End Sub

Private Sub Form_Load()
   Dim varKrnlNames As Variant
   Dim i As Long
   varKrnlNames = Array("Custom", "Rectangle", "Cylinder", "Gaussian", "Cone", "Pyramid", "Jinc", "Sinc", "Peak", "ExpDecay", "Laplacian")
   Combo2.AddItem varKrnlNames(0)
   For i = 1 To 10
       Combo1.AddItem 1 + i * 2 & " x " & 1 + i * 2
       Combo2.AddItem varKrnlNames(i)
   Next i
   For i = 0 To 440
       If i > 0 Then Load txtFilter(i)
       txtFilter(i).Move (i Mod 21) * txtFilter(0).Width, Int(i / 21) * txtFilter(0).Height
       txtFilter(i).Visible = True
       txtFilter(i) = i
   Next i
   Combo1.ListIndex = CustomFilterSize
   Combo2.ListIndex = 0
   ShowFilter
   ResizeForm
End Sub

Private Sub ShowFilter()
   Dim i As Long, j As Long, nSize As Long
   nSize = Combo1.ListIndex
   For i = 0 To 2 + nSize * 2
       For j = 0 To 2 + nSize * 2
           txtFilter(i * 21 + j) = CustomFilter(i - (3 + nSize * 2) \ 2, j - (3 + nSize * 2) \ 2)
       Next j
   Next i
End Sub

Private Sub ResizeForm()
   Dim nSize As Long
   Dim nWidth As Long, nHeight As Long
   nSize = Combo1.ListIndex
   Frame1.Width = (3 + nSize * 2) * (txtFilter(0).Width)
   Frame1.Height = (3 + nSize * 2) * (txtFilter(0).Height)
   Label1.Move Frame1.Left, Frame1.Height + 100
   Label2.Move Frame1.Left + Label1.Width + 150, Frame1.Height + 100
   Combo1.Move Label1.Left, Label1.Top + Label1.Height
   Combo2.Move Label2.Left, Label1.Top + Label1.Height
   Check1.Move Combo2.Left + Combo2.Width + 120, Combo2.Top
   nWidth = Frame1.Width + 210
   If nWidth < 3000 Then nWidth = 3000
   nHeight = Frame1.Height + Combo1.Height * 2 + 450
   Me.Move Screen.Width / 2 - nWidth / 2, Screen.Height / 2 - nHeight / 2, nWidth, nHeight
   Image1.Move Frame1.Left + Frame1.Width, Frame1.Top, ScaleWidth - Frame1.Left - Frame1.Width, Frame1.Height
   Image1.Visible = (nSize = 0)
End Sub
