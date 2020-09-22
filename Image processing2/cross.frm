VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmdecomp 
   Caption         =   "Form2"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form2"
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   4680
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   840
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   4680
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   840
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   600
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   480
      Width           =   3405
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save 4"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save 3"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save 2"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save 1"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   600
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   2
      Top             =   3840
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   4320
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   4320
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   3840
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      Filter          =   "*.bmp,*.gif,*.jpeg"
   End
   Begin VB.Label Label4 
      Caption         =   "Picture4"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Picture3"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Picture2"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Picture 1"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmdecomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
'Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dialog.ShowSave
Print Dialog.FileName
SavePicture picture1.Picture, "jhgj.bmp"
End Sub

Private Sub Form_Activate()
   
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    w = frmPicViewer.picCopy.ScaleWidth
    h = frmPicViewer.picCopy.ScaleHeight
    For pr = 1 To 100
        If pr / 100 * w > 200 Then Exit For
        If pr / 100 * h > 200 Then Exit For
    Next
    wp1 = pr / 100 * w
    hp1 = pr / 100 * h
    
    frmdecomp.Picture5.Width = Int(w / 2)
    frmdecomp.Picture5.Height = Int(h / 2)
    frmdecomp.Picture6.Width = Int(w / 2)
    frmdecomp.Picture6.Height = Int(h / 2)
    frmdecomp.Picture7.Width = Int(w / 2)
    frmdecomp.Picture7.Height = Int(h / 2)
    frmdecomp.Picture8.Width = Int(w / 2)
    frmdecomp.Picture8.Height = Int(h / 2)
    frmdecomp.picture1.Width = wp1
    frmdecomp.picture1.Height = hp1
    frmdecomp.Picture2.Width = wp1
    frmdecomp.Picture2.Height = hp1
    frmdecomp.Picture3.Width = wp1
    frmdecomp.Picture3.Height = hp1
    frmdecomp.Picture4.Width = wp1
    frmdecomp.Picture4.Height = hp1
    
    gh = SetStretchBltMode(Picture5.hdc, 3)
    gh = SetStretchBltMode(Picture6.hdc, 3)
    gh = SetStretchBltMode(Picture7.hdc, 3)
    gh = SetStretchBltMode(Picture8.hdc, 3)
    gh = SetStretchBltMode(picture1.hdc, 3)
    gh = SetStretchBltMode(Picture2.hdc, 3)
    gh = SetStretchBltMode(Picture3.hdc, 3)
    gh = SetStretchBltMode(Picture4.hdc, 3)
   
      
    
For X = 0 To frmPicViewer.picCopy.ScaleWidth - 1 Step 2
    For Y = 0 To frmPicViewer.picCopy.ScaleHeight - 1 Step 2
        SetPixelV frmdecomp.Picture5.hdc, X / 2, Y / 2, GetPixel(frmPicViewer.picCopy.hdc, X, Y)
        SetPixelV frmdecomp.Picture6.hdc, X / 2, Y / 2, GetPixel(frmPicViewer.picCopy.hdc, X + 1, Y)
        SetPixelV frmdecomp.Picture7.hdc, X / 2, Y / 2, GetPixel(frmPicViewer.picCopy.hdc, X, Y + 1)
        SetPixelV frmdecomp.Picture8.hdc, X / 2, Y / 2, GetPixel(frmPicViewer.picCopy.hdc, X + 1, Y + 1)
    Next
Next

    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture5.hdc, 0, 0, Int(w / 2), Int(h / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture6.hdc, 0, 0, Int(w / 2), Int(h / 2), vbSrcCopy)
    mresult = StretchBlt(Picture3.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(w / 2), Int(h / 2), vbSrcCopy)
    mresult = StretchBlt(Picture4.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(w / 2), Int(h / 2), vbSrcCopy)

End Sub
