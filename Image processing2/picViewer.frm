VERSION 5.00
Begin VB.Form frmPicViewer 
   Caption         =   "Auto image viewer"
   ClientHeight    =   7590
   ClientLeft      =   300
   ClientTop       =   345
   ClientWidth     =   10020
   ClipControls    =   0   'False
   Icon            =   "picViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2289.275
   ScaleMode       =   0  'User
   ScaleWidth      =   3602.896
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   14280
      TabIndex        =   55
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Next-->"
      Height          =   375
      Left            =   14040
      TabIndex        =   54
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   12960
      TabIndex        =   53
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox piccmdFitInLargeContainer 
      Height          =   315
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   645
      TabIndex        =   46
      ToolTipText     =   "Fit in"
      Top             =   4800
      Width           =   705
      Begin VB.CommandButton cmdViewLarge 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Full"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Enlarge viewport"
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.PictureBox piccmdOrigSizeContainer 
      Height          =   315
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   690
      TabIndex        =   45
      Top             =   4800
      Width           =   744
      Begin VB.CommandButton cmdOrigSize 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Original"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Return to original size"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.ComboBox cboResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "+/- percentage"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox piccmdFitInContainer 
      Height          =   315
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   675
      TabIndex        =   43
      ToolTipText     =   "Fit in"
      Top             =   4800
      Width           =   734
      Begin VB.CommandButton cmdFitIn 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fit"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Image fit in"
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.ComboBox cboPattern 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   240
      Width           =   1545
   End
   Begin VB.PictureBox picViewZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5055
      Left            =   240
      ScaleHeight     =   420.189
      ScaleMode       =   0  'User
      ScaleWidth      =   421
      TabIndex        =   31
      Top             =   5160
      Width           =   6375
      Begin VB.PictureBox picViewTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   150
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   33
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox picCopy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   1080
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   32
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox picViewX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4995
         Left            =   0
         ScaleHeight     =   420.317
         ScaleMode       =   0  'User
         ScaleWidth      =   417
         TabIndex        =   34
         Top             =   0
         Width           =   6285
      End
   End
   Begin VB.FileListBox FilList 
      Height          =   4185
      Left            =   3720
      TabIndex        =   30
      Top             =   240
      Width           =   3165
   End
   Begin VB.VScrollBar VsbView 
      Height          =   5055
      Left            =   6600
      TabIndex        =   29
      Top             =   5160
      Width           =   255
   End
   Begin VB.HScrollBar HsbView 
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Top             =   10200
      Width           =   6225
   End
   Begin VB.DirListBox DirList 
      Height          =   3915
      Left            =   210
      TabIndex        =   27
      Top             =   570
      Width           =   3345
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   1920
      TabIndex        =   26
      Top             =   240
      Width           =   1635
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00800000&
      Caption         =   "Search all subdir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   9480
      TabIndex        =   25
      ToolTipText     =   "List all image files in subdir, of a search pattern"
      Top             =   240
      Width           =   1755
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00800000&
      Caption         =   "Images in dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   7320
      TabIndex        =   24
      ToolTipText     =   "Disp all images in dir, of a file pattern"
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.PictureBox picAuto 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   8400
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   9240
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame fraImagesPanel 
      Caption         =   "List panel - images in directory"
      ForeColor       =   &H00800000&
      Height          =   9855
      Left            =   7080
      TabIndex        =   4
      Top             =   600
      Width           =   7815
      Begin VB.PictureBox picProgressContainer 
         AutoRedraw      =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   477
         TabIndex        =   51
         Top             =   840
         Width           =   7215
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   0
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   52
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox picImagesInDirContainer 
         Height          =   585
         Left            =   120
         ScaleHeight     =   525
         ScaleWidth      =   7185
         TabIndex        =   20
         Top             =   240
         Width           =   7245
         Begin VB.Label lblImagesPanelDir 
            Caption         =   "lblImagesPanelDir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   30
            TabIndex        =   21
            Top             =   0
            Width           =   4125
         End
      End
      Begin VB.CommandButton cmdListImages 
         Height          =   345
         Left            =   540
         Picture         =   "picViewer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Disp all  images in dir"
         Top             =   1230
         Width           =   345
      End
      Begin VB.PictureBox picAutoImagesContainer 
         Height          =   405
         Left            =   120
         Picture         =   "picViewer.frx":04D4
         ScaleHeight     =   345
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   1200
         Width           =   375
         Begin VB.CommandButton cmdAutoImagesOn 
            Height          =   345
            Left            =   0
            Picture         =   "picViewer.frx":069E
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Toggle auto images on/off.  Current status is On."
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdAutoImagesOff 
            Height          =   345
            Left            =   0
            Picture         =   "picViewer.frx":0868
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Toggle auto images on/off.  Current status is Off"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.ComboBox cboPicListLot 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select lot number"
         Top             =   8640
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.VScrollBar VsbPicList 
         Height          =   7995
         Left            =   7320
         TabIndex        =   7
         Top             =   1710
         Width           =   240
      End
      Begin VB.PictureBox picListZ 
         Height          =   7995
         Left            =   120
         ScaleHeight     =   529
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   479
         TabIndex        =   5
         Top             =   1710
         Width           =   7245
         Begin VB.PictureBox picListX 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7875
            Left            =   0
            ScaleHeight     =   523
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   475
            TabIndex        =   6
            Top             =   0
            Width           =   7155
         End
      End
      Begin VB.PictureBox piccmdPanelRefContainer 
         Height          =   315
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   1305
         TabIndex        =   38
         Top             =   8880
         Visible         =   0   'False
         Width           =   1365
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   330
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Label lblLotRef 
         Alignment       =   1  'Right Justify
         Caption         =   "Lot:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   50
         Top             =   8880
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblPicsCount 
         Alignment       =   1  'Right Justify
         Caption         =   "Total images in dir:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1020
         TabIndex        =   10
         Top             =   1170
         Width           =   3225
      End
      Begin VB.Label lblPicsOnPanel 
         Alignment       =   1  'Right Justify
         Caption         =   "Images in current lot:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   960
         TabIndex        =   9
         Top             =   1380
         Width           =   3315
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "List panel - search all subdirectories"
      ForeColor       =   &H00800000&
      Height          =   9915
      Left            =   7080
      TabIndex        =   2
      Top             =   570
      Width           =   7828
      Begin VB.PictureBox picSearchAllSubdirContainer 
         Height          =   705
         Left            =   150
         ScaleHeight     =   645
         ScaleWidth      =   7185
         TabIndex        =   22
         Top             =   330
         Width           =   7245
         Begin VB.Label lblSearchPanelDir 
            Caption         =   "lblSearchPanelDir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   7065
         End
      End
      Begin VB.CommandButton cmdStopSearch 
         Height          =   315
         Left            =   4320
         Picture         =   "picViewer.frx":0A32
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Stop search"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   315
         Left            =   3840
         Picture         =   "picViewer.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Search for an image file pattern, in all subdir"
         Top             =   1350
         Width           =   345
      End
      Begin VB.ComboBox cboSearchPattern 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Text            =   "cboSearchPattern"
         Top             =   1350
         Width           =   2295
      End
      Begin VB.ListBox lisFilesFound 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00800000&
         Height          =   7665
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   7335
      End
      Begin VB.Label lblSearchPattern 
         Caption         =   "Search pattern"
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label lblFoundNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Total found: 0"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   9120
         Width           =   1425
      End
   End
   Begin VB.PictureBox piccmdEditContainer 
      Height          =   315
      Left            =   1440
      ScaleHeight     =   255
      ScaleWidth      =   405
      TabIndex        =   48
      Top             =   4080
      Visible         =   0   'False
      Width           =   465
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00000080&
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Edit image with system paint program"
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Label lblFilesCount 
      Alignment       =   1  'Right Justify
      Caption         =   "Files in dir:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   960
      TabIndex        =   56
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblPicSizeH 
      Caption         =   "h="
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   270
      TabIndex        =   47
      ToolTipText     =   "+/- percentage"
      Top             =   4920
      Width           =   825
   End
   Begin VB.Label lblSearchInProgress 
      Caption         =   "Search in progress ........"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1380
      TabIndex        =   37
      Top             =   990
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   210
      Top             =   240
      Width           =   1605
   End
   Begin VB.Label lblPicSizeW 
      Caption         =   "w= "
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   270
      TabIndex        =   35
      ToolTipText     =   "+/- percentage"
      Top             =   4680
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   315
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4545
   End
   Begin VB.Menu popFilList 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popFilListRenameFile 
         Caption         =   "&Rename file (within same dir)"
      End
      Begin VB.Menu PopFilListSaveFile 
         Caption         =   "&Save file (i.e. make a copy)"
      End
      Begin VB.Menu PopFilListMoveFile 
         Caption         =   "&Move file (to a different dir)"
      End
      Begin VB.Menu popFilListDeleteFile 
         Caption         =   "&Delete file"
      End
   End
   Begin VB.Menu popPicView 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popPicViewRenameFile 
         Caption         =   "&Rename file (within same dir)"
      End
      Begin VB.Menu PopPicviewSaveFile 
         Caption         =   "&Save file (i.e. make a copy)"
      End
      Begin VB.Menu PopPicviewMoveFile 
         Caption         =   "&Move file (to a different dir)"
      End
      Begin VB.Menu popPicViewDeleteFile 
         Caption         =   "&Delete file"
      End
   End
   Begin VB.Menu popDirList 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popDirListCreateDir 
         Caption         =   "&Create directory"
      End
      Begin VB.Menu popDirListRenameDir 
         Caption         =   "&Rename directory"
      End
      Begin VB.Menu popDirListDeleteDir 
         Caption         =   "&Delete directory"
      End
   End
End
Attribute VB_Name = "frmPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private WithEvents frmScreen As frmFullScreen
Attribute frmScreen.VB_VarHelpID = -1
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long

  
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
 
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Dim strCurrDir As String
Dim strCurrPattern As String
Dim arrPicFileSpec() As String
Dim picsCount As Integer
Dim picsOnCurrList As Integer
Dim filesCount As Integer

Dim picListRow As Integer
Dim picListCol As Integer
Dim picListSeq As Integer

Dim picListRegionFlag As Boolean
Dim SuspendFlag As Boolean
Dim StopSearchFlag As Boolean
Dim AutoListSwitching As Boolean
Dim MagnifyGlassOn As Boolean
Dim GlassInDrag As Boolean

Dim Xcurr As Single
Dim Ycurr As Single
Dim XRegOffset As Single
Dim YRegOffset As Single

Dim picViewXFileSpec As String

Dim OldX1 As Integer
Dim OldX2 As Integer
Dim OldY1 As Integer
Dim OldY2 As Integer

Dim OrigW As Long
Dim OrigH As Long

Dim X1Reg As Single
Dim X2Reg As Single
Dim Y1Reg As Single
Dim Y2Reg As Single


'---------------------------------------------------------
' Calculation of positions depends on the following values
' For placements and displacements, refer these values.
'---------------------------------------------------------
Const StdW = 45
Const StdH = 45
Const xColNums = 6
Const xRowsPerPanel = 6
Const xPanelNums = 6
Const xRowNums = xRowsPerPanel * xPanelNums
Const xMaxNums = xColNums * xRowNums
Const xGapW = 32
Const xGapH = 32
Const xTextH = 20
Const xTotalGapH = xGapH + xTextH

Dim maxImageWidth As Single
Dim maxImageHeight As Single
Const minImageWidth = 16
Const minImageheight = 16

Const defaultcboResizeIndex = 3
Dim oldcboResizeIndex As Integer

Dim origPicViewLargeBarW As Single
Dim origPicViewLargeContainerW As Single
Dim origPicViewLargeContainerH As Single
Dim origVsbViewLargeH As Single
Dim origVsbViewLargeLeft As Single
Dim origHsbViewLargeW As Single
Dim origHsbViewLargeTop As Single

   ' You may add others to exclude
Const NonPicFileExt = "EXE/COM/DLL/RES/OCX/SYS/BAT/TXT/RTF/LOG/ERR/INI/MDB/PRG/C/HTM/ASP"
Dim mresult
Dim gcdg As Object
'---------------------------------------------------------

Private Sub Command5_Click()
Unload Me
End
End Sub

Private Sub Command6_Click()
 'Delete ("test.jpg")
 frmPicViewer.Hide
 SavePicture frmPicViewer.picViewX.Picture, "test1.jpg"
 'Unload Me
 'Print frmTest.lblGrip.Count
'Load frmTest
frmTest.Show
End Sub

Private Sub Form_Load()
    SuspendFlag = True
    Me.ScaleMode = vbPixels
    picTemp.AutoSize = False
    picTemp.Width = picTemp.Width - picTemp.ScaleWidth + StdW
    picTemp.Height = picTemp.Height - picTemp.ScaleHeight + StdH
    picTemp.Visible = False
    
    picAuto.AutoSize = True
    picAuto.Visible = False
    
      ' Keep a copy of existing values, will be used when cmdViewLargeExit
    'origPicViewLargeBarW = picViewLargeBar.Width
    'origPicViewLargeContainerW = picViewLargeContainer.Width
    'origPicViewLargeContainerH = picViewLargeContainer.Height
    'origVsbViewLargeH = VsbViewLarge.Height
    'origVsbViewLargeLeft = VsbViewLarge.Left
    'origHsbViewLargeW = HsbViewLarge.Width
    'origHsbViewLargeTop = HsbViewLarge.Top
    
      ' Align
    picViewX.Move 0, 0
    'picViewLargeContainer.Top = picViewLargeBar.Top + picViewLargeBar.Height
    'picViewLargeContainer.Left = picViewLargeBar.Left
    'picViewLarge.Move 0, 0

    VsbPicList.min = 0
    VsbPicList.Max = (StdH + xTotalGapH) * xRowNums * _
            ((xRowsPerPanel - 1) / xRowsPerPanel) + (StdH + xTotalGapH)
    VsbPicList.SmallChange = 1
    VsbPicList.LargeChange = StdH + xTotalGapH
    
     ' We display xColNums icons each row
    picListX.Width = (StdW + xGapH) * xColNums + xGapW / 2
     ' xRowNums each column, allow extra xTextH pixels for file name.
    picListX.Height = (StdH + xTotalGapH) * xRowNums
    
    cboPattern.Clear
    cboPattern.AddItem "*.*"
    cboPattern.AddItem "*.bmp"
    cboPattern.AddItem "*.gif"
    cboPattern.AddItem "*.jpeg"
    cboPattern.AddItem "*.jpg"
    cboPattern.AddItem "*.ico"
    cboPattern.AddItem "*.cur"
    cboPattern.AddItem "*.wfm"
    cboPattern.AddItem "*.emf"
    cboPattern.ListIndex = 0
    
    cboSearchPattern.Clear
    cboSearchPattern.AddItem "*.*"
    cboSearchPattern.AddItem "*.bmp"
    cboSearchPattern.AddItem "*.gif"
    cboSearchPattern.AddItem "*.jpeg"
    cboSearchPattern.AddItem "*.jpg"
    cboSearchPattern.AddItem "*.ico"
    cboSearchPattern.AddItem "*.cur"
    cboSearchPattern.AddItem "*.wfm"
    cboSearchPattern.AddItem "*.emf"
    cboSearchPattern.ListIndex = 0
    
    Dim i, j
    
    For i = 1.5 To 2.6 Step 0.1
         'cboMagnify.AddItem i
    Next i
    'cboMagnify.ListIndex = 5
    'picGlass.Width = picGlass.Width - picGlass.ScaleWidth + 150
    'picGlass.Height = picGlass.Height - picGlass.ScaleHeight + 150
    
    maxImageWidth = Screen.Width / Screen.TwipsPerPixelX
    maxImageHeight = Screen.Height / Screen.TwipsPerPixelY
    i = maxImageWidth / picViewZ.Width
    j = maxImageHeight / picViewZ.Height
    HsbView.Max = (i - 1) * picViewZ.Width
    VsbView.Max = (j - 1) * picViewZ.Height
    
    'HsbViewLarge.Max = HsbView.Max
    'VsbViewLarge.Max = VsbView.Max
    
    For i = -75 To -25 Step 25
        cboResize.AddItem i
    Next i
    For i = 0 To 800 Step 50
        i = "+" & i
        cboResize.AddItem i
    Next i
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
      
    lblSearchInProgress.Visible = False
    cmdStopSearch.Visible = False
    StopSearchFlag = False
    picListRegionFlag = False
    AutoListSwitching = False
    picViewXFileSpec = ""
    ClearImagesPanelDisp
    ClearSearchDisp
    strCurrDir = ""
    strCurrPattern = ""
    
    cmdAutoImagesOn.Visible = True
    cmdListImages.Enabled = False
    
    'picViewLargeBar.Visible = False
    picProgressContainer.Visible = False
    'picViewLargeContainer.Visible = False
    cboPicListLot.Appearance = 0
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    lblImagesPanelDir.Caption = DirList.Path
    lblSearchPanelDir.Caption = DirList.Path
    ListImagesInDir
    'Set gcdg = CommonDialog1
    SuspendFlag = False
End Sub



Private Sub cmdExit_Click()
    Screen.MousePointer = vbNormal
    Unload Me
End Sub

Private Sub lblFilListHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  File patterns:" & vbCrLf
     tmp = tmp & "      As and when you change a selection in File Pattern Box, or switch" & vbCrLf
     tmp = tmp & "      between folders, all images of the selected pattern under the folde will be" & vbCrLf
     tmp = tmp & "      displayed automatically in the List Panel (if that panel is set to Images In" & vbCrLf
     tmp = tmp & "      Dir).  There is no limit to the No. of images to be displayed in a folder." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  View image in original size:" & vbCrLf
     tmp = tmp & "      To view an image in its original size, click the file name in the File List Box," & vbCrLf
     tmp = tmp & "      or an image in the List Panel if any.  Selection through the File List Box is" & vbCrLf
     tmp = tmp & "      automatically reflected in the List Panel as well, and vice versa.  (If the panel" & vbCrLf
     tmp = tmp & "      is set to Search All Subdir, then click a displayed file spec there)" & vbCrLf & vbCrLf
     tmp = tmp & "(3)  File functions:" & vbCrLf
     tmp = tmp & "      At the File List Box, right click the mouse to bring up a popup menu.  (You" & vbCrLf
     tmp = tmp & "      may also right click the mouse in the Viewport if there is an image in it)" & vbCrLf & vbCrLf
     tmp = tmp & "(4)  Directory functions:" & vbCrLf
     tmp = tmp & "      At the Dir List Box, right click the mouse to bring up a popup menu." & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblPicViewHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  Usage of Viewport:" & vbCrLf
     tmp = tmp & "      This Viewport is for viewing an image in its original size: (a) When the" & vbCrLf
     tmp = tmp & "      panel is set to Images In Dir, click the file name in File List Box (or an" & vbCrLf
     tmp = tmp & "      image in List Panel).  (b) If the panel is set to Search All Subdir, then" & vbCrLf
     tmp = tmp & "      click a displayed file spec in List Panel." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  Image resizing:" & vbCrLf
     tmp = tmp & "      You may zoom in or zoom out the image until the maximum/mininum width" & vbCrLf
     tmp = tmp & "      &/or height is reached.  Values in combo box are for percentage +/-." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  Image Fit-In or Viewport enlarging:" & vbCrLf
     tmp = tmp & "      Click Fit-In button or Enlarge Viewport button.  A scalable magnifying" & vbCrLf
     tmp = tmp & "      glass is also available in the enlarged viewport." & vbCrLf & vbCrLf
     tmp = tmp & "(4)  Edit image:" & vbCrLf
     tmp = tmp & "      You can click Edit button to invoke your system paint program to edit" & vbCrLf
     tmp = tmp & "       the image - No action will be taken if image is found incompatible" & vbCrLf
     tmp = tmp & "       with your system paint program, e.g. cannot edit ico file." & vbCrLf & vbCrLf
     tmp = tmp & "(5)  File functions:" & vbCrLf
     tmp = tmp & "      If there is an image in the Viewport, right click the mouse there." & vbCrLf
     tmp = tmp & "      (Otherwise use the functions in File List Box)" & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblImagesPanelHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  No. of images per Lot:" & vbCrLf
     tmp = tmp & "      If there are many images, they will be divided into Lots.  Lot Nos. are" & vbCrLf
     tmp = tmp & "      displayed in Lot Box for selection; each lot consists of a max of " & CStr(xMaxNums) & vbCrLf
     tmp = tmp & "      images.  Since there is no limit to the No. of images to be displayed" & vbCrLf
     tmp = tmp & "      under a dir, there is no limit to the No. of Lots a dir may have." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  No. of panels per Lot:" & vbCrLf
     tmp = tmp & "      Eash lot is made of four continuous panels;  upto " & CStr(xMaxNums / xPanelNums) & _
                           " images in each" & vbCrLf
     tmp = tmp & "      visible panel.  Use the vertical scroll bar to move along the panels," & vbCrLf
     tmp = tmp & "      or click a Panel Ref No. to reach that specific panel directly." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  View image in original Size:" & vbCrLf
     tmp = tmp & "      To view an image in its original size, click a displayed image in the" & vbCrLf
     tmp = tmp & "      List Panel.  Alternatively, you may click its file name in File List Box." & vbCrLf
     tmp = tmp & "      Selection through the List Panel is automatically reflected in the File" & vbCrLf
     tmp = tmp & "      List Box as well, and vice versa, i.e. they are synchronized." & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblSearchPanelHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  Search for a file pattern:" & vbCrLf
     tmp = tmp & "      Supply a file pattern in the Search Pattern Box on the" & vbCrLf
     tmp = tmp & "      List Panel (select or type), switch to the directory of" & vbCrLf
     tmp = tmp & "      which all subdirectories are to be searched, then click" & vbCrLf
     tmp = tmp & "      the Search button." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  View image of a found file spec:" & vbCrLf
     tmp = tmp & "      Click a displayed file spec in the List Panel." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  View full file spec:" & vbCrLf
     MsgBox tmp & "      Double click a displayed file spec." & vbCrLf & vbCrLf
End Sub




Private Sub optMode_Click(Index As Integer)
     If Index = 0 Then
          fraImagesPanel.Visible = True
          fraSearch.Visible = False
          
          ClearViewDisp
          If cmdAutoImagesOn.Visible Then
               AutoListSwitching = True
               ListImagesInDir
               AutoListSwitching = False      ' Set back to normal
          End If
     Else
          fraImagesPanel.Visible = False
          fraSearch.Visible = True
          ClearViewDisp
          ClearImagesPanelDisp
     End If
     FilList.Refresh
End Sub




Private Sub cmdAutoImagesOn_Click()
     cmdAutoImagesOn.Visible = False
     picListX.SetFocus
     
     cmdAutoImagesOff.Visible = True
     ClearImagesPanelDisp
     cmdListImages.Enabled = True
End Sub



Private Sub cmdAutoImagesOff_Click()
     cmdAutoImagesOn.Visible = True
     cmdAutoImagesOff.Visible = False
     AutoListSwitching = True       ' Signal to ListImagesInDir not to early exit
     picListX.SetFocus
     
     DirList_Change                 ' Will call ListImagesInDir there
     AutoListSwitching = False      ' Set back to normal
     cmdListImages.Enabled = False
End Sub



Private Sub cmdListImages_Click()
     AutoListSwitching = True
     ListImagesInDir
     AutoListSwitching = False
     picListX.SetFocus
End Sub




Private Sub DrvList_Change()
     On Error GoTo ErrHandler               ' Trap e.g. drive not ready
     DirList.Path = drvList.Drive
     Exit Sub
     
ErrHandler:
     drvList.Drive = DirList.Path
     ErrMsgProc "DrvList_Change"
End Sub



Private Sub DirList_Change()
    FilList.Path = DirList.Path
    If lblSearchInProgress.Visible = False Then
        FilList.Pattern = cboPattern.Text
        ClearViewDisp
        If cmdAutoImagesOn.Visible Then
            ListImagesInDir
        End If
    Else
        FilList.Pattern = cboSearchPattern.Text
    End If
    
    lblImagesPanelDir.Caption = DirList.Path
    lblSearchPanelDir.Caption = DirList.Path
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
End Sub



Private Sub DirList_LostFocus()
    DirList.Path = DirList.List(DirList.ListIndex)
End Sub



Private Sub DirList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    PopupMenu Me.popDirList
End Sub





Private Sub popDirListCreateDir_click()
    On Error GoTo ErrHandler
    Dim currdir As String, newDir As String
    currdir = DirList.List(DirList.ListIndex)
again:
    newDir = InputBox("Type full directory specification:", _
        "Create directory", currdir)
    If newDir = "" Then
         Exit Sub
    End If
    MkDir newDir
    DoEvents
    DirList.Refresh
    Exit Sub
ErrHandler:
    If Err.Number = 75 Then
        MsgBox "Directory already exists/access error"
        GoTo again
    End If
    ErrMsgProc "popDirListCreateDir_click"
End Sub



Private Sub popDirListRenameDir_click()
    On Error GoTo ErrHandler
    Dim newDir As String, origDirAsFile As String
    Dim origFullPath As String, origPathDir As String
    origFullPath = DirList.List(DirList.ListIndex)
    origPathDir = PathSection(origFullPath, 1)
    origDirAsFile = PathSection(origFullPath, 2)
    
again:
    newDir = InputBox("Type new name", "Rename directory", origDirAsFile)
    If newDir = "" Then
         Exit Sub
    End If
    If InStr(newDir, "\") <> 0 Then
         MsgBox "Rename directory within same path only"
         GoTo again
    ElseIf PathSection(origPathDir & newDir, 1) <> PathSection(origFullPath, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newDir = origPathDir & newDir
    Name origFullPath As newDir
    DirList.Path = newDir
    Exit Sub

ErrHandler:
    ErrMsgProc "popDirListRenameDir_click"
End Sub



Private Sub popDirListDeleteDir_click()
    On Error GoTo ErrHandler
    If MsgBox("Sure to delete " & DirList.List(DirList.ListIndex) & vbLf & _
           "and all its contents?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim delDir As String
    delDir = DirList.List(DirList.ListIndex)
    DelFolder delDir
      ' Update
    DirList.Path = drvList.Drive
    Exit Sub
ErrHandler:
    ErrMsgProc "popDirListDeleteDir_click"
End Sub



Public Sub DelFolder(ByVal inDir As String)
    On Error Resume Next
    Dim mFile As String
       ' Safety
    If Len(Dir(inDir, vbDirectory)) = 0 Then
        Exit Sub
    End If
       ' Find first, i.e. retrieve first entry
    mFile = Dir(inDir & "\", vbDirectory + vbHidden + vbSystem)
       ' Loop through
    Do While mFile <> ""
        If mFile = "." Or mFile = ".." Then
              ' Call Dir again without arguments to return the next file in same dir
            mFile = Dir
        Else
              ' Try to use bitwise comparison to see if mFile is a directory
            If (GetAttr(inDir & "\" & mFile) And vbDirectory) = vbDirectory Then
                 DelFolder inDir & "\" & mFile
                 mFile = Dir(inDir & "\", vbDirectory + vbHidden + vbSystem)
            Else
                   ' Avoid run-time error
                 SetAttr inDir & "\" & mFile, vbNormal
                 Kill inDir & "\" & mFile
                   ' Find next
                 mFile = Dir
            End If
        End If
    Loop
    RmDir inDir
End Sub



Private Sub FilList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    PopupMenu Me.popFilList
End Sub




Private Sub popFilListRenameFile_Click()
    On Error GoTo ErrHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    
    Dim newName As String, origName As String
    Dim newFileSpec As String, oldDir As String
    Dim mfilespec As String
    mfilespec = FilList.Path & "\" & FilList.FileName
    oldDir = PathSection(mfilespec, 1)
    origName = PathSection(mfilespec, 2)
    
again:
    newName = InputBox("Type new name including ext", "Rename file", origName)
    If newName = "" Then
         Exit Sub
    End If
    
    If InStr(newName, "\") <> 0 Then
         MsgBox "Rename file within same directory only"
         GoTo again
    ElseIf PathSection(oldDir & newName, 1) <> PathSection(mfilespec, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newFileSpec = oldDir & newName
    Name mfilespec As newFileSpec
    DoEvents
    FilList.Refresh
    If mfilespec = picViewXFileSpec Then
         picViewXFileSpec = newFileSpec
         AutoListSwitching = True
         ListImagesInDir
         AutoListSwitching = False
         ClearViewDisp
    End If
    Exit Sub
    
ErrHandler:
    ErrMsgProc "popFilListRenameFile"
End Sub




Private Sub popFilListSaveFile_Click()
    On Error GoTo ErrHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    gcdg.FileName = mfilespec           ' Will show only FilList.FileName though
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.DefaultExt = "bmp"
    gcdg.FilterIndex = 1
    gcdg.flags = cdlOFNOverwritePrompt
    gcdg.CancelError = True
    gcdg.ShowSave
    
afterCreatingDir:
    If mfilespec <> gcdg.FileName Then
        FileCopy mfilespec, gcdg.FileName
           ' Same dir? (dir returned from PathSection includes "\")
        If PathSection(mfilespec, 1) = PathSection(gcdg.FileName, 1) Then
             FilList.Refresh
             AutoListSwitching = True
             ListImagesInDir
             AutoListSwitching = False
             ClearViewDisp
             lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
         End If
         ChDir mPath
    End If
    Exit Sub
    
ErrHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListSaveFile"
         Else                                   ' Dir not already exists
              If MsgBox("Dir " & PathSection(mfilespec, 1) & " does not exist" _
                   & vbCrLf & "Create it?", vbYesNo + vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(mfilespec, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
     End If
End Sub




Private Sub popFilListMoveFile_Click()
    On Error GoTo ErrHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    Dim mPath As String, mFullSpec As String
    Dim oldDir As String
    mPath = CurDir
    Dim mfilespec As String
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    gcdg.FileName = mfilespec
    
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.CancelError = True
    
again:
    gcdg.ShowSave
    
    mFullSpec = drvList.List(drvList.ListIndex) & gcdg.FileName
    
    If PathSection(mFullSpec, 1) = PathSection(mfilespec, 1) Then
        MsgBox "Cannot move to the same directory"
        GoTo again
    End If
    
afterCreatingDir:
    FileCopy mfilespec, gcdg.FileName
    Kill mfilespec
    DoEvents
    ChDir mPath
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    ClearViewDisp
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    Exit Sub
    
ErrHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListMoveFile_Click"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(gcdg.FileName, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(gcdg.FileName, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popFilListDeleteFile_Click()
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    If MsgBox("Sure to delete " & FilList.FileName & vbCrLf, _
           vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim mfilespec As String
    
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    Kill mfilespec
    DoEvents
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
End Sub




Private Sub picViewZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    If picViewXFileSpec = "" Then
         Exit Sub
    End If
    PopupMenu Me.popPicView
End Sub




Private Sub picViewX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picViewZ_MouseUp Button, Shift, x, y
End Sub



Private Sub popPicViewRenameFile_Click()
    On Error GoTo ErrHandler
    Dim newName As String, origName As String
    Dim newFileSpec As String, oldDir As String
    
    oldDir = PathSection(picViewXFileSpec, 1)
    origName = PathSection(picViewXFileSpec, 2)
    
again:
    newName = InputBox("Type new name including ext", "Rename file", origName)
    If newName = "" Then
         Exit Sub
    End If
    
    If InStr(newName, "\") <> 0 Then
         MsgBox "Rename file within same directory only"
         GoTo again
    ElseIf PathSection(oldDir & newName, 1) <> PathSection(picViewXFileSpec, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newFileSpec = oldDir & newName
    Name picViewXFileSpec As newFileSpec
    DoEvents
    FilList.Refresh
    picViewXFileSpec = newFileSpec
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    ClearViewDisp
    Exit Sub
    
ErrHandler:
    ErrMsgProc "popPicViewRenameFile"
End Sub



Private Sub popPicViewSaveFile_Click()
    On Error GoTo ErrHandler
    
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    gcdg.FileName = picViewXFileSpec
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.DefaultExt = "bmp"
    gcdg.FilterIndex = 1
    gcdg.flags = cdlOFNOverwritePrompt
    gcdg.CancelError = True
    
afterCreatingDir:
    gcdg.ShowSave
    
    SavePicture picViewX.Picture, gcdg.FileName
    DoEvents
    
    mfilespec = gcdg.FileName         ' mfilespec is a full file spec
    
    If mfilespec <> picViewXFileSpec Then
        If PathSection(mfilespec, 1) = PathSection(picViewXFileSpec, 1) Then
             FilList.Refresh
             AutoListSwitching = True
             ListImagesInDir
             AutoListSwitching = False
             lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
             ClearViewDisp
         End If
         ChDir mPath
    End If
    Exit Sub
    
ErrHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListSaveFile"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(mfilespec, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(mfilespec, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popPicViewMoveFile_Click()
    On Error GoTo ErrHandler
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    mfilespec = drvList.List(drvList.ListIndex) & gcdg.FileName
     
    gcdg.FileName = ""
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.CancelError = True
    
again:
    gcdg.ShowSave
    
    If PathSection(mfilespec, 1) = PathSection(picViewXFileSpec, 1) Then
        MsgBox "Cannot move to the same directory"
        GoTo again
    End If
    
afterCreatingDir:
    SavePicture picViewX.Picture, gcdg.FileName
    Kill picViewXFileSpec
    DoEvents
    ChDir mPath
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
    Exit Sub
    
ErrHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popPicViewMoveFile_Click"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(gcdg.FileName, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(gcdg.FileName, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popPicviewDeleteFile_Click()
    If MsgBox("Sure to delete " & picViewXFileSpec, _
           vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Kill picViewXFileSpec
    DoEvents
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
End Sub



Private Sub cmdEdit_Click()
    On Error Resume Next
    Dim mfilespec As String
    mfilespec = picViewXFileSpec
    mresult = ShellExecute(Me.hWnd, "Open", mfilespec, &H0&, &H0&, SW_RESTORE)
        FilList.Path = DirList.Path
    If lblSearchInProgress.Visible = False Then
        FilList.Pattern = cboPattern.Text
        ClearViewDisp
        If cmdAutoImagesOn.Visible Then
            ListImagesInDir
        End If
    Else
        FilList.Pattern = cboSearchPattern.Text
    End If
    
    lblImagesPanelDir.Caption = DirList.Path
    lblSearchPanelDir.Caption = DirList.Path
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
End Sub




Private Sub cboResize_Click()
    If SuspendFlag = True Then
        Exit Sub
    End If
    If OrigW = 0 Or OrigH = 0 Then
        Exit Sub
    End If
    
    Dim W, H
    Dim newW, newH
    Dim mfactor, morigArea, mnewArea
    Dim tmp
   
    W = picViewX.ScaleWidth
    H = picViewX.ScaleHeight
    
    HsbView.Value = 0
    VsbView.Value = 0
    Select Case val(cboResize.Text)
       Case Is > 0
           If OrigW > maxImageWidth Or OrigH > maxImageHeight Then
                  MsgBox "Orig image width/height already exceeded the maximum allowed (" & _
                     CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                  GoTo earlyExit
           End If
           
            ' Checking existing size
           If W >= maxImageWidth Or H >= maxImageHeight Then
               If OrigW < maxImageWidth Or OrigH < maxImageHeight Then
                  MsgBox "Image width/height cannot exceed the maximum allowed (" & _
                     CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
               Else
                  If W >= OrigW Or H >= OrigH Then
                      MsgBox "Image width/height exceeded the max allowed (" & _
                         CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                      GoTo earlyExit
                      Exit Sub
                  End If
             End If
           End If
      Case Is < 0
           If OrigW <= minImageWidth Or OrigH <= minImageheight Then
                  MsgBox "Orig image width/height already reached the maximum allowed (" & _
                     CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
           End If
           
           If W <= minImageWidth Or H <= minImageheight Then
              If OrigW > minImageWidth Or OrigH > minImageheight Then
                  MsgBox "Image width/height cannot go below the minimum allowed (" & _
                     CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
             Else
                  If W <= OrigW Or H <= OrigH Then
                      MsgBox tmp & "Image width/height below the minimum allowed (" & _
                         CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                      GoTo earlyExit
                      Exit Sub
                  End If
             End If
          End If
      Case Else
          If W <> OrigW Or H <> OrigH Then
             picViewX.Width = picViewX.Width - W + OrigW
             picViewX.Height = picViewX.Height - H + OrigH
             picViewX.Picture = LoadPicture()
             picViewX.Picture = picCopy.Picture
          End If
          picViewX.SetFocus
          Exit Sub
    End Select
    
    morigArea = OrigW * OrigH
    mnewArea = morigArea * (100 + val(cboResize.Text)) / 100
    
    mfactor = Sqr(mnewArea / morigArea)
    
    newW = OrigW * mfactor
    newH = OrigH * mfactor
    
    If val(cboResize.Text) > 0 Then
         If newW >= maxImageWidth Or newH >= maxImageHeight Then
             If Not (OrigW > newW Or OrigH > newH) Then
                 MsgBox "Will exceed max allowed " & CStr(maxImageWidth) & "/" & _
                     CStr(maxImageHeight) & " pixels"
                 GoTo earlyExit
                 Exit Sub
             End If
         End If
    Else
         If newW <= minImageWidth Or newH <= minImageheight Then
             If Not (OrigW < newW Or OrigH < newH) Then
                 MsgBox "Will fall below min allowed " & CStr(minImageWidth) & "/" & _
                     CStr(minImageheight) & " pixels"
                 GoTo earlyExit
                 Exit Sub
             End If
         End If
    End If
    
    Screen.MousePointer = vbHourglass

    picTemp.Picture = LoadPicture()
    picTemp.Cls
    picTemp.Width = picTemp.Width - picTemp.ScaleWidth + newW
    picTemp.Height = picTemp.Height - picTemp.ScaleHeight + newH
    
    mresult = StretchPic(picTemp, newW, newH, picViewX, W, H)
    If mresult = 0 Then
        GoTo earlyExit
    End If
    
    picViewX.Picture = LoadPicture()
    picViewX.Cls
    picViewX.Width = picViewX.Width - picViewX.ScaleWidth + newW
    picViewX.Height = picViewX.Height - picViewX.ScaleHeight + newH
    
    picTemp.Picture = picTemp.Image
    
    BitBlt picViewX.hdc, 0, 0, newW, newH, picTemp.hdc, 0, 0, vbSrcCopy
    
    picTemp.Cls
    picViewX.SetFocus
    
    oldcboResizeIndex = cboResize.ListIndex
    Screen.MousePointer = vbDefault
    Exit Sub
    
earlyExit:
    SuspendFlag = True
    cboResize.ListIndex = oldcboResizeIndex
    SuspendFlag = False
    picViewX.SetFocus
End Sub




Private Sub cmdOrigSize_Click()
    HsbView.Value = 0
    VsbView.Value = 0
    picViewX.Width = picViewX.Width - picViewX.ScaleWidth + OrigW
    picViewX.Height = picViewX.Height - picViewX.ScaleHeight + OrigH
    picViewX.Picture = LoadPicture()
    picViewX.Picture = picCopy.Image
    picViewX.SetFocus
    SuspendFlag = True
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
    SuspendFlag = False
End Sub



Private Sub cmdFitIn_Click()
    Dim W, H
    Dim newW, newH
    Dim i
        
    Screen.MousePointer = vbHourglass
    HsbView.Value = 0
    VsbView.Value = 0
    If OrigW <= picViewZ.ScaleWidth And OrigH <= picViewZ.ScaleHeight Then
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + OrigW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + OrigH
        picViewX.Picture = picCopy.Picture
    Else
        W = picCopy.ScaleWidth
        H = picCopy.ScaleHeight
        
        newW = OrigW
        newH = OrigH
        
        i = 0
        Do While newW > picViewZ.ScaleWidth Or newH > picViewZ.ScaleHeight
             i = i + 1
             newW = OrigW * (100 - i) / 100
             newH = OrigH * (100 - i) / 100
             If newW < 16 Or newH < 16 Then
                 Screen.MousePointer = vbDefault
                 MsgBox "Due to the relative proportion of width to height of this image," & _
                    vbCrLf & "unable to implement Fit In"
                    Exit Sub
             End If
        Loop
             
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + newW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + newH
        
        mresult = StretchPic(picViewX, newW, newH, picCopy, W, H)
        If mresult = 0 Then
            GoTo ErrHandler
        End If
        
        picViewX.Picture = picViewX.Image
    
        picViewX.SetFocus
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
ErrHandler:
End Sub




Private Sub cmdViewLarge_Click()
  Set frmScreen = New frmFullScreen
  frmScreen.Show 0, Me

End Sub





Private Sub PrepareSnapShot()
    Dim W, H
    Dim maxW, maxH
    Dim ratioW, ratioH
    picGlass.Picture = LoadPicture()
      ' We take an enlarged copy from the clean picCopy, rather than from picViewLarge
      ' We take the enlarged image from a smaller area, e.g. if 2 times, then we take
      ' from 0 to 1/2 of the size of imgGlass in the picViewLarge.  In order to cover
      ' any part of picViewLarge, X2Reg may exceed picViewLarge.ScaleWidth and Y2Reg
      ' may exceed picViewLarge.ScaleHeight.
    W = (X2Reg - X1Reg) / val(cboMagnify.Text)        ' Enlarge this much
    H = (Y2Reg - Y1Reg) / val(cboMagnify.Text)
    If X2Reg < picViewLarge.ScaleWidth And Y2Reg < picViewLarge.ScaleHeight Then
         StretchBlt picGlass.hdc, 0, 0, picGlass.ScaleWidth, picGlass.ScaleHeight, _
             picCopy.hdc, X1Reg, Y1Reg, W, H, vbSrcCopy
    Else
         maxW = (picViewLarge.ScaleWidth - 1 - X1Reg) / val(cboMagnify.Text)
         maxH = (picViewLarge.ScaleHeight - 1 - Y1Reg) / val(cboMagnify.Text)
         ratioW = maxW / W
         ratioH = maxH / H
         StretchBlt picGlass.hdc, 0, 0, picGlass.ScaleWidth * ratioW, _
             picGlass.ScaleHeight * ratioH, picCopy.hdc, _
             X1Reg, Y1Reg, maxW, maxH, vbSrcCopy
    End If
    picGlass.Picture = picGlass.Image
End Sub



Private Sub UpdateDragging()
       ' Take a fresh copy from picCopy to picViewLarge
    BitBlt picViewLarge.hdc, 0, 0, picViewLarge.ScaleWidth, picViewLarge.ScaleHeight, _
          picCopy.hdc, 0, 0, vbSrcCopy
    picViewLarge.Picture = picViewLarge.Image
       ' Put in the enlarged image
    BitBlt picViewLarge.hdc, X1Reg, Y1Reg, X2Reg, Y2Reg, picGlass.hdc, _
        0, 0, vbSrcCopy
    picViewLarge.Picture = picViewLarge.Image
    DrawRegionLines
End Sub



Private Sub DrawRegionLines()
    picViewLarge.DrawMode = vbInvert
    picViewLarge.Line (X2Reg, Y2Reg)-(X1Reg, Y1Reg), , B
End Sub



Private Sub picViewLarge_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then
         Exit Sub
    End If
    If MagnifyGlassOn = False Then
         Exit Sub
    End If
    
    If Not ((x >= X1Reg And x <= X2Reg) And (y >= Y1Reg And y <= Y2Reg)) Then
         Exit Sub
    End If
    Xcurr = x
    Ycurr = y
    XRegOffset = 0
    YRegOffset = 0
    GlassInDrag = True
End Sub



Private Sub picViewLarge_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetMousePointer x, y
    If MagnifyGlassOn = False Or GlassInDrag = False Then
         Exit Sub
    End If
    XRegOffset = x - Xcurr
    YRegOffset = y - Ycurr
    X1Reg = X1Reg + XRegOffset
    X2Reg = X2Reg + XRegOffset
    Y1Reg = Y1Reg + YRegOffset
    Y2Reg = Y2Reg + YRegOffset
    Xcurr = x
    Ycurr = y
        ' Set borders within which dragging of image is allowed.  As we take the enlarged
        ' image from a smaller area, e.g. if 2 times, we take from 0 to 1/2 of the size of
        ' imgGlass in the picViewLarge (see PrepareSnapShot), we should take that into
        ' account, so that all parts of picViewLarge can be covered when we drag
    If (X1Reg < 0) Or (X2Reg > picViewLarge.ScaleWidth + picGlass.ScaleWidth / val(cboMagnify.Text)) Then
         X1Reg = X1Reg - XRegOffset
         X2Reg = X2Reg - XRegOffset
    End If
    If (Y1Reg < 0) Or (Y2Reg > picViewLarge.ScaleHeight + picGlass.ScaleHeight / val(cboMagnify.Text)) Then
         Y1Reg = Y1Reg - YRegOffset
         Y2Reg = Y2Reg - YRegOffset
    End If
    
    PrepareSnapShot
    UpdateDragging
End Sub



Private Sub picviewlarge_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GlassInDrag = False
End Sub



Private Sub SetMousePointer(inX, inY)
    If MagnifyGlassOn Then
         Dim W, H
         W = X1Reg + picGlass.ScaleWidth
         H = Y1Reg + picGlass.ScaleHeight
         If (inX > X1Reg And inX < W) And (inY > Y1Reg And inY < H) Then
              picViewLarge.MousePointer = vbSizeAll
         Else
              picViewLarge.MousePointer = vbDefault
         End If
    Else
        picViewLarge.MousePointer = vbDefault
    End If
End Sub






Private Sub HsbView_Change()
    picViewX.Left = -HsbView.Value
End Sub



Private Sub VsbView_Change()
    picViewX.Top = -VsbView.Value
End Sub


Private Sub VsbPicList_Change()
    picListX.Top = -VsbPicList.Value
End Sub




Private Sub HsbViewLarge_Change()
    picViewLarge.Left = -HsbViewLarge.Value
End Sub



Private Sub VsbViewLarge_Change()
    picViewLarge.Top = -VsbViewLarge.Value
End Sub




Private Sub PicListx_MouseDown(Button As Integer, Shift As Integer, _
             x As Single, y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    GetpicSeq x, y
    If picListSeq = 0 Then
        Exit Sub
    End If
    
    SuspendFlag = True
    lisFilesFound.Refresh
    
    If FilList.ListCount > 0 Then
         Dim i, j
         If cboPattern.Text <> "*.*" Then
              i = (CInt(cboPicListLot.Text) - 1) * xMaxNums + picListSeq
              FilList.ListIndex = i - 1
         Else
              i = (CInt(cboPicListLot.Text) - 1) * xMaxNums
              j = CInt(arrPicFileSpec(i + picListSeq - 1, 2))
              FilList.ListIndex = j
         End If
         HsbView.Value = 0
         VsbView.Value = 0
    Else
         picListX.Cls                ' Don't leave a region
    End If
    SuspendFlag = False
    
    If picListRegionFlag = False Then
        DispPicListRegion
    Else
        If OldX1 < x Or OldX2 > x Or OldY1 < y Or OldY2 > y Then
             ClearPicListRegion
        End If
        DispPicListRegion
    End If
   
End Sub



Private Sub cboPattern_Click()
        FilList.Pattern = cboPattern.Text
    If optMode(0).Value = True Then
         ClearViewDisp            ' Clear current picture
         ListImagesInDir          ' List dir related display
    Else
         ClearSearchDisp
    End If
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
End Sub



Private Sub ClearViewDisp()
    If picListRegionFlag Then
        ClearPicListRegion
    End If
    picViewX.Picture = LoadPicture()
    HsbView.Value = 0
    VsbView.Value = 0
    
    picViewX.Width = picViewZ.Width
    picViewX.Height = picViewZ.Height
    lblPicSizeW.Caption = ""
    lblPicSizeH.Caption = ""
    
    piccmdEditContainer.Visible = False
    cboResize.Visible = False
    If SuspendFlag = False Then
         cboResize.ListIndex = defaultcboResizeIndex
         oldcboResizeIndex = defaultcboResizeIndex
    End If
    piccmdOrigSizeContainer.Visible = False
    piccmdFitInContainer.Visible = False
    piccmdFitInLargeContainer.Visible = False
    
    picViewXFileSpec = ""
    DoEvents
End Sub
    


Private Sub ClearImagesPanelDisp()
    lblPicsOnPanel.Caption = ""
    lblPicsCount.Caption = ""
    
    piccmdPanelRefContainer.Visible = False
    
    picListX.Picture = LoadPicture()
    picListRegionFlag = False
    VsbPicList.Value = 0
    cboPicListLot.Clear
End Sub



Private Sub ClearSearchDisp()
    lblFoundNum.Caption = ""
    lisFilesFound.Clear
End Sub



Private Sub FilList_Click()
    On Error GoTo ErrHandler
    If optMode(1).Value = True Then
         MsgBox "When the List Panel is for Search All Subdir," & vbCrLf & _
            "click file spec listed in it instead (when there" & vbCrLf & _
            "are entries there)."
         Exit Sub
    End If
    
    If cboPicListLot.ListCount < 1 Then
         picListX.Cls                ' Don't leave a region
         Exit Sub
    End If
                 
    Screen.MousePointer = vbHourglass
    Dim i, j
    Dim mPath As String
    Dim mfilespec As String
    If Right(DirList.Path, 1) <> "\" Then
         mPath = DirList.Path & "\"
    Else
         mPath = DirList.Path                   ' e.g. root "C:\"
    End If
    
    HsbView.Value = 0
    VsbView.Value = 0
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
         
    
    If SuspendFlag = False And cmdAutoImagesOn.Visible = True Then
         lisFilesFound.Refresh
          
         If picListRegionFlag Then
             ClearPicListRegion
         End If
         
         If cboPattern.Text <> "*.*" Then      ' We may use this approach also
               i = Int(FilList.ListIndex / xMaxNums) + 1
               If i <> CInt(cboPicListLot.Text) Then
                    cboPicListLot.ListIndex = (i - 1)
               End If
         
               picListSeq = FilList.ListIndex Mod xMaxNums
               picListRow = Int(picListSeq / xColNums) + 1
               picListCol = picListSeq Mod xColNums + 1          ' Remainder, & +1
         
         Else               ' If "*.*" then we must use this approach
              mfilespec = mPath & FilList.List(FilList.ListIndex)
              picListSeq = -1                             ' As a marking
              For i = 0 To UBound(arrPicFileSpec) - 1
                    If arrPicFileSpec(i, 1) = mfilespec Then
                         picListSeq = CInt(arrPicFileSpec(i, 3))
                    End If
              Next i
              If picListSeq = -1 Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Not an image file/invalid file"
                    Exit Sub
              End If
              
              i = Int(picListSeq / xMaxNums) + 1
              If i <> CInt(cboPicListLot.Text) Then
                    cboPicListLot.ListIndex = (i - 1)
              End If
         
              picListSeq = picListSeq Mod xMaxNums
              picListRow = Int(picListSeq / xColNums) + 1
              picListCol = picListSeq Mod xColNums + 1          ' Remainder, & +1
         End If
         
         i = Int(picListSeq / (xColNums * xRowsPerPanel))
         
         VsbPicList.Value = 0
         If i > 0 Then
              j = Int(VsbPicList.Max / (xPanelNums - 1)) * (i)
              If j > VsbPicList.Max Then
                   j = VsbPicList.Max
              End If
              VsbPicList.Value = j
         End If
         DoEvents
         
         DispPicListRegion
    End If
    
    mfilespec = mPath & FilList.List(FilList.ListIndex)
    On Error Resume Next
    Err = False
    picViewX.Picture = LoadPicture()
    picCopy.Picture = LoadPicture()
    picViewX.Picture = LoadPicture(mfilespec)
    If Err Then
         Screen.MousePointer = vbNormal
         MsgBox "Not an image file/invalid image"
         Exit Sub
    End If
    
    picViewXFileSpec = mfilespec
    
    OrigW = picViewX.ScaleWidth
    OrigH = picViewX.ScaleHeight
    lblPicSizeW.Caption = "w=" & OrigW
    lblPicSizeH.Caption = "h=" & OrigH
    
    picCopy.Picture = LoadPicture()
    picCopy.Picture = picViewX.Image
    
    piccmdEditContainer.Visible = True
    cboResize.Visible = True
    piccmdOrigSizeContainer.Visible = True
    
    If OrigW > picViewZ.ScaleWidth Or OrigH > picViewZ.ScaleHeight Then
         piccmdFitInContainer.Visible = True
         piccmdFitInLargeContainer.Visible = True
    Else
         piccmdFitInContainer.Visible = False
         piccmdFitInLargeContainer.Visible = False
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
    ErrMsgProc "FilList_Click"
End Sub




Private Sub ListImagesInDir()
    On Error Resume Next
    If lblSearchInProgress.Visible = True Then
         Exit Sub
    End If
    
    If (DirList.Path = strCurrDir) And (cboPattern = strCurrPattern) Then
         If AutoListSwitching = False Then
               Exit Sub
         End If
    End If
    AutoListSwitching = False
    
    Screen.MousePointer = vbHourglass
    
    ClearImagesPanelDisp
    
    strCurrDir = DirList.Path
    strCurrPattern = cboPattern
    
    Dim mImageCount As Integer
    Dim mPath As String, mFile As String, mFullSpec As String
    Dim mfilespec As String
    Dim i As Integer, j As Integer, k As Integer
    Dim mPercent
    
    cboPicListLot.Clear
    If Right(DirList.Path, 1) <> "\" Then
        mPath = DirList.Path & "\"
    Else
        mPath = DirList.Path                   ' e.g. root "C:\"
    End If

    mFile = Dir(mPath & cboPattern, vbNormal)  ' Retrieve the first entry.
    If mFile = "" Then                 ' Cannot find first
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    
    ReDim arrPicFileSpec(FilList.ListCount - 1, 3) As String
    
    mImageCount = 0
    picProgressContainer.Visible = True
    k = FilList.ListCount - 1
    For i = 0 To k
         mFullSpec = mPath & FilList.List(i)
         mPercent = Int(i / (k - 1)) * 100
         PlotProgress mPercent
         Err = False
         If cboPattern.Text = "*.*" Then
              mfilespec = UCase(Right(mFullSpec, 3))
              If InStr(NonPicFileExt, mfilespec) = 0 Then
                   picTemp.Picture = LoadPicture(mFullSpec)
              Else
                   Err = True
              End If
              If Err = False Then
                  arrPicFileSpec(mImageCount, 1) = mFullSpec
                  arrPicFileSpec(mImageCount, 2) = i
                     ' Store valid seq in third dimension
                  arrPicFileSpec(mImageCount, 3) = mImageCount
                  mImageCount = mImageCount + 1
              End If
         Else
              arrPicFileSpec(mImageCount, 1) = mFullSpec
              arrPicFileSpec(mImageCount, 2) = i
              arrPicFileSpec(mImageCount, 3) = i
              mImageCount = mImageCount + 1
         End If
    Next i
    picProgressContainer.Visible = False
    
    picsCount = mImageCount
    
    If picsCount = 0 Then
        Screen.MousePointer = vbDefault
        ReDim arrPicFileSpec(0)
        Exit Sub
    End If
    
    SuspendFlag = True
    If picsCount <= xMaxNums Then
        cboPicListLot.AddItem 1                 ' Lot 1 is enough
    Else
        j = Int(picsCount / xMaxNums) + 1
        For i = 0 To (j - 1)
             cboPicListLot.AddItem (i + 1)
        Next i
    End If
    cboPicListLot.ListIndex = 0
    SuspendFlag = False
    
    PRINTPICLIST
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub PlotProgress(ByVal inPercent As Integer)
    picProgressContainer.Cls
    picProgress.Cls
    picProgress.Width = picProgressContainer.ScaleWidth * (CInt(inPercent) / 100)
    picProgressContainer.CurrentX = picProgressContainer.ScaleWidth / 2
    picProgressContainer.CurrentY = 1
    picProgress.CurrentX = picProgress.ScaleWidth / 2
    picProgress.CurrentY = 1
    If picProgress.CurrentX < picProgressContainer.CurrentX Then
         picProgressContainer.Print CStr(inPercent) & "%"
    Else
         picProgress.Print CStr(inPercent) & "%"
    End If
    DoEvents
End Sub




Private Sub PRINTPICLIST()
    On Error Resume Next
    picListRegionFlag = False
    picListX.Picture = LoadPicture()
    VsbPicList.Value = 0                  ' Ensure scrollbar back to 0 pos
    
    If cboPicListLot.ListCount = 0 Then
         Exit Sub
    End If
    
    Dim x, y
    Dim W, H
    Dim i, j, k
    Dim mIconNo
    Dim mFile
    
    j = (val(cboPicListLot.Text) - 1) * xMaxNums
    
    picsOnCurrList = picsCount - j
    If picsOnCurrList > xMaxNums Then
         picsOnCurrList = xMaxNums
    End If
    
    lblPicsOnPanel.Caption = "Images in current lot: " & CStr(picsOnCurrList) & _
               "  (max " & CStr(xMaxNums) & " per lot)"
    If cboPicListLot.ListCount = 1 Then
        lblPicsCount.Caption = "Total images: " & CStr(picsCount) & _
              ",  contained in 1 lot"
    Else
        lblPicsCount.Caption = "Total images: " & CStr(picsCount) & _
              ",  contained in " & CStr(cboPicListLot.ListCount) & " lots."
    End If
    
    k = j                                  ' Keep this k value unchanged
    
    x = xGapH / 2
    y = xGapW / 2
    
    For i = j To picsCount - 1
         Err = False
         picAuto.Picture = LoadPicture()
         picAuto.Picture = LoadPicture(arrPicFileSpec(i, 1))
         
         If Err = False Then
             W = picAuto.ScaleWidth
             H = picAuto.ScaleHeight
             picAuto.Refresh
         
             picTemp.Picture = LoadPicture()
             mresult = StretchPic(picTemp, StdW, StdH, picAuto, W, H)
             If mresult = 0 Then
                  mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
             Else
                  mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, picTemp.hdc, 0, 0, vbSrcCopy)
                  If mresult = 0 Then
                       mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
                  Else
                       mFile = PathSection(arrPicFileSpec(i, 1), 3)
                       picListX.CurrentX = (x - 7)        ' So that text somewhat in middle of pic
                       picListX.CurrentY = (y + StdH + 2)
                       If Len(mFile) > 8 Then
                           picListX.Print Left(mFile, 8)
                       Else
                           picListX.Print mFile
                       End If
                  End If
             End If
        Else
             mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
        End If
        
        picListX.Refresh
        
        If (i - k) > (xMaxNums - 1) Then
             Exit For
        End If
         
        x = x + (StdW + xGapW)
        If x > (StdW + xGapW) * xColNums Then
             x = xGapW / 2                        ' Start from first column again
             y = y + (StdH + xTotalGapH)          ' An icon's size plus the gap
        End If
    Next i
    DoEvents
    
    piccmdPanelRefContainer.Visible = True
    j = xMaxNums / xPanelNums
    For i = 0 To j - 1
        cmdPanelRef(i).Enabled = False
    Next i
    j = Int(picsOnCurrList / (xColNums * xRowsPerPanel))
    For i = 0 To j
        cmdPanelRef(i).Enabled = True
    Next i
End Sub



Private Function StretchPic(inDestPic, newW, newH, inSrcPic, W, H)
    On Error GoTo ErrHandler
    Dim mOrigTone       As Long
    inDestPic.Picture = LoadPicture()
    mOrigTone = SetStretchBltMode(inDestPic.hdc, STRETCH_HALFTONE)
    StretchBlt inDestPic.hdc, 0, 0, newW, newH, inSrcPic.hdc, 0, 0, W, H, vbSrcCopy
    inDestPic.Picture = inDestPic.Image
    mresult = SetStretchBltMode(inDestPic.hdc, mOrigTone)
    StretchPic = 1
    Exit Function
    
ErrHandler:
    StretchPic = 0
End Function



Private Sub GetpicSeq(inX As Single, inY As Single)
    Dim a As Single, B As Single
    a = picListX.ScaleWidth / xColNums       ' i.e. /3
    B = picListX.ScaleHeight / xRowNums      '      /14
    For picListCol = 1 To xColNums
        If inX < (a * picListCol) Then
             Exit For
        End If
    Next picListCol
    For picListRow = 1 To xRowNums
        If inY < (B * picListRow) Then
             Exit For
        End If
    Next picListRow
    
    picListSeq = (picListRow - 1) * xColNums + picListCol
    If picListSeq > picsOnCurrList Then
        picListSeq = 0
    End If
End Sub




Private Sub DispPicListRegion()
    If cboPicListLot.ListCount = 0 Then
         picListX.Cls
         Exit Sub
    End If
    
    Dim xStart, yStart
    xStart = (picListCol - 1) * (StdW + xGapH) + 16        ' We start from 16
    yStart = (picListRow - 1) * (StdH + xTotalGapH) + 16
    
    OldX1 = xStart - 11
    OldX2 = xStart + StdW + 11
    OldY1 = yStart - 2
    OldY2 = yStart + StdH + xTextH + 2
    picListRegionFlag = True
    
    picListX.DrawMode = vbInvert
    picListX.DrawStyle = vbDot
    picListX.Line (OldX1, OldY1)-(OldX2, OldY2), , B
    picListX.DrawMode = vbCopyPen
    picListX.DrawStyle = vbSolid
End Sub



Private Sub ClearPicListRegion()
    picListRegionFlag = False
    picListX.DrawMode = vbInvert
    picListX.DrawStyle = vbDot
    picListX.Line (OldX1, OldY1)-(OldX2, OldY2), , B
    picListX.DrawMode = vbCopyPen
    picListX.DrawStyle = vbSolid
End Sub



Private Sub cbopicListLot_Click()
    If SuspendFlag = False Then
        If val(cboPicListLot.Text) > 0 Then
               PRINTPICLIST
         End If
    End If
End Sub



Private Sub cmdPanelRef_Click(Index As Integer)
    If cboPicListLot.ListCount < 1 Then
         Exit Sub
    End If
    VsbPicList.Value = Int(VsbPicList.Max / (xPanelNums - 1)) * Index
    picListX.SetFocus
End Sub



Private Sub cmdSearch_Click()
    If Trim(cboSearchPattern.Text) = "" Then
        MsgBox "No search pattern yet"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    lisFilesFound.Clear
    
    StopSearchFlag = False
    SuspendFlag = True     ' Avoid updates in picListX resulting from change of dir
    lisFilesFound.SetFocus
    
    ToggleSearchVisibles False
      
    Dim mFirstPath As String
    Dim mErrDirDiver As Boolean
    Dim mDirCount As Integer
    Dim mNumFiles As Integer
    
    If DirList.Path <> DirList.List(DirList.ListIndex) Then
         DirList.Path = DirList.List(DirList.ListIndex)
         SuspendFlag = False
         Screen.MousePointer = vbDefault
         Exit Sub         ' Exit so user can take a look before re-search
    End If

    FilList.Pattern = cboSearchPattern.Text

    mFirstPath = DirList.Path
    mDirCount = DirList.ListCount

    filesCount = 0                     ' Reset found files indicator
    mErrDirDiver = DirDiver(mFirstPath, mDirCount, "")
    
    If StopSearchFlag Then
         strCurrDir = ""
         SuspendFlag = False
         ToggleSearchVisibles True
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    
    If mErrDirDiver = True Then
         lisFilesFound.Clear
         filesCount = 0
         DirList.Path = CurDir
         drvList.Drive = DirList.Path        ' Reset the path.
         SuspendFlag = False
         ToggleSearchVisibles True
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    If filesCount > 0 Then
        lblFoundNum.Caption = "Total found: " & CStr(filesCount)
    Else
        lblFoundNum.Caption = "Total found: 0"
    End If
    FilList.Path = DirList.Path
     
    ToggleSearchVisibles True
    
    If cboPattern.Text <> cboSearchPattern.Text Then
        DirList_Change
    End If
       
    Screen.MousePointer = vbDefault
    If lisFilesFound.ListCount = 0 Then
        MsgBox "No file found matching the search pattern"
    End If
End Sub




Private Function DirDiver(NewPath As String, mDirCount As Integer, BackUp As String) As Integer
    If StopSearchFlag Then
        Exit Function
    End If

    Dim mDirToPeek As Integer
    Dim mAbandon As Integer
    Dim mOldPath As String
    Dim mCurrPath As String
    Dim mEntry As String
    Dim mRetVal As Integer
    Dim i As Integer
    
    DirDiver = False             ' Assumed first. Set to False if there is an error.
    
    mRetVal = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If StopSearchFlag Then
        DirDiver = True
        Exit Function
    End If
    
    On Local Error GoTo ErrHandler:
    
    mDirToPeek = DirList.ListCount    ' How many directories below this?
    
    Do While mDirToPeek > 0 And StopSearchFlag = False
        mOldPath = DirList.Path                      ' Save old path for next recursion.
        DirList.Path = NewPath
        If DirList.ListCount > 0 Then
            DirList.Path = DirList.List(mDirToPeek - 1)
            mAbandon = DirDiver((DirList.Path), mDirCount%, mOldPath)
        End If
        mDirToPeek = mDirToPeek - 1
        If mAbandon = True Then
            StopSearchFlag = True
            Exit Function
        End If
    Loop
    
    If FilList.ListCount Then
        If Len(DirList.Path) <= 3 Then             ' Check for 2 bytes/character
            mCurrPath = DirList.Path                  ' If at root level, leave as is...
        Else
            mCurrPath = DirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For i = 0 To FilList.ListCount - 1        ' Add conforming files in this directory to the list box.
            mEntry = mCurrPath + FilList.List(i)
            lisFilesFound.AddItem mEntry
            filesCount = filesCount + 1
            lblFoundNum = "Total found: " & CStr(filesCount)
       Next i
    End If
    If BackUp <> "" Then                ' If there is a superior directory, move it.
        DirList.Path = BackUp
    End If
    Exit Function
    
ErrHandler:
    lblSearchInProgress.Visible = False
    cmdStopSearch.Visible = False
    lblFoundNum.Visible = True
    If Err = 7 Then              ' If Out of Memory error occurs, assume the list box just got full.
        MsgBox "Out of memory. Abandoning search..."
    Else                         ' Otherwise display error message and quit.
        ErrMsgProc "frmPicViewer DirDiver"
    End If
End Function




Private Sub cmdStopSearch_Click()
    lisFilesFound.SetFocus
    StopSearchFlag = True
    strCurrDir = ""
End Sub



Private Sub lisFilesFound_dblClick()
    If lisFilesFound.ListCount > 0 Then
        MsgBox "File specification:" & vbCrLf & vbCrLf & lisFilesFound.Text
    End If
End Sub



Private Sub lisFilesFound_Click()
    On Error Resume Next
    If picListRegionFlag Then
         ClearPicListRegion
         picListRegionFlag = False
    End If
    FilList.Refresh
    HsbView.Value = 0
    VsbView.Value = 0
    Err = False
    picViewX.Picture = LoadPicture(lisFilesFound.Text)
    If Err Then
         MsgBox "Not an image file/invalid image"
         Exit Sub
    End If
    
    picViewXFileSpec = lisFilesFound.Text
    
    lblPicSizeW.Caption = "w=" & picViewX.ScaleWidth
    lblPicSizeH.Caption = "h=" & picViewX.ScaleHeight
    
    OrigW = picViewX.ScaleWidth
    OrigH = picViewX.ScaleHeight
    
    picCopy.Picture = LoadPicture()
    picCopy.Picture = picViewX.Image
    picCopy.Picture = picViewX.Image
    
    piccmdEditContainer.Visible = True
    cboResize.Visible = True
    piccmdOrigSizeContainer.Visible = True
    
    If OrigW > picViewZ.ScaleWidth Or OrigH > picViewZ.ScaleHeight Then
         piccmdFitInContainer.Visible = True
         piccmdFitInLargeContainer.Visible = True
    Else
         piccmdFitInContainer.Visible = False
         piccmdFitInLargeContainer.Visible = False
    End If
End Sub




Private Sub ToggleSearchVisibles(OnOff As Boolean)
    lblSearchInProgress.Visible = Not OnOff
    cmdStopSearch.Visible = Not OnOff
    lblFoundNum.Visible = OnOff
    
    Shape1.Visible = OnOff
    cboPattern.Visible = OnOff
    drvList.Visible = OnOff
    DirList.Visible = OnOff
    FilList.Visible = OnOff
    lblFilesCount.Visible = OnOff
End Sub



Private Sub StartLockWindow(ByVal lHWnd As Long)
    mresult = LockWindowUpdate(lHWnd)
End Sub



Private Sub StopLockWindow()
    mresult = LockWindowUpdate(0)
End Sub



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub


Function PathSection(ByVal inPath As String, inReturnType As Integer)
    Dim DriveLetter As String
    Dim DirPath As String
    Dim FName As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If inReturnType <> 0 And inReturnType <> 1 And inReturnType <> 2 And inReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If
    DriveLetter = ""
    DirPath = ""
    FName = ""
    Extension = ""
    If Mid(inPath, 2, 1) = ":" Then
        DriveLetter = Left(inPath, 2)
        inPath = Mid(inPath, 3)
    End If
    PathLength = Len(inPath)
    For Offset = PathLength To 1 Step -1
        Select Case Mid(inPath, Offset, 1)
            Case ".":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                Extension = Mid(inPath, Offset, ThisLength + 1)
            End If
            inPath = Left(inPath, Offset - 1)
            Case "\":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                FName = Mid(inPath, Offset + 1, ThisLength)
                inPath = Left(inPath, Offset)
                FileNameFound = True
                Exit For
            End If
            Case Else
        End Select
    Next Offset
    If FileNameFound = False Then
        FName = inPath
    Else
        DirPath = inPath
    End If
    If inReturnType = 0 Then
        PathSection = DriveLetter
    ElseIf inReturnType = 1 Then
        PathSection = DirPath
    ElseIf inReturnType = 2 Then
        PathSection = FName & Extension
    ElseIf inReturnType = 3 Then
        PathSection = FName
    ElseIf inReturnType = 4 Then
        PathSection = Extension
    End If
End Function

