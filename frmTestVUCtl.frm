VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTestVUCtl 
   BackColor       =   &H8000000A&
   Caption         =   "VU Demo"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Lights"
      TabPicture(0)   =   "frmTestVUCtl.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "fraLights"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bars"
      TabPicture(1)   =   "frmTestVUCtl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBars"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Gradient"
      TabPicture(2)   =   "frmTestVUCtl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraGrad"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Meters"
      TabPicture(3)   =   "frmTestVUCtl.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraMeters"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraMeters 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   8175
         Begin VB.ComboBox cboSep 
            Height          =   315
            ItemData        =   "frmTestVUCtl.frx":0070
            Left            =   5640
            List            =   "frmTestVUCtl.frx":0072
            TabIndex        =   67
            Top             =   6120
            Width           =   1455
         End
         Begin VB.ComboBox cboTh 
            Height          =   315
            ItemData        =   "frmTestVUCtl.frx":0074
            Left            =   3960
            List            =   "frmTestVUCtl.frx":0076
            TabIndex        =   65
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Frame Frame7 
            Caption         =   "Colors"
            Height          =   855
            Left            =   1680
            TabIndex        =   62
            Top             =   5640
            Width           =   2175
            Begin VB.OptionButton optMCol 
               Caption         =   "Meter3 BackColor"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   64
               Top             =   480
               Width           =   1815
            End
            Begin VB.OptionButton optMCol 
               Caption         =   "Needle Color"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Orientation"
            Height          =   855
            Left            =   0
            TabIndex        =   59
            Top             =   5640
            Width           =   1455
            Begin VB.OptionButton optOrM 
               Caption         =   "Vertical"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   61
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton optOrM 
               Caption         =   "Horizontal"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin TestVUCtl.Meter Meter 
            Height          =   1365
            Index           =   3
            Left            =   4080
            TabIndex        =   58
            Top             =   2640
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   2408
            Style           =   3
         End
         Begin TestVUCtl.Meter Meter 
            Height          =   1380
            Index           =   2
            Left            =   240
            TabIndex        =   57
            Top             =   2640
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   2434
            Style           =   2
            NeedleColor     =   12632256
         End
         Begin TestVUCtl.Meter Meter 
            Height          =   1140
            Index           =   1
            Left            =   4080
            TabIndex        =   56
            Top             =   120
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   2011
            Style           =   1
            NeedleColor     =   128
         End
         Begin TestVUCtl.Meter Meter 
            Height          =   1020
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   1799
            NeedleColor     =   12632256
         End
         Begin VB.CommandButton cmdStopM 
            Caption         =   "Stop"
            Height          =   255
            Left            =   7320
            TabIndex        =   54
            Top             =   6240
            Width           =   855
         End
         Begin VB.CommandButton cmdStartM 
            Caption         =   "Start"
            Height          =   255
            Left            =   7320
            TabIndex        =   53
            Top             =   5880
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Meter Separation"
            Height          =   255
            Left            =   5640
            TabIndex        =   68
            Top             =   5880
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Needle Thickness"
            Height          =   255
            Left            =   3960
            TabIndex        =   66
            Top             =   5880
            Width           =   1455
         End
      End
      Begin VB.Frame fraGrad 
         BorderStyle     =   0  'None
         Caption         =   "Gradient"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   8055
         Begin VB.CommandButton cmdPrevG 
            Caption         =   "Preview"
            Height          =   255
            Left            =   2160
            TabIndex        =   73
            Top             =   3720
            Width           =   735
         End
         Begin VB.Frame Frame8 
            Caption         =   "Colors"
            Height          =   1215
            Left            =   120
            TabIndex        =   48
            Top             =   3600
            Width           =   1455
            Begin VB.OptionButton optGCol 
               Caption         =   "Color 1"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optGCol 
               Caption         =   "BackColor"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   50
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optGCol 
               Caption         =   "Color 2"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   49
               Top             =   840
               Width           =   1215
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Size"
            Height          =   1215
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   1455
            Begin VB.OptionButton optSizeG 
               Caption         =   "Small"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optSizeG 
               Caption         =   "Medium"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   46
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton optSizeG 
               Caption         =   "Large"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   45
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdStartG 
            Caption         =   "Start"
            Height          =   255
            Left            =   2040
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdStopG 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   42
            Top             =   2160
            Width           =   855
         End
         Begin VB.Frame Frame10 
            Caption         =   "Orientation"
            Height          =   855
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   1455
            Begin VB.OptionButton optOrG 
               Caption         =   "Horizontal"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optOrG 
               Caption         =   "Vertical"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   1215
            End
         End
         Begin TestVUCtl.GradV GradVR 
            Height          =   1335
            Left            =   4680
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2355
            Channel         =   1
            BackColor       =   0
         End
         Begin TestVUCtl.GradV GradVL 
            Height          =   1335
            Left            =   4560
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2355
            BackColor       =   0
         End
         Begin TestVUCtl.GradH GradHR 
            Height          =   90
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   159
            Channel         =   1
            BackColor       =   0
         End
         Begin TestVUCtl.GradH GradHL 
            Height          =   90
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   159
            BackColor       =   0
         End
         Begin VB.Label Label5 
            Caption         =   "Note Gradients don't show anything at run time if no sound is coming from the sound card"
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   6240
            Width           =   6735
         End
      End
      Begin VB.Frame fraBars 
         BorderStyle     =   0  'None
         Caption         =   "Bars"
         Height          =   5895
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   5655
         Begin VB.CommandButton cmdPrevB 
            Caption         =   "Preview"
            Height          =   255
            Left            =   2160
            TabIndex        =   72
            Top             =   3720
            Width           =   735
         End
         Begin VB.Frame Frame3 
            Caption         =   "Orientation"
            Height          =   855
            Left            =   120
            TabIndex        =   30
            Top             =   2640
            Width           =   1455
            Begin VB.OptionButton optOrB 
               Caption         =   "Vertical"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton optOrB 
               Caption         =   "Horizontal"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdStopB 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdStartB 
            Caption         =   "Start"
            Height          =   255
            Left            =   2040
            TabIndex        =   28
            Top             =   1800
            Width           =   855
         End
         Begin VB.Frame Frame5 
            Caption         =   "Size"
            Height          =   1215
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   1455
            Begin VB.OptionButton optSizeB 
               Caption         =   "Large"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   840
               Width           =   975
            End
            Begin VB.OptionButton optSizeB 
               Caption         =   "Medium"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton optSizeB 
               Caption         =   "Small"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Colors"
            Height          =   975
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "After selection, Choose a color from the palette"
            Top             =   3600
            Width           =   1455
            Begin VB.OptionButton optBCol 
               Caption         =   "BackColor"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optBCol 
               Caption         =   "ForeColor"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   1215
            End
         End
         Begin TestVUCtl.BarsH BarsHL 
            Height          =   90
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   159
         End
         Begin TestVUCtl.BarsV BarsVR 
            Height          =   1335
            Left            =   4680
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2355
            Channel         =   1
         End
         Begin TestVUCtl.BarsV BarsVL 
            Height          =   1335
            Left            =   4560
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2355
         End
         Begin TestVUCtl.BarsH BarsHR 
            Height          =   90
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   159
            Channel         =   1
         End
      End
      Begin VB.Frame fraLights 
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   5655
         Begin VB.Frame Frame1 
            Caption         =   "Size"
            Height          =   1215
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1455
            Begin VB.OptionButton optSizeL 
               Caption         =   "Small"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optSizeL 
               Caption         =   "Medium"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton optSizeL 
               Caption         =   "Large"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdStartL 
            Caption         =   "Start"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdStopL 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            Top             =   2160
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Caption         =   "Orientation"
            Height          =   855
            Left            =   120
            TabIndex        =   4
            Top             =   2640
            Width           =   1455
            Begin VB.OptionButton optOrL 
               Caption         =   "Horizontal"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optOrL 
               Caption         =   "Vertical"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Width           =   1215
            End
         End
         Begin TestVUCtl.LightsV LightsVL 
            Height          =   1350
            Left            =   4560
            TabIndex        =   7
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2381
         End
         Begin TestVUCtl.LightsH LightsHL 
            Height          =   90
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   159
         End
         Begin TestVUCtl.LightsH LightsHR 
            Height          =   90
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   159
            Channel         =   1
         End
         Begin TestVUCtl.LightsV LightsVR 
            Height          =   1350
            Left            =   4680
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   90
            _ExtentX        =   159
            _ExtentY        =   2381
         End
      End
      Begin VB.Label Label4 
         Caption         =   "No color selection here.  What you see is what you get."
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   6720
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Note Bars don't show anything at run time if no sound is coming from the sound card"
         Height          =   255
         Left            =   -74640
         TabIndex        =   69
         Top             =   6720
         Width           =   6615
      End
   End
   Begin VB.PictureBox pPal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   9000
      Picture         =   "frmTestVUCtl.frx":0078
      ScaleHeight     =   155
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   0
      Top             =   600
      Width           =   2385
   End
   Begin VB.Label lblClick 
      Alignment       =   2  'Center
      Caption         =   "Click here for color selection"
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmTestVUCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note this does little if there is no
'sound coming from the soundcard
'So start your media player/cd/mike/whatever
Private TheColor As Long
Private Enum eSelVU
 Bars
 Grad
 Meter3Needle
End Enum
Private SelVU As eSelVU

Private Sub Form_Load()
 Dim i As Long
 Dim x As Long, y As Long
 'the ppal.picture is just a copy from
 ' ms paint's new color definition
 ' this puts the basic colors in
 For i = 0 To 15
  pPal.Line (x, 145)-(x + 9, 155), QBColor(i), BF
  x = x + 10
 Next
 For i = 1 To 5 'meter thumbwidths
  cboTh.AddItem i
 Next
 cboTh.ListIndex = 0
 For i = 0 To 600 Step 30 'meter separation
  cboSep.AddItem i
 Next
 cboSep.ListIndex = 0
 SSTab1.Tab = 0
 SelVU = Bars
End Sub
Private Sub cboSep_Click()
 Dim i As Long
 For i = 0 To 3
  Meter(i).Separation = cboSep.Text
 Next
 Refresh
End Sub

Private Sub cboTh_Click()
 Dim i As Long
 For i = 0 To 3
  Meter(i).NeedleWidth = cboTh.Text
 Next
 Refresh
End Sub
Private Sub cmdStartM_Click()
 Dim i As Long
 For i = 0 To 3
  Meter(i).StartVU
 Next
End Sub
Private Sub cmdStopM_Click()
 Dim i As Long
 For i = 0 To 3
  Meter(i).StopVU
 Next
End Sub
Private Sub cmdStartB_Click()
 BarsHL.StartVU
 BarsHR.StartVU
 BarsVL.StartVU
 BarsVR.StartVU
 cmdStartB.Enabled = False
 cmdPrevB.Enabled = False
 cmdStopB.Enabled = True
End Sub
Private Sub cmdStopB_Click()
 BarsHL.StopVU
 BarsHR.StopVU
 BarsVL.StopVU
 BarsVR.StopVU
 cmdStopB.Enabled = False
 cmdStartB.Enabled = True
 cmdPrevB.Enabled = True
End Sub
Private Sub cmdPrevB_Click()
 BarsHL.Preview
 BarsHR.Preview
 BarsVL.Preview
 BarsVR.Preview
End Sub
Private Sub cmdStartG_Click()
 GradHL.StartVU
 GradHR.StartVU
 GradVL.StartVU
 GradVR.StartVU
 cmdStartG.Enabled = False
 cmdPrevG.Enabled = False
 cmdStopG.Enabled = True
End Sub
Private Sub cmdStopG_Click()
 GradHL.StopVU
 GradHR.StopVU
 GradVL.StopVU
 GradVR.StopVU
 cmdStartG.Enabled = True
 cmdPrevG.Enabled = True
 cmdStopG.Enabled = False
End Sub
Private Sub cmdPrevG_Click()
 GradHL.Preview
 GradHR.Preview
 GradVL.Preview
 GradVR.Preview
End Sub

Private Sub cmdStartL_Click()
 LightsHL.StartVU
 LightsHR.StartVU
 LightsVL.StartVU
 LightsVR.StartVU
 cmdStartL.Enabled = False
 cmdStopL.Enabled = True
End Sub
Private Sub cmdStopL_Click()
 LightsHL.StopVU
 LightsHR.StopVU
 LightsVL.StopVU
 LightsVR.StopVU
 cmdStopL.Enabled = False
 cmdStartL.Enabled = True
End Sub

Private Sub optBCol_Click(index As Integer)
 SelVU = Bars
 TheColor = index
End Sub

Private Sub optGCol_Click(index As Integer)
 SelVU = Grad
 TheColor = index
End Sub
Private Sub optMCol_Click(index As Integer)
 SelVU = Meter3Needle
 TheColor = index
End Sub


Private Sub optOrB_Click(index As Integer)
 Select Case index
  Case 0
   BarsVL.Visible = False
   BarsVR.Visible = False
   BarsHL.Visible = True
   BarsHR.Visible = True
  Case 1
   BarsVL.Visible = True
   BarsVR.Visible = True
   BarsHL.Visible = False
   BarsHR.Visible = False
 End Select

End Sub

Private Sub optOrG_Click(index As Integer)
 Select Case index
  Case 0
   GradVL.Visible = False
   GradVR.Visible = False
   GradHL.Visible = True
   GradHR.Visible = True
  Case 1
   GradVL.Visible = True
   GradVR.Visible = True
   GradHL.Visible = False
   GradHR.Visible = False
 End Select

End Sub

Private Sub optOrL_Click(index As Integer)
 Select Case index
  Case 0
   LightsVL.Visible = False
   LightsVR.Visible = False
   LightsHL.Visible = True
   LightsHR.Visible = True
  Case 1
   LightsVL.Visible = True
   LightsVR.Visible = True
   LightsHL.Visible = False
   LightsHR.Visible = False
 End Select
End Sub

Private Sub optOrM_Click(index As Integer)
 Dim i As Long
 For i = 0 To 3
  Meter(i).Orientation = index
 Next
 Refresh
End Sub

Private Sub optSizeB_Click(index As Integer)
 If optOrB(0).Value = True Then
  BarsVL.VUSize = index
  BarsVR.VUSize = index
  BarsVR.Move BarsVL.Left + BarsVL.Width + 30, BarsVL.Top
  BarsHL.VUSize = index
  BarsHR.VUSize = index
  BarsHR.Move BarsHL.Left, BarsHL.Top + BarsHL.Height + 30
 Else
  BarsHL.VUSize = index
  BarsHR.VUSize = index
  BarsHR.Move BarsHL.Left, BarsHL.Top + BarsHL.Height + 30
  BarsVL.VUSize = index
  BarsVR.VUSize = index
  BarsVR.Move BarsVL.Left + BarsVL.Width + 30, BarsVL.Top
 End If
End Sub

Private Sub optSizeG_Click(index As Integer)
 If optOrG(0).Value = True Then
  GradVL.VUSize = index
  GradVR.VUSize = index
  GradVR.Move GradVL.Left + GradVL.Width + 30, GradVL.Top
  GradHL.VUSize = index
  GradHR.VUSize = index
  GradHR.Move GradHL.Left, GradHL.Top + GradHL.Height + 30
 Else
  GradHL.VUSize = index
  GradHR.VUSize = index
  GradHR.Move GradHL.Left, GradHL.Top + GradHL.Height + 30
  GradVL.VUSize = index
  GradVR.VUSize = index
  GradVR.Move GradVL.Left + GradVL.Width + 30, GradVL.Top
 End If
End Sub

Private Sub optSizeL_Click(index As Integer)
 If optOrL(0).Value = True Then
  LightsVL.VUSize = index
  LightsVR.VUSize = index
  LightsVR.Move LightsVL.Left + LightsVL.Width + 30, LightsVL.Top
  LightsHL.VUSize = index
  LightsHR.VUSize = index
  LightsHR.Move LightsHL.Left, LightsHL.Top + LightsHL.Height + 30
 Else
  LightsHL.VUSize = index
  LightsHR.VUSize = index
  LightsHR.Move LightsHL.Left, LightsHL.Top + LightsHL.Height + 30
  LightsVL.VUSize = index
  LightsVR.VUSize = index
  LightsVR.Move LightsVL.Left + LightsVL.Width + 30, LightsVL.Top
 End If
End Sub
Private Sub optSizeBL_Click(index As Integer)
 Select Case index
  Case 0
   BarsHL.VUSize = SmallLH
   BarsHR.VUSize = SmallLH
  Case 1
   BarsHL.VUSize = MediumLH
   BarsHR.VUSize = MediumLH
  Case 2
   BarsHL.VUSize = LargeLH
   BarsHR.VUSize = LargeLH
 End Select
 BarsHR.Move BarsHL.Left, BarsHL.Top + BarsHL.Height + 30
End Sub

Private Sub pPal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 SetColor pPal.Point(x, y)
End Sub
Private Sub SetColor(ByVal NewCol As Long)
 Dim i As Long
 Select Case SelVU
  Case Grad
   Select Case TheColor
    Case 0 'bc
     GradHL.BackColor = NewCol
     GradHR.BackColor = NewCol
     GradVL.BackColor = NewCol
     GradVR.BackColor = NewCol
    Case 1 'c1
     GradHL.GradientColor1 = NewCol
     GradHR.GradientColor1 = NewCol
     GradVL.GradientColor1 = NewCol
     GradVR.GradientColor1 = NewCol
    Case 2 'c2
     GradHL.GradientColor2 = NewCol
     GradHR.GradientColor2 = NewCol
     GradVL.GradientColor2 = NewCol
     GradVR.GradientColor2 = NewCol
   End Select
   cmdPrevG_Click
  Case Bars
   Select Case TheColor
    Case 0 'bc
     BarsHL.BackColor = NewCol
     BarsHR.BackColor = NewCol
     BarsVL.BackColor = NewCol
     BarsVR.BackColor = NewCol
    Case 1 'fc
     BarsHL.ForeColor = NewCol
     BarsHR.ForeColor = NewCol
     BarsVL.ForeColor = NewCol
     BarsVR.ForeColor = NewCol
   End Select
   cmdPrevB_Click
  Case Meter3Needle
   Select Case TheColor
    Case 0 'needle
     For i = 0 To 3
      Meter(i).NeedleColor = NewCol
     Next
    Case 1 'm3bc
     Meter(2).Meter3BackColor = NewCol
   End Select
 End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 Select Case SSTab1.Tab
  Case 1
   SelVU = Bars
  Case 2
   SelVU = Grad
  Case 3
   SelVU = Meter3Needle
 End Select
End Sub

