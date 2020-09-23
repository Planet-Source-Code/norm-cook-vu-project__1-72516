VERSION 5.00
Begin VB.UserControl Meter 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   3.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6045
   ScaleWidth      =   4410
   ToolboxBitmap   =   "Meter.ctx":0000
   Begin VB.PictureBox pMeter4R 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   2160
      Picture         =   "Meter.ctx":0312
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
      Begin VB.Line linMeter4R 
         BorderColor     =   &H00000000&
         X1              =   67
         X2              =   19
         Y1              =   96
         Y2              =   55
      End
   End
   Begin VB.PictureBox pMeter4L 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   0
      Picture         =   "Meter.ctx":95C8
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   6
      Top             =   4560
      Width           =   2055
      Begin VB.Line linMeter4L 
         BorderColor     =   &H00000000&
         X1              =   67
         X2              =   19
         Y1              =   96
         Y2              =   55
      End
   End
   Begin VB.PictureBox pMeter3R 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   1800
      Picture         =   "Meter.ctx":1287E
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   5
      Top             =   3120
      Width           =   1380
      Begin VB.Line linMeter3R 
         BorderColor     =   &H00C0C0C0&
         X1              =   46
         X2              =   12
         Y1              =   74
         Y2              =   40
      End
   End
   Begin VB.PictureBox pMeter3L 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   120
      Picture         =   "Meter.ctx":18BF0
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   4
      Top             =   3120
      Width           =   1380
      Begin VB.Line linMeter3L 
         BorderColor     =   &H00C0C0C0&
         X1              =   46
         X2              =   12
         Y1              =   73
         Y2              =   40
      End
   End
   Begin VB.PictureBox pMeter2R 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   2040
      Picture         =   "Meter.ctx":1EF62
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   3
      Top             =   1800
      Width           =   1845
      Begin VB.Line linMeter2R 
         BorderColor     =   &H00000080&
         X1              =   59
         X2              =   20
         Y1              =   88
         Y2              =   33
      End
   End
   Begin VB.PictureBox pMeter2L 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   120
      Picture         =   "Meter.ctx":25E14
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   2
      Top             =   1800
      Width           =   1845
      Begin VB.Line linMeter2L 
         BorderColor     =   &H00000080&
         X1              =   59
         X2              =   20
         Y1              =   88
         Y2              =   33
      End
   End
   Begin VB.PictureBox pMeter1R 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   1920
      Picture         =   "Meter.ctx":2CCC6
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      Begin VB.Line linMeter1R 
         BorderColor     =   &H00C0C0C0&
         X1              =   52
         X2              =   20
         Y1              =   63
         Y2              =   30
      End
   End
   Begin VB.PictureBox pMeter1L 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   240
      Picture         =   "Meter.ctx":320F8
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      Begin VB.Line linMeter1L 
         BorderColor     =   &H00C0C0C0&
         X1              =   52
         X2              =   20
         Y1              =   63
         Y2              =   30
      End
   End
End
Attribute VB_Name = "Meter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eStyle
 Meter1
 Meter2
 Meter3
 Meter4
End Enum
Public Enum eOrientation
 Horizontal
 Vertical
End Enum

Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private intSamples() As Integer
Private vM1 As Variant
Private vM2 As Variant
Private vM3 As Variant
Private vM4 As Variant
Private mStyle As eStyle
Private mSeparation As Long
Private mOrientation As eOrientation
Private mMeter3BackColor As OLE_COLOR
Public Property Get NeedleColor() As OLE_COLOR
 Select Case mStyle
  Case Meter1
   NeedleColor = linMeter1L.BorderColor
  Case Meter2
   NeedleColor = linMeter2L.BorderColor
  Case Meter3
   NeedleColor = linMeter3L.BorderColor
  Case Meter4
   NeedleColor = linMeter4L.BorderColor
 End Select
End Property
Public Property Let NeedleColor(ByVal NewCol As OLE_COLOR)
 Select Case mStyle
  Case Meter1
   linMeter1L.BorderColor = NewCol
   linMeter1R.BorderColor = NewCol
  Case Meter2
   linMeter2L.BorderColor = NewCol
   linMeter2R.BorderColor = NewCol
  Case Meter3
   linMeter3L.BorderColor = NewCol
   linMeter3R.BorderColor = NewCol
  Case Meter4
   linMeter4L.BorderColor = NewCol
   linMeter4R.BorderColor = NewCol
 End Select
End Property
Public Property Get NeedleWidth() As Long
 Select Case mStyle
  Case Meter1
   NeedleWidth = linMeter1L.BorderWidth
  Case Meter2
   NeedleWidth = linMeter2L.BorderWidth
  Case Meter3
   NeedleWidth = linMeter3L.BorderWidth
  Case Meter4
   NeedleWidth = linMeter4L.BorderWidth
 End Select
End Property
Public Property Let NeedleWidth(ByVal NewVal As Long)
 Select Case mStyle
  Case Meter1
   linMeter1L.BorderWidth = NewVal
   linMeter1R.BorderWidth = NewVal
  Case Meter2
   linMeter2L.BorderWidth = NewVal
   linMeter2R.BorderWidth = NewVal
  Case Meter3
   linMeter3L.BorderWidth = NewVal
   linMeter3R.BorderWidth = NewVal
  Case Meter4
   linMeter4L.BorderWidth = NewVal
   linMeter4R.BorderWidth = NewVal
 End Select
End Property
Public Property Get BackColor() As OLE_COLOR
 BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewCol As OLE_COLOR)
 UserControl.BackColor = NewCol
End Property
Public Property Get Separation() As Long
 Separation = mSeparation
End Property
Public Property Let Separation(ByVal NewVal As Long)
 mSeparation = NewVal
 UserControl_ReSize
End Property
Public Property Get Orientation() As eOrientation
 Orientation = mOrientation
End Property
Public Property Let Orientation(ByVal NewVal As eOrientation)
 mOrientation = NewVal
 UserControl_ReSize
End Property
'Allow user to change Meter3's backcolor
Public Property Get Meter3BackColor() As OLE_COLOR
 Meter3BackColor = mMeter3BackColor
End Property
Public Property Let Meter3BackColor(ByVal NewCol As OLE_COLOR)
 If mStyle = Meter3 Then
  mMeter3BackColor = NewCol
  M3BackColor
 End If
End Property
Public Property Get Style() As eStyle
 Style = mStyle
End Property
Public Property Let Style(ByVal NewVal As eStyle)
 mStyle = NewVal
 UserControl_ReSize
 PropertyChanged "Style"
End Property
Public Sub StartVU()
 If Not oRec.IsRecording Then
  oRec.StartRecord 44100, 2
 End If
End Sub
Public Sub StopVU()
 Graphics 0, 0
 oRec.StopRecord
 Graphics 0, 0
End Sub
'Note no timer needed since the event below
'is called when the WinProc receives MM_WIM_DATA
Private Sub oRec_GotData(intBuffer() As Integer, lngLen As Long)
 Dim lngMaxL As Long, lngMaxR As Long
 intSamples = intBuffer
 lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
 lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
 Graphics lngMaxL / 32768#, lngMaxR / 32768#
End Sub

'================Worker functions=============
Private Function GetArrayMaxAbs(intArray() As Integer, _
    Optional ByVal offStart As Long = 0, _
    Optional ByVal steps As Long = 1) As Long
 Dim lngTemp As Long
 Dim lngMax  As Long
 Dim i       As Long
 For i = offStart To UBound(intArray) Step steps
  lngTemp = Abs(CLng(intArray(i)))
  If lngTemp > lngMax Then
   lngMax = lngTemp
  End If
 Next
 If lngMax = 0 Then lngMax = 1
 GetArrayMaxAbs = lngMax
End Function
'Since Meter3 is rounded, allows user to
' change its backcolor to match form
'Note ExtFloodFill requires pic.fillstyle=vbfssolid=0
Private Sub M3BackColor()
 With pMeter3L
  .FillColor = mMeter3BackColor 'color to fill with
  ExtFloodFill .hdc, 0, 0, .Point(0, 0), FLOODFILLSURFACE
  ExtFloodFill .hdc, 91, 0, .Point(91, 0), FLOODFILLSURFACE
  ExtFloodFill .hdc, 8, 85, .Point(8, 85), FLOODFILLSURFACE
  ExtFloodFill .hdc, 87, 89, .Point(87, 89), FLOODFILLSURFACE
  .Refresh
 End With
 With pMeter3R
  .FillColor = mMeter3BackColor 'color to fill with
  ExtFloodFill .hdc, 0, 0, .Point(0, 0), FLOODFILLSURFACE
  ExtFloodFill .hdc, 91, 0, .Point(91, 0), FLOODFILLSURFACE
  ExtFloodFill .hdc, 8, 85, .Point(8, 85), FLOODFILLSURFACE
  ExtFloodFill .hdc, 87, 89, .Point(87, 89), FLOODFILLSURFACE
  .Refresh
 End With
End Sub
'LLev & RLev are singles, 0 to 1
' where 0=no sound, 1=max sound
Private Sub Graphics(ByVal LLev As Single, ByVal RLev As Single)
 Dim TopL As Long, TopR As Long
 Static Cnt As Long
 Cnt = Cnt + 1
 If Cnt < 2 Then Exit Sub
 Cnt = 0
 Select Case mStyle
  Case Meter1
   TopL = 35 * LLev '35 is the ubound for vM1
   If (TopL And 1) > 0 Then 'ensure on even boundary
    TopL = TopL + 1
   End If
   TopR = 35 * RLev
   If (TopR And 1) > 0 Then
    TopR = TopR + 1
   End If
   If TopL > 34 Then TopL = 34 '34 & 35 are last xy pair
   If TopR > 34 Then TopR = 34
   linMeter1L.X2 = CSng(vM1(TopL))
   linMeter1L.Y2 = CSng(vM1(TopL + 1))
   linMeter1R.X2 = CSng(vM1(TopR))
   linMeter1R.Y2 = CSng(vM1(TopR + 1))
  Case Meter2
   TopL = 27 * LLev
   If (TopL And 1) > 0 Then
    TopL = TopL + 1
   End If
   TopR = 27 * RLev
   If (TopR And 1) > 0 Then
    TopR = TopR + 1
   End If
   If TopL > 26 Then TopL = 26
   If TopR > 26 Then TopR = 26
   linMeter2L.X2 = CSng(vM2(TopL))
   linMeter2L.Y2 = CSng(vM2(TopL + 1))
   linMeter2R.X2 = CSng(vM2(TopR))
   linMeter2R.Y2 = CSng(vM2(TopR + 1))
  Case Meter3
   TopL = 33 * LLev
   If (TopL And 1) > 0 Then
    TopL = TopL + 1
   End If
   TopR = 33 * RLev
   If (TopR And 1) > 0 Then
    TopR = TopR + 1
   End If
   If TopL > 32 Then TopL = 32
   If TopR > 32 Then TopR = 32
   linMeter3L.X2 = CSng(vM3(TopL))
   linMeter3L.Y2 = CSng(vM3(TopL + 1))
   linMeter3R.X2 = CSng(vM3(TopR))
   linMeter3R.Y2 = CSng(vM3(TopR + 1))
  Case Meter4
   TopL = 41 * LLev
   If (TopL And 1) > 0 Then
    TopL = TopL + 1
   End If
   TopR = 41 * RLev
   If (TopR And 1) > 0 Then
    TopR = TopR + 1
   End If
   If TopL > 40 Then TopL = 40
   If TopR > 40 Then TopR = 40
   linMeter4L.X2 = CSng(vM4(TopL))
   linMeter4L.Y2 = CSng(vM4(TopL + 1))
   linMeter4R.X2 = CSng(vM4(TopR))
   linMeter4R.Y2 = CSng(vM4(TopR + 1))
 End Select
End Sub
'============UserControl Events===========
Private Sub UserControl_Initialize()
 Set oRec = New WaveInRecorder
 ReDim intSamples(FFT_SAMPLES - 1) As Integer
End Sub
Private Sub UserControl_Terminate()
 Erase intSamples
 oRec.StopRecord
 Set oRec = Nothing
End Sub
Private Sub UserControl_InitProperties()
 mStyle = Meter1
 mMeter3BackColor = UserControl.BackColor
 mOrientation = Horizontal
 mSeparation = 0
End Sub
Private Sub UserControl_ReSize()
 Static Busy As Boolean
 Dim NW As Long, NH As Long
 UserControl.Cls
 UserControl.Picture = LoadPicture
 pMeter1L.Visible = False: pMeter1R.Visible = False
 pMeter2L.Visible = False: pMeter2R.Visible = False
 pMeter3L.Visible = False: pMeter3R.Visible = False
 pMeter3L.Visible = False: pMeter3R.Visible = False
 
 Select Case mStyle
  Case Meter1
   pMeter1L.Visible = True: pMeter1R.Visible = True
   If mOrientation = Horizontal Then
    NW = pMeter1L.Width * 2 + mSeparation
    NH = pMeter1L.Height
    pMeter1L.Move 0, 0
    pMeter1R.Move pMeter1L.Width + mSeparation, 0
   Else
    NW = pMeter1L.Width
    NH = pMeter1L.Height * 2 + mSeparation
    pMeter1L.Move 0, 0
    pMeter1R.Move 0, pMeter1L.Height + mSeparation
   End If
  Case Meter2
   pMeter2L.Visible = True: pMeter2R.Visible = True
   If mOrientation = Horizontal Then
    NW = pMeter2L.Width * 2 + mSeparation
    NH = pMeter2L.Height
    pMeter2L.Move 0, 0
    pMeter2R.Move pMeter2L.Width + mSeparation, 0
   Else
    NW = pMeter2L.Width
    NH = pMeter2L.Height * 2 + mSeparation
    pMeter2L.Move 0, 0
    pMeter2R.Move 0, pMeter2L.Height + mSeparation
   End If
  Case Meter3
   pMeter3L.Visible = True: pMeter3R.Visible = True
   If mOrientation = Horizontal Then
    NW = pMeter3L.Width * 2 + mSeparation
    NH = pMeter3L.Height
    pMeter3L.Move 0, 0
    pMeter3R.Move pMeter3L.Width + mSeparation, 0
   Else
    NW = pMeter3L.Width
    NH = pMeter3L.Height * 2 + mSeparation
    pMeter3L.Move 0, 0
    pMeter3R.Move 0, pMeter3L.Height + mSeparation
   End If
  Case Meter4
   pMeter4L.Visible = True: pMeter4R.Visible = True
   If mOrientation = Horizontal Then
    NW = pMeter4L.Width * 2 + mSeparation
    NH = pMeter4L.Height
    pMeter4L.Move 0, 0
    pMeter4R.Move pMeter4L.Width + mSeparation, 0
   Else
    NW = pMeter4L.Width
    NH = pMeter4L.Height * 2 + mSeparation
    pMeter4L.Move 0, 0
    pMeter4R.Move 0, pMeter4L.Height + mSeparation
   End If
 End Select
 If Not Busy Then
  Busy = True
  UserControl.Width = NW
  UserControl.Height = NH
  Busy = False
 End If
 UserControl.Refresh
 DoEvents
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  mStyle = .ReadProperty("Style", 0)
  mMeter3BackColor = .ReadProperty("Meter3BackColor", vbButtonFace)
  mOrientation = .ReadProperty("Orientation", 0)
  mSeparation = .ReadProperty("Separation", 0)
  UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
  Select Case mStyle
   Case Meter1
    linMeter1L.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter1R.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter1L.BorderWidth = .ReadProperty("NeedleWidth", 1)
    linMeter1R.BorderWidth = .ReadProperty("NeedleWidth", 1)
   Case Meter2
    linMeter2L.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter2R.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter2L.BorderWidth = .ReadProperty("NeedleWidth", 1)
    linMeter2R.BorderWidth = .ReadProperty("NeedleWidth", 1)
   Case Meter3
    linMeter3L.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter3R.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter3L.BorderWidth = .ReadProperty("NeedleWidth", 1)
    linMeter3R.BorderWidth = .ReadProperty("NeedleWidth", 1)
   Case Meter4
    linMeter4L.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter4R.BorderColor = .ReadProperty("NeedleColor", 0)
    linMeter4L.BorderWidth = .ReadProperty("NeedleWidth", 1)
    linMeter4R.BorderWidth = .ReadProperty("NeedleWidth", 1)
  End Select
 End With
 'These x,y pairs define needle (line control) position
 vM1 = Array(19, 29, 22, 25, 25, 23, 30, 20, 35, 18, 38, 16, 44, 13, 49, 13, 53, 13, 58, 13, 67, 14, 72, 15, 75, 17, 78, 19, 82, 22, 86, 25, 89, 27, 90, 28)
 vM2 = Array(20, 33, 25, 30, 30, 28, 36, 25, 45, 25, 54, 23, 58, 23, 62, 22, 69, 22, 76, 23, 83, 26, 88, 28, 93, 30, 98, 32)
 vM3 = Array(12, 41, 16, 39, 21, 35, 26, 32, 30, 30, 34, 29, 39, 27, 41, 27, 46, 26, 50, 27, 55, 29, 60, 30, 64, 30, 69, 32, 74, 35, 75, 37, 79, 41)
 vM4 = Array(19, 55, 22, 51, 26, 48, 30, 44, 36, 41, 41, 38, 46, 38, 52, 35, 57, 33, 63, 34, 67, 34, 73, 34, 79, 34, 83, 34, 90, 35, 98, 39, 101, 39, 103, 41, 107, 43, 109, 46, 112, 50)
 If mStyle = Meter3 Then
  M3BackColor
 End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Style", mStyle, 0
  .WriteProperty "Meter3BackColor", mMeter3BackColor, vbButtonFace
  .WriteProperty "Orientation", mOrientation, 0
  .WriteProperty "Separation", mSeparation, 0
  .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
  Select Case mStyle
   Case Meter1
    .WriteProperty "NeedleColor", linMeter1L.BorderColor, 0
    .WriteProperty "NeedleWidth", linMeter1L.BorderWidth, 1
   Case Meter2
    .WriteProperty "NeedleColor", linMeter2L.BorderColor, 0
    .WriteProperty "NeedleWidth", linMeter2L.BorderWidth, 1
   Case Meter3
    .WriteProperty "NeedleColor", linMeter3L.BorderColor, 0
    .WriteProperty "NeedleWidth", linMeter3L.BorderWidth, 1
   Case Meter4
    .WriteProperty "NeedleColor", linMeter4L.BorderColor, 0
    .WriteProperty "NeedleWidth", linMeter4L.BorderWidth, 1
  End Select

 End With
End Sub


