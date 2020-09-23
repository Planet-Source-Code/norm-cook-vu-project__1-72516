VERSION 5.00
Begin VB.UserControl BarsV 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ToolboxBitmap   =   "BarsV.ctx":0000
   Begin VB.PictureBox pBarsS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      Picture         =   "BarsV.ctx":0312
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.PictureBox pBarsM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   120
      Picture         =   "BarsV.ctx":7C50
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox pBarsL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   120
      Picture         =   "BarsV.ctx":25852
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   460
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   6900
   End
End
Attribute VB_Name = "BarsV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eBVSize
 SmallBV
 MediumBV
 LargeBV
End Enum
Public Enum eBVChannel
 LeftChan
 RightChan
End Enum
Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private mSize As eBVSize
Private mChan As eBVChannel
Private intSamples() As Integer
Public Sub StartVU()
 If Not oRec.IsRecording Then
  oRec.StartRecord 44100, 2
 End If
End Sub
Public Sub StopVU()
 Graphics 0
 oRec.StopRecord
End Sub
Public Sub Preview()
 Graphics 0.5
End Sub
Private Sub oRec_GotData(intBuffer() As Integer, lngLen As Long)
 Dim lngMaxL As Long, lngMaxR As Long
 intSamples = intBuffer
 lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
 lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
 If mChan = LeftChan Then
  Graphics lngMaxL / 32768#
 Else
  Graphics lngMaxR / 36738#
 End If
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

Public Property Get Channel() As eBVChannel
 Channel = mChan
End Property
Public Property Let Channel(ByVal NewChan As eBVChannel)
 mChan = NewChan
End Property
Public Property Get VUSize() As eBVSize
 VUSize = mSize
End Property
Public Property Let VUSize(ByVal NewSiz As eBVSize)
 mSize = NewSiz
 UserControl_ReSize
 DoColors
End Property
Private Sub Graphics(ByVal Lev As Single)
 Select Case mSize
  Case SmallBV
   BitBlt hdc, 0, 0, 6, 89, _
     pBarsS.hdc, 5 * CLng(22 * Lev), 0, vbSrcCopy
  Case MediumBV
   BitBlt hdc, 0, 0, 12, 176, _
     pBarsM.hdc, 10 * CLng(22 * Lev), 0, vbSrcCopy
  Case LargeBV
   BitBlt hdc, 0, 0, 24, 352, _
     pBarsL.hdc, 20 * CLng(22 * Lev), 0, vbSrcCopy
 End Select
 Refresh
End Sub

Private Sub UserControl_Initialize()
 Set oRec = New WaveInRecorder
 ReDim intSamples(FFT_SAMPLES - 1) As Integer
End Sub

Private Sub UserControl_ReSize()
 Static Busy As Boolean
 Dim NW As Long, NH As Long
 Select Case mSize
  Case SmallBV
   NH = 1335: NW = 90
  Case MediumBV
   NH = 2640: NW = 150
  Case LargeBV
   NH = 5280: NW = 300
 End Select
 If Not Busy Then
  Busy = True
  UserControl.Width = NW
  UserControl.Height = NH
  Busy = False
 End If
 If Ambient.UserMode Then
  Graphics 0
 Else
  Graphics 0.5
 End If
End Sub
Private Sub UserControl_InitProperties()
 mSize = SmallBV
 mChan = LeftChan
 UserControl.ForeColor = vbGreen
 UserControl.BackColor = vbBlack
End Sub
Private Sub UserControl_Terminate()
 oRec.StopRecord
 Set oRec = Nothing
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  mChan = .ReadProperty("Channel", 0)
  mSize = .ReadProperty("VUSize", 0)
  UserControl.BackColor = .ReadProperty("BackColor", vbBlack)
  UserControl.ForeColor = .ReadProperty("ForeColor", vbGreen)
 End With
 DoColors
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Channel", mChan, 0
  .WriteProperty "VUSize", mSize, 0
  .WriteProperty "BackColor", UserControl.BackColor, vbBlack
  .WriteProperty "ForeColor", UserControl.ForeColor, vbGreen
 End With
End Sub
Private Sub DoColors()
 Select Case mSize
  Case SmallBV
   BarVS
  Case MediumBV
   BarVM
  Case LargeBV
   BarVL
 End Select
 If Not UserControl.Ambient Then
  Graphics 0.5
 End If
End Sub
'Complete redraw
Private Sub BarVS()
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsS.Line (0, 0)-(116, 89), BackColor, BF
 XCnt = 1
 For y = 1 To 85 Step 4
  x = 111
  For i = 1 To XCnt
   pBarsS.Line (x, y)-(x + 3, y + 2), ForeColor, BF
   x = x - 5
  Next
  XCnt = XCnt + 1
 Next
End Sub
Private Sub BarVM()
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsM.Line (0, 0)-(230, 176), BackColor, BF
 XCnt = 1
 For y = 1 To 169 Step 8
  x = 221
  For i = 1 To XCnt
   pBarsM.Line (x, y)-(x + 7, y + 5), ForeColor, BF
   x = x - 10
  Next
  XCnt = XCnt + 1
 Next
End Sub
Private Sub BarVL()
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsL.Line (0, 0)-(460, 352), BackColor, BF
 XCnt = 1
 For y = 2 To 342 Step 16
  x = 442
  For i = 1 To XCnt
   pBarsL.Line (x, y)-(x + 14, y + 10), ForeColor, BF
   x = x - 20
  Next
  XCnt = XCnt + 1
 Next
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
 BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 UserControl.BackColor() = New_BackColor
 PropertyChanged "BackColor"
 DoColors
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
 ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
 UserControl.ForeColor() = New_ForeColor
 PropertyChanged "ForeColor"
 DoColors
End Property


