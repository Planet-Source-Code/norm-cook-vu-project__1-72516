VERSION 5.00
Begin VB.UserControl BarsH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   ToolboxBitmap   =   "BarsH.ctx":0000
   Begin VB.PictureBox pBarsL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6900
      Left            =   4320
      Picture         =   "BarsH.ctx":0312
      ScaleHeight     =   460
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.PictureBox pBarsM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   1560
      Picture         =   "BarsH.ctx":76CD4
      ScaleHeight     =   230
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox pBarsS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   120
      Picture         =   "BarsH.ctx":94776
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "BarsH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eBHSize
 SmallBH
 MediumBH
 LargeBH
End Enum
Public Enum eBHChannel
 LeftChanBH
 RightChanBH
End Enum
Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private mSize As eBHSize
Private mChan As eBHChannel
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
Public Sub Preview() 'just threw this in for the demo
 Graphics 0.5
End Sub
'The class's only event
Private Sub oRec_GotData(intBuffer() As Integer, lngLen As Long)
 Dim lngMaxL As Long, lngMaxR As Long
 intSamples = intBuffer
 'left is the even numbers, right is odd
 lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
 lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
 If mChan = LeftChanBH Then
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

Public Property Get Channel() As eBHChannel
 Channel = mChan
End Property
Public Property Let Channel(ByVal NewChan As eBHChannel)
 mChan = NewChan
End Property
Public Property Get VUSize() As eBHSize
 VUSize = mSize
End Property
Public Property Let VUSize(ByVal NewSiz As eBHSize)
 mSize = NewSiz
 UserControl_ReSize
 DoColors
End Property
Private Sub Graphics(ByVal Lev As Single)
 Select Case mSize
  Case SmallBH
   BitBlt hdc, 0, 0, 89, 6, _
     pBarsS.hdc, 0, 5 * CLng(22 * Lev), vbSrcCopy
  Case MediumBH
   BitBlt hdc, 0, 0, 176, 12, _
     pBarsM.hdc, 0, 10 * CLng(22 * Lev), vbSrcCopy
  Case LargeBH
   BitBlt hdc, 0, 0, 352, 24, _
     pBarsL.hdc, 0, 20 * CLng(22 * Lev), vbSrcCopy
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
  Case SmallBH
   NW = 1335: NH = 90
  Case MediumBH
   NW = 2640: NH = 150
  Case LargeBH
   NW = 5280: NH = 300
 End Select
 If Not Busy Then
  Busy = True 'prevent recursive resizing
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
Private Sub UserControl_Terminate()
 oRec.StopRecord
 Set oRec = Nothing
End Sub
Private Sub UserControl_InitProperties()
 mSize = SmallBH
 mChan = LeftChanBH
 UserControl.ForeColor = vbGreen
 UserControl.BackColor = vbBlack
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
'Completely redraws the bar pics
' with the desired back/forecolor
Private Sub DrawHS() 'small
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsS.Line (0, 0)-(89, 116), BackColor, BF
 XCnt = 1
 For y = 6 To 111 Step 5
  x = 1
  For i = 1 To XCnt
   pBarsS.Line (x, y)-(x + 2, y + 3), ForeColor, BF
   x = x + 4
  Next
  XCnt = XCnt + 1
 Next
End Sub

Private Sub DrawHM() 'medium
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsM.Line (0, 0)-(176, 230), BackColor, BF
 XCnt = 1
 For y = 11 To 221 Step 10
  x = 1
  For i = 1 To XCnt
   pBarsM.Line (x, y)-(x + 5, y + 7), ForeColor, BF
   x = x + 8
  Next
  XCnt = XCnt + 1
 Next
End Sub
Private Sub DrawHL() 'large
 Dim x As Long, y As Long, i As Long, XCnt As Long
 pBarsL.Line (0, 0)-(352, 460), BackColor, BF
 XCnt = 1
 For y = 22 To 458 Step 20
  x = 2
  For i = 1 To XCnt
   pBarsL.Line (x, y)-(x + 11, y + 15), ForeColor, BF
   x = x + 16
  Next
  XCnt = XCnt + 1
 Next
End Sub
Private Sub DoColors()
 Select Case mSize
  Case SmallBH
   DrawHS
  Case MediumBH
   DrawHM
  Case LargeBH
   DrawHL
 End Select
 If Not UserControl.Ambient Then
  Graphics 0.5
 End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
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
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
 ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
 UserControl.ForeColor() = New_ForeColor
 PropertyChanged "ForeColor"
 DoColors
End Property

