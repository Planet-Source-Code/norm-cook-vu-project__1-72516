VERSION 5.00
Begin VB.UserControl GradH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ToolboxBitmap   =   "GradH.ctx":0000
End
Attribute VB_Name = "GradH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eGHSize
 SmallGH
 MediumGH
 LargeGH
End Enum
Public Enum eGHChannel
 LeftChanGH
 RightChanGH
End Enum
Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private mSize As eGHSize
Private mChan As eGHChannel
Private mGradientColor1 As Long
Private mGradientColor2 As Long
Private intSamples() As Integer
Public Property Get Channel() As eGHChannel
 Channel = mChan
End Property
Public Property Let Channel(ByVal NewChan As eGHChannel)
 mChan = NewChan
End Property
Public Property Get VUSize() As eGHSize
 VUSize = mSize
End Property
Public Property Let VUSize(ByVal NewSiz As eGHSize)
 mSize = NewSiz
 UserControl_ReSize
End Property
Public Property Get GradientColor1() As OLE_COLOR
 GradientColor1 = mGradientColor1
End Property
Public Property Let GradientColor1(ByVal NewCol As OLE_COLOR)
 mGradientColor1 = NewCol
 DoColors
End Property
Public Property Get GradientColor2() As OLE_COLOR
 GradientColor2 = mGradientColor2
End Property
Public Property Let GradientColor2(ByVal NewCol As OLE_COLOR)
 mGradientColor2 = NewCol
 DoColors
End Property

Private Sub oRec_GotData(intBuffer() As Integer, lngLen As Long)
 Dim lngMaxL As Long, lngMaxR As Long
 intSamples = intBuffer
 lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
 lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
 If mChan = LeftChanGH Then
  Graphics lngMaxL / 32768#
 Else
  Graphics lngMaxR / 36738#
 End If
End Sub
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
 mGradientColor1 = vbGreen
 mGradientColor2 = vbRed
 UserControl.BackColor = vbBlack
End Sub
Private Sub Graphics(ByVal Lev As Single)
 Cls
 Select Case mSize
  Case SmallGH
   GradientFillRectDC hdc, 0, 0, Lev * ScaleWidth, ScaleHeight, mGradientColor1, mGradientColor2, GF_RECTHORIZ
  Case MediumGH
   GradientFillRectDC hdc, 0, 0, Lev * ScaleWidth, ScaleHeight, mGradientColor1, mGradientColor2, GF_RECTHORIZ
  Case LargeGH
   GradientFillRectDC hdc, 0, 0, Lev * ScaleWidth, ScaleHeight, mGradientColor1, mGradientColor2, GF_RECTHORIZ
 End Select
 Refresh
End Sub

Private Sub UserControl_ReSize()
 Static Busy As Boolean
 Dim NW As Long, NH As Long
 Select Case mSize
  Case SmallGH
   NW = 1335: NH = 90
  Case MediumGH
   NW = 2640: NH = 150
  Case LargeGH
   NW = 5280: NH = 300
 End Select
 If Not Busy Then
  Busy = True
  UserControl.Width = NW
  UserControl.Height = NH
  Busy = False
 End If
 DoColors
End Sub
Private Sub DoColors()
 If Ambient.UserMode Then
  Graphics 0
 Else
  Graphics 0.5
 End If
End Sub
Public Property Get BackColor() As OLE_COLOR
 BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 UserControl.BackColor() = New_BackColor
 PropertyChanged "BackColor"
 DoColors
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  mChan = .ReadProperty("Channel", 0)
  mSize = .ReadProperty("VUSize", 0)
  UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
  mGradientColor1 = .ReadProperty("GradientColor1", vbGreen)
  mGradientColor2 = .ReadProperty("GradientColor2", vbRed)
 End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Channel", mChan, 0
  .WriteProperty "VUSize", mSize, 0
  .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
  .WriteProperty "GradientColor1", mGradientColor1, vbGreen
  .WriteProperty "GradientColor2", mGradientColor2, vbRed
 End With
End Sub

