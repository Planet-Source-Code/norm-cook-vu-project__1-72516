VERSION 5.00
Begin VB.UserControl LightsH 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ToolboxBitmap   =   "LightsH.ctx":0000
   Begin VB.PictureBox pLitesM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   360
      Picture         =   "LightsH.ctx":0312
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox pLitesL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   2760
      Picture         =   "LightsH.ctx":19854
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.PictureBox pLitesS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   840
      Picture         =   "LightsH.ctx":7EC96
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "LightsH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eLHSize
 SmallLH
 MediumLH
 LargeLH
End Enum
Public Enum eLHChannel
 LeftChanLH
 RightChanLH
End Enum
Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private mSize As eLHSize
Private mChan As eLHChannel
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
Private Sub oRec_GotData(intBuffer() As Integer, lngLen As Long)
 Dim lngMaxL As Long, lngMaxR As Long
 intSamples = intBuffer
 lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
 lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
 If mChan = LeftChanLH Then
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

Public Property Get Channel() As eLHChannel
 Channel = mChan
End Property
Public Property Let Channel(ByVal NewChan As eLHChannel)
 mChan = NewChan
End Property
Public Property Get VUSize() As eLHSize
 VUSize = mSize
End Property
Public Property Let VUSize(ByVal NewSiz As eLHSize)
 mSize = NewSiz
 UserControl_ReSize
End Property
Private Sub Graphics(ByVal Lev As Single)
 Select Case mSize
  Case SmallLH
   BitBlt hdc, 0, 0, 90, 6, _
     pLitesS.hdc, 0, 6 * CLng(16 * Lev), vbSrcCopy
  Case MediumLH
   BitBlt hdc, 0, 0, 180, 12, _
     pLitesM.hdc, 0, 12 * CLng(16 * Lev), vbSrcCopy
  Case LargeLH
   BitBlt hdc, 0, 0, 360, 24, _
     pLitesL.hdc, 0, 24 * CLng(16 * Lev), vbSrcCopy
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
  Case SmallLH
   NW = 1350: NH = 90
  Case MediumLH
   NW = 2700: NH = 180
  Case LargeLH
   NW = 5400: NH = 360
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
 mSize = SmallLH
 mChan = LeftChanLH
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  mChan = .ReadProperty("Channel", 0)
  mSize = .ReadProperty("VUSize", 0)
 End With
End Sub

Private Sub UserControl_Terminate()
 oRec.StopRecord
 Set oRec = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Channel", mChan, 0
  .WriteProperty "VUSize", mSize, 0
 End With
End Sub
