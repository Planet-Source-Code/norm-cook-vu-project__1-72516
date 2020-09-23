VERSION 5.00
Begin VB.UserControl LightsV 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ToolboxBitmap   =   "LightsV.ctx":0000
   Begin VB.PictureBox pLitesL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   3720
      Picture         =   "LightsV.ctx":0312
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.PictureBox pLitesM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   600
      Picture         =   "LightsV.ctx":65754
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox pLitesS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   360
      Picture         =   "LightsV.ctx":7EC96
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "LightsV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eLVSize
 SmallLV
 MediumLV
 LargeLV
End Enum
Public Enum eLVChannel
 LeftChanLV
 RightChanLV
End Enum
Private WithEvents oRec As WaveInRecorder
Attribute oRec.VB_VarHelpID = -1
Private mSize As eLVSize
Private mChan As eLVChannel
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
 If mChan = LeftChanLV Then
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

Public Property Get Channel() As eLVChannel
 Channel = mChan
End Property
Public Property Let Channel(ByVal NewChan As eLVChannel)
 mChan = NewChan
End Property
Public Property Get VUSize() As eLVSize
 VUSize = mSize
End Property
Public Property Let VUSize(ByVal NewSiz As eLVSize)
 mSize = NewSiz
 UserControl_ReSize
End Property
Private Sub Graphics(ByVal Lev As Single)
 Select Case mSize
  Case SmallLV
   BitBlt hdc, 0, 0, 6, 90, _
     pLitesS.hdc, 6 * CLng(16 * Lev), 0, vbSrcCopy
  Case MediumLV
   BitBlt hdc, 0, 0, 12, 180, _
     pLitesM.hdc, 12 * CLng(16 * Lev), 0, vbSrcCopy
  Case LargeLV
   BitBlt hdc, 0, 0, 24, 360, _
     pLitesL.hdc, 24 * CLng(16 * Lev), 0, vbSrcCopy
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
  Case SmallLV
   NH = 1350: NW = 90
  Case MediumLV
   NH = 2700: NW = 180
  Case LargeLV
   NH = 5400: NW = 360
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
 mSize = SmallLV
 mChan = LeftChanLV
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


