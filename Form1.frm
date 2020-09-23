VERSION 5.00
Begin VB.Form FrmVisualize 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   5970
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   11700
   ControlBox      =   0   'False
   DrawWidth       =   3
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Line s2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   1680
      X2              =   3240
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line s 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7800
      X2              =   9360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1st, RUN WINAMP. Play a song. Keep it running and then run the Magic Lasers!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   3615
   End
End
Attribute VB_Name = "FrmVisualize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Coded by Kailash Nadh, http://bnsoft.net , kailash@bnsoft.net

' Dancing Musical lasers which dance with a song!

' Actually, the code to get the frequency from
' the sound card is not mine! I just made the laser part!

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim Barwidth As Integer, a As Integer
Dim Gap As Integer
Dim Skip As Integer
Dim NoOfLines
Dim Falls As Integer
Dim ColorS(2) As Integer
Dim Pointx As Integer
Dim PointFalls As Integer

Dim d As Integer
Dim w(600) As Integer
Dim p(600) As Integer
Dim Px As Integer

Dim start As Long
Dim ender As Long
Dim xSD As Long

Dim X As Integer
Dim Y As Boolean
Option Explicit
Dim linked As Boolean
Dim hmixer As Long
Dim inputVolCtrl As MIXERCONTROL
Dim outputVolCtrl As MIXERCONTROL
Dim rc As Long
Dim OK As Boolean
Dim mxcd As MIXERCONTROLDETAILS
Dim vol As MIXERCONTROLDETAILS_SIGNED
Dim volume As Long
Dim volHmem As Long
Private VU As VULights
Private FreqNum As Frequen
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim picX As Integer
Dim picY As Integer
Dim newX As Integer
Dim newY As Integer
Dim oldX As Integer
Dim Oldy As Integer
Dim Started As Boolean
Dim sFile As String
Dim sOutput As String
Dim TopS As Long
Dim SetF As Integer

Dim picX2, picY2, oldX2, Oldy2, newX2, newY2

Private Sub Form_Click()
Unload Me
End
End Sub

Private Sub form_load()
On Error Resume Next
Dim ForCol As OLE_COLOR, BackCol As OLE_COLOR
' Open Mixer
  rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   Else
   End If
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   ' Set Maximum Volume
   TopS = outputVolCtrl.lMaximum
   
Me.Show
Me.SetFocus
Do
DoEvents
Dim ValU As Single
' get volume
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    
    ' convert volume into perc
   ValU = (volume / TopS)
    ValU = ValU * 100
    ValU = Int(ValU)
    Laser ValU ' Call circle function
    


    If GetKeyState(vbKeyEscape) < -5 Then End

    
Loop

End Sub

'###########################################
'###########################################
'###########################################
'###########################################

Function Laser(Percent As Single)

'Percent is the sound's percent. You can use
'it to make your own effects

If Percent <= 0 Then
'If percent is less than 0, slow down the lasers
Me.Cls
s2.Y1 = Percent * 20
s.Y2 = Percent * 20
Me.s.BorderColor = RGB(Rnd * 25, Rnd * 25, Rnd * 2500)
Me.s2.BorderColor = RGB(Rnd * 25, Rnd * 25, Rnd * 2500)
Me.Circle (s2.X1, s2.Y1), 50, vbBlue
Me.Circle (s.X2, s.Y2), 50, vbBlue
Exit Function
End If

'Do normal dancing
s2.Y1 = Percent * 100
s.Y2 = Percent * 100
Me.Cls
Me.Circle (s2.X1, s2.Y1), 50, vbBlue
Me.Circle (s.X2, s.Y2), 50, vbBlue

Me.s.BorderColor = RGB(Rnd * 25, Rnd * 25, Rnd * 2500)
Me.s2.BorderColor = RGB(Rnd * 25, Rnd * 25, Rnd * 2500)

'If the frequency is high, make it Rock!
If Percent > 70 Then
Me.Circle (5000, 5000), Percent * 10, vbRed
End If

'Extreme!!! Make the circles appear
If Percent = 100 Then
Me.Circle (5000, 5000), Percent * 4, vbBlue
End If
'###########################################

End Function

