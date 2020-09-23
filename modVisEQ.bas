Attribute VB_Name = "modSound"
Option Explicit
Public hmixer As Long
Public inputVolCtrl As MIXERCONTROL
Public outputVolCtrl As MIXERCONTROL
Public rc As Long
Public OK As Boolean
Public mxcd As MIXERCONTROLDETAILS
Public vol As MIXERCONTROLDETAILS_SIGNED
Public volume As Long
Public volHmem As Long

Public VolValue As Double

Public Enum Frequen
    Freq60Hz = 60
    Freq170Hz = 170
    Freq310Hz = 310
    Freq600Hz = 600
    Freq3kHz = 3
    Freq6kHz = 6
    Freq12kHz = 12
    Freq14kHz = 14
    Freq31Hz = 32.3
    Freq62Hz = 62.5
    Freq125Hz = 125
    Freq250Hz = 250
    Freq500Hz = 500
    Freq1kHz = 1
    Freq2kHz = 2
    Freq4kHz = 4
    Freq8kHz = 8
    Freq16kHz = 16.1
End Enum

Public Type VULights
    VUOn As Boolean
    InOutLev As Double
    VolLev As Double
    VULev As Double
    VolUnit As Variant
    VUArray As Long
    Freq(0 To 9) As Double
    FreqNum As Integer
    FreqVal As Double
End Type

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm" () As Long

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const Flags& = SND_ASYNC Or SND_NODEFAULT

Public Const CALLBACK_FUNCTION = &H30000
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1
Public Const GMEM_FIXED = &H0

Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
                           

   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type

Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type

Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long

Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long

Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long
Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32

Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)

Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
            (ByVal hmxobj As Long, _
            pmxcd As MIXERCONTROLDETAILS, _
            ByVal fdwDetails As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" _
                  (ByVal uMxId As Long, _
                  ByVal pmxcaps As MIXERCAPS, _
                  ByVal cbmxcaps As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, _
                                             pumxID As Long, _
                                             ByVal fdwId As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" _
                  (ByVal hmxobj As Long, _
                  pmxlc As MIXERLINECONTROLS, _
                  ByVal fdwControls As Long) As Long

Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" _
                     (ByVal hmxobj As Long, _
                     pmxl As MIXERLINE, _
                     ByVal fdwInfo As Long) As Long

Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long

Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, _
                                                ByVal uMsg As Long, _
                                                ByVal dwParam1 As Long, _
                                                ByVal dwParam2 As Long) As Long

Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, _
                                             ByVal uMxId As Long, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" _
         (ByVal hmxobj As Long, _
         pmxcd As MIXERCONTROLDETAILS, _
         ByVal fdwDetails As Long) As Long

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                             ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Type MIXERCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
   fdwSupport As Long
   cDestinations As Long
End Type

Type MIXERCONTROL
   cbStruct As Long
   dwControlID As Long
   dwControlType As Long
   fdwControl As Long
   cMultipleItems As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   lMinimum As Long
   lMaximum As Long
   Reserved(10) As Long
End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long
   dwControlID As Long
   cChannels As Long
   item As Long
   cbDetails As Long
   paDetails As Long
End Type

Type MIXERCONTROLDETAILS_SIGNED
   lValue As Long
End Type

Type MIXERLINE
   cbStruct As Long
   dwDestination As Long
   dwSource As Long
   dwLineID As Long
   fdwLine As Long
   dwUser As Long
   dwComponentType As Long
   cChannels As Long
                           
   cConnections As Long
                           
   cControls As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
                                                  
                                                   
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long
   dwLineID As Long
   dwControl As Long
   cControls As Long
   cbmxctrl As Long
   pamxctrl As Long
End Type

Public i As Integer
Public j As Integer
Public msg As String * 200
Public hWaveIn As Long
Public format1 As WAVEFORMAT

Public Const NUM_BUFFERS = 2
Public Const BUFFER_SIZE = 8192
Public Const DEVICEID = 0
Public hmem(NUM_BUFFERS) As Long
Public inHdr(NUM_BUFFERS) As WAVEHDR

Public fRecording As Boolean

Function GetControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
      mxlc.cbStruct = Len(mxlc)
      mxlc.dwLineID = mxl.dwLineID
      mxlc.dwControl = ctrlType
      mxlc.cControls = 1
      mxlc.cbmxctrl = Len(mxc)
      mxlc.pamxctrl = 9
      
      hmem = GlobalAlloc(GMEM_FIXED, Len(mxc))
      mxlc.pamxctrl = GlobalLock(hmem)
      mxc.cbStruct = Len(mxc)
      
      rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
            
      If (MMSYSERR_NOERROR = rc) Then
         GetControl = True
         
         CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
      Else
         GetControl = False
      End If
      GlobalFree (hmem)
      Exit Function
   End If
   
   GetControl = False
End Function

Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format1.wFormatTag = 1
    format1.nChannels = 1
    format1.wBitsPerSample = 8
    format1.nSamplesPerSec = 8000
    format1.nBlockAlign = format1.nChannels * format1.wBitsPerSample / 8
    format1.nAvgBytesPerSec = format1.nSamplesPerSec * format1.nBlockAlign
    format1.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format1, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

Sub StopInput()

    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub
