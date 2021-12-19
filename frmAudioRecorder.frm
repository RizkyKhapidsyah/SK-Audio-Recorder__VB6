VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Recorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AudioRecorder"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmAudioRecorder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Height          =   495
      Left            =   5970
      TabIndex        =   10
      ToolTipText     =   "Change rate, stereo/mono, 8/16 bits and program an automatic recording"
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "You can choose a beginning for playing the recording"
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   500
      SmallChange     =   100
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web"
      Height          =   495
      Left            =   4995
      TabIndex        =   7
      ToolTipText     =   "Visit the home page of me!! (Maybe a new version is available...)"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "To start a new recording and adjusting all settings"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4020
      TabIndex        =   3
      ToolTipText     =   "Save the recording as as WAV file"
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   5760
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   " "
      Orientation     =   2
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3045
      TabIndex        =   2
      ToolTipText     =   "Play the recording"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2070
      TabIndex        =   1
      ToolTipText     =   "Stop recording or playing"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   495
      Left            =   1095
      TabIndex        =   0
      ToolTipText     =   "Start recording immediate"
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Starting position for play (in milliseconds)"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4815
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   5160
      Top             =   2400
   End
   Begin VB.Frame Frame4 
      Caption         =   "Statistics"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
      Begin VB.Label StatLbl 
         BackColor       =   &H00000000&
         Caption         =   " "
         ForeColor       =   &H0000FF00&
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Information about the recording"
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Recorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright: E. de Vries
'e-mail: eeltje@geocities.com
'This code can be used as freeware

Const AppName = "Recorder"

Private Sub cmdSave_Click()
    Dim sName As String
    
    If WavMidi = "" Then
        sName = "Record_from_" & CStr(WaRecoStTi) & "_to_" & CStr(WavRecStopT)
        sName = Replace(sName, ":", "-")
        sName = Replace(sName, " ", "_")
        sName = Replace(sName, "/", "-")
    Else
        sName = WavMidi
        sName = Replace(sName, "MID", "wav")
    End If
  
    ComDlg.FileName = sName
    ComDlg.CancelError = True
    On Error GoTo ErrHandler1
    ComDlg.Filter = "WAV file (*.wav*)|*.wav"
    ComDlg.Flags = &H2 Or &H400
    ComDlg.ShowSave
    sName = ComDlg.FileName
    
    WaveSaveAs (sName)
    Exit Sub
ErrHandler1:
End Sub

Private Sub cmdRecord_Click()
    Dim settings As String
    Dim Alignment As Integer
      
    Alignment = Channels * Res / 8
    
    settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Res) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)
    WaveReset
    WaveSet
    WaveRecord
    WaRecoStTi = Now
    cmdStop.Enabled = True   'Enable the STOP BUTTON
    cmdPlay.Enabled = False  'Disable the "PLAY" button
    cmdSave.Enabled = False  'Disable the "SAVE AS" button
    cmdRecord.Enabled = False 'Disable the "RECORD" button
End Sub

Private Sub cmdSettings_Click()
Dim strWhat As String
    ' show the user entry form modally
    strWhat = MsgBox("If you continue your data will be lost!", vbOKCancel)
    If strWhat = vbCancel Then
        Exit Sub
    End If
    Slider1.Max = 10
    Slider1.Value = 0
    Slider1.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("Recorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("Recorder", "StartUp", "Channels", "1"))
    Res = CInt(GetSetting("Recorder", "StartUp", "Res", "16"))
    WavFN = GetSetting("Recorder", "StartUp", "WavFN", "C:\Radio.wav")
    WavAutoSave = GetSetting("Recorder", "StartUp", "WavAutoSave", "True")

    WaRecIm = True
    WavRecR = False
    WavRecord = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    frmSettings.optRecordImmediate.Value = True
    frmSettings.Show vbModal
End Sub

Private Sub cmdStop_Click()
    WaveStop
    cmdSave.Enabled = True  'Enable the "SAVE AS" button
    cmdPlay.Enabled = True  'Enable the "PLAY" button
    cmdStop.Enabled = False 'Disable the "STOP" button
    If WavePosition = 0 Then
        Slider1.Max = 10
    Else
        If WaRecIm And (Not WavePlaying) Then Slider1.Max = WavePosition
        If (Not WaRecIm) And WavRecord Then Slider1.Max = WavePosition
    End If
    If WavRecord Then WavRecR = True
    WavRecStopT = Now
    WavRecord = False
    WavePlaying = False
    frmSettings.optRecordProgrammed.Value = False
    frmSettings.optRecordImmediate.Value = True
    frmSettings.lblTimes.Visible = False
End Sub

Private Sub cmdPlay_Click()
    WavePlayFrom (Slider1.Value)
    WavePlaying = True
    cmdStop.Enabled = True
    cmdPlay.Enabled = False
End Sub


Private Sub cmdWeb_Click()
  Dim ret&
  ret& = ShellExecute(Me.hwnd, "Open", "http://home.wxs.nl/~eeltjevr/", "", App.Path, 1)
End Sub




Private Sub cmdReset_Click()
    Slider1.Max = 10
    Slider1.Value = 0
    Slider1.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("Recorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("Recorder", "StartUp", "Channels", "1"))
    Res = CInt(GetSetting("Recorder", "StartUp", "Res", "16"))
    WavFN = GetSetting("Recorder", "StartUp", "WavFN", "C:\Radio.wav")
    WavAutoSave = GetSetting("Recorder", "StartUp", "WavAutoSave", "True")

    WaRecIm = True
    WavRecR = False
    WavRecord = False
    WavePlaying = False
    WavMidi = ""
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    If WavRenaNeces Then
        Name WavShFN As WavLongFN
        WavRenaNeces = False
        WavShFN = ""
    End If
End Sub

Private Sub Form_Load()
    WaveReset
    
    Rate = CLng(GetSetting("Recorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("Recorder", "StartUp", "Channels", "1"))
    Res = CInt(GetSetting("Recorder", "StartUp", "Res", "16"))
    WavFN = GetSetting("Recorder", "StartUp", "WavFN", "C:\Radio.wav")
    WavAutoSave = GetSetting("Recorder", "StartUp", "WavAutoSave", "True")

    WaRecIm = True
    WavRecR = False
    WavRecord = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    WaRecoStTi = Now + TimeSerial(0, 15, 0)
    WavRecStopT = WaRecoStTi + TimeSerial(0, 15, 0)
    WavMidi = ""
    WavRenaNeces = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WaveClose
    Call SaveSetting("Recorder", "StartUp", "Rate", CStr(Rate))
    Call SaveSetting("Recorder", "StartUp", "Channels", CStr(Channels))
    Call SaveSetting("Recorder", "StartUp", "Res", CStr(Res))
    Call SaveSetting("Recorder", "StartUp", "WavFN", WavFN)
    Call SaveSetting("Recorder", "StartUp", "WavAutoSave", CStr(WavAutoSave))
    If WavRenaNeces Then
        Name WavShFN As WavLongFN
        WavRenaNeces = False
        WavShFN = ""
    End If
    End
End Sub



Private Sub Timer2_Timer()
    Dim RecordingTimes As String
    Dim msg As String
    
    RecordingTimes = "Start time:  " & WaRecoStTi & vbCrLf _
                    & "Stop time:  " & WavRecStopT
    
    WaveStatistics
    If Not WaRecIm Then
        WavSticMsg = WavSticMsg & "Programmed recording"
        If WavAutoSave Then
            WavSticMsg = WavSticMsg & " (automatic save)"
        Else
            WavSticMsg = WavSticMsg & " (manual save)"
        End If
        WavSticMsg = WavSticMsg & vbCrLf & vbCrLf & RecordingTimes
    End If
    StatLbl.Caption = WavSticMsg
    
    WaveStatus
    If WavStatMsg <> Recorder.Caption Then Recorder.Caption = WavStatMsg
    If InStr(Recorder.Caption, "stopped") > 0 Then
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
    End If
    
    If RecordingTimes <> frmSettings.lblTimes.Caption Then frmSettings.lblTimes.Caption = RecordingTimes
    
    If (Now > WaRecoStTi) _
            And (Not WavRecR) _
            And (Not WaRecIm) _
            And (Not WavRecord) Then
        WaveReset
        WaveSet
        WaveRecord
        WavRecord = True
        cmdStop.Enabled = True   'Enable the STOP BUTTON
        cmdPlay.Enabled = False  'Disable the "PLAY" button
        cmdSave.Enabled = False  'Disable the "SAVE AS" button
        cmdRecord.Enabled = False 'Disable the "RECORD" button
    End If
    
    If (Now > WavRecStopT) And (Not WavRecR) And (Not WaRecIm) Then
        WaveStop
        cmdSave.Enabled = True 'Enable the "SAVE AS" button
        cmdPlay.Enabled = True 'Enable the "PLAY" button
        cmdStop.Enabled = False 'Disable the "STOP" button
        If WavePosition > 0 Then
            Slider1.Max = WavePosition
        Else
            Slider1.Max = 10
        End If
        WavRecord = False
        WavRecR = True
        If WavAutoSave Then
            WavFN = "Radio_from_" & CStr(WaRecoStTi) & "_to_" & CStr(WavRecStopT)
            WavFN = Replace(WavFN, ":", ".")
            WavFN = Replace(WavFN, " ", "_")
            WavFN = WavFN & ".wav"
            WaveSaveAs (WavFN)
            msg = "Recording has been saved" & vbCrLf
            msg = msg & "Filename: " & WavFN
            MsgBox (msg)
        Else
            msg = "Recording is ready" & vbCrLf
            msg = msg & "Don't forget to save recording..."
            MsgBox (msg)
        End If
        frmSettings.optRecordProgrammed.Value = False
        frmSettings.optRecordImmediate.Value = True
    End If

End Sub
