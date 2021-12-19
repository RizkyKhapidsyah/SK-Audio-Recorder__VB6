Attribute VB_Name = "modWave"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Rate As Long
    
Public Channels As Integer

Public Res As Integer

Public WavStatMsg As String * 255

Public WavSticMsg As String

Public WaRecIm As Boolean

Public WaRecoStTi As Date

Public WavRecStopT As Date

Public WavRecR As Boolean

Public WavRecord As Boolean

Public WavePlaying As Boolean

Public WavAutoSave As Boolean

Public WavFN As String

Public WavMidi As String

Public WavLongFN As String
Public WavShFN As String
Public WavRenaNeces As Boolean

'These were the public variables
'===============================================================================
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrrtning As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
   
Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
     
Private Declare Function FindFirstFile& Lib "kernel32" _
       Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
       As WIN32_FIND_DATA)

Private Declare Function FindClose Lib "kernel32" _
       (ByVal hFindFile As Long) As Long
       
Private Const MAX_PATH = 260

Private Type FILETIME ' 8 Bytes
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
  
Private Type WIN32_FIND_DATA ' 318 Bytes
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved¯ As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Function FileExist(strFileName As String) As Boolean

Dim lpFindFileData As WIN32_FIND_DATA
Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)
    If hFindFirst > 0 Then
        FindClose hFindFirst
        FileExist = True
    Else
        FileExist = False
    End If
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    If lRetVal = 0 Then 'The file does not exist, first create it!
        Open sLongFileName For Random As #1
        Close #1
        lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
        'Now another try!
        Kill (sLongFileName)
        'Delete file now!
    End If
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Private Function Has_Space(sName As String) As Boolean
    Dim b As Boolean
    Dim i As Long
        
    b = False 'not yet any spaces found
    i = InStr(sName, " ")
    If i <> 0 Then b = True
    Has_Space = b
End Function
  
Public Sub WaveReset()
    Dim rtn As String
    Dim i As Long
    
    rtn = Space$(260)
    'Close any MCI operations from previous VB programs
    i = mciSendString("close all", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Closing all MCI operations failed!")
    'Open a new WAV with MCI Command...
    i = mciSendString("open new type waveaudio alias capture", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Opening new wave failed!")
End Sub

Public Sub WaveSet()
    Dim rtn As String
    Dim i As Long
    Dim settings As String
    Dim Alignment As Integer
       
    rtn = Space$(260)
  
    Alignment = Channels * Res / 8
    
    settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Res) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)

    'Samples Per Second that are supported:
    '11025     low quality
    '22050     medium quality
    '44100     high quality (CD music quality)
    'Bits per sample is 16 or 8
    'Channels are 1 (mono) or 2 (stereo)
 
    i = mciSendString("seek capture to start", rtn, Len(rtn), 0) 'Always start at the beginning
    If i <> 0 Then MsgBox ("Starting recording failed!")
    'You can use at least the following combinations
     
    ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 44100 channels 2 bytespersec 176400", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 44100 channels 1 bytespersec 88200", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 22050 channels 2 bytespersec 88200", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 22050 channels 1 bytespersec 44100", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 11025 channels 2 bytespersec 44100", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 11025 channels 1 bytespersec 22050", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 11025 channels 2 bytespersec 22050", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 11025 channels 1 bytespersec 11025", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 8000 channels 2 bytespersec 16000", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 8000 channels 1 bytespersec 8000", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 6000 channels 2 bytespersec 12000", rtn, Len(rtn), 0)
    ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 6000 channels 1 bytespersec 6000", rtn, Len(rtn), 0)
    
    i = mciSendString(settings, rtn, Len(rtn), 0)

    If i <> 0 Then MsgBox ("Settings for recording not consistent")
    ' If the combination is not supported you get an error!
 End Sub
 
 Public Sub WaveRecord()
    Dim rtn As String
    Dim i As Long
    Dim msg As String
    
    rtn = Space$(260)
    
    If WavMidi <> "" Then
  
        If WaRecIm Then MsgBox ("Midi file " & WavMidi & " will be recorded")
        i = mciSendString("open " & WavMidi & " type sequencer alias midi", rtn, Len(rtn), 0)
        If i <> 0 Then MsgBox ("Opening midi file failed!")

        i = mciSendString("play midi", rtn, Len(rtn), 0)  'Start the recording
        If i <> 0 Then MsgBox ("Playing midi file failed!")
    End If
   
    i = mciSendString("record capture", rtn, Len(rtn), 0)  'Start the recording
    If i <> 0 Then MsgBox ("Recording not possible, please restart your computer...")
 End Sub

Public Sub WaveSaveAs(sName As String)
   Dim rtn As String
   Dim i As Long
   
   'If file already exists then remove it
   
    If FileExist(sName) Then
        Kill (sName)
    End If
 
    'The mciSendString API call doesn't seem to like'
    'long filenames that have spaces in them, so we
    'will make another API call to get the short
    'filename version.
    'This is accomplished by the function GetShortName
            
    'MCI command to save the WAV file
     If Has_Space(sName) Then
        WavShFN = GetShortName(sName)
        WavLongFN = sName
        WavRenaNeces = True
        ' These are necessary in order to be able to rename file
        i = mciSendString("save capture " & WavShFN, rtn, Len(rtn), 0)
     Else
        i = mciSendString("save capture " & sName, rtn, Len(rtn), 0)
     End If
     If i <> 0 Then MsgBox ("Saving file failed, file name was: " & sName)
End Sub

Public Sub WaveStop()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("stop capture", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Stopping recording failed!")
    If WavMidi <> "" Then
        i = mciSendString("stop midi", rtn, Len(rtn), 0)
        If i <> 0 Then MsgBox ("Stopping playing midi file failed!")
    End If
End Sub

Public Sub WavePlay()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("play capture from 0", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Start playing failed!")
End Sub

Public Sub WaveStatus()
    Dim i As Long
    WavStatMsg = Space(255)
    i = mciSendString("status capture mode", WavStatMsg, 255, 0)
    If i <> 0 Then MsgBox ("Failure getting wave status...")
    WavStatMsg = "Recorder: " & WavStatMsg
End Sub

Public Sub WaveStatistics()
    Dim mssg As String * 255
    Dim i As Long
    i = mciSendString("set capture time format ms", 0&, 0, 0)
    If i <> 0 Then MsgBox ("Setting time format in milliseconds failed!")
    i = mciSendString("status capture length", mssg, 255, 0)
    mssg = CStr(CLng(mssg) / 1000)
    If i <> 0 Then MsgBox ("Finding length recording in milliseconds failed!")
    WavSticMsg = "Length recording " & Str(mssg) & " s"

    i = mciSendString("set capture time format bytes", 0&, 0, 0)
    If i <> 0 Then MsgBox ("Setting time format in bytes failed!")
    i = mciSendString("status capture length", mssg, 255, 0)
    If i <> 0 Then MsgBox ("Finding length recording in bytes failed!")
    WavSticMsg = WavSticMsg & " (" & Str(mssg) & " bytes)" & vbCrLf

    i = mciSendString("status capture channels", mssg, 255, 0)
    If i <> 0 Then MsgBox ("Finding number of channels failed!")
    If Str(mssg) = 1 Then
        WavSticMsg = WavSticMsg & "Mono - "
        ElseIf Str(mssg) = 2 Then
            WavSticMsg = WavSticMsg & "Stereo - "
    End If

    i = mciSendString("status capture bitspersample", mssg, 255, 0)
    If i <> 0 Then MsgBox ("Finding Res failed!")
    WavSticMsg = WavSticMsg & Str(mssg) & " bits - "

    i = mciSendString("status capture samplespersec", mssg, 255, 0)
    If i <> 0 Then MsgBox ("Finding sample rate failed!")
    WavSticMsg = WavSticMsg & Str(mssg) & " samples per second " & vbCrLf & vbCrLf
End Sub

Public Sub WaveClose()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("close capture", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Closing MCI failed!")
End Sub

Public Function WavePosition() As Long
    Dim rtn As String
    Dim i As Long
    Dim pos As String
    rtn = Space(255)
    pos = Space(255)
    
    i = mciSendString("set capture time format ms", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Setting format in milliseconds failed!")
    i = mciSendString("status capture position", pos, 255, 0)
    If i <> 0 Then MsgBox ("Finding position failed!")
    If i <> 0 Then MsgBox ("Error in position")
    WavePosition = CLng(pos)
End Function

Public Sub WavePlayFrom(Position As Long)
    Dim rtn As String
    Dim i As Long
    Dim pos As String
    pos = CStr(Position)
    i = mciSendString("set capture time format ms", 0&, 0, 0)
    If i <> 0 Then MsgBox ("Setting format in milliseconds failed!")
    i = mciSendString("play capture from " & pos, rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Playing from indicated position failed!")
    If i <> 0 Then MsgBox ("Play from position doesn't work....")
End Sub
