VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Karaoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long


Private mvarTitles As Collection
Private mvarIsLyric As Boolean
Private mvarQuarterNote As String
Public Event Playing(ByVal CurrentWord As String, ByVal WordLen As Long, ByVal WordStart As Long, ByVal CurrentText As String, ByVal PreviousText As String, ByVal NextText As String)
Public Event TimePosition(ByVal CurrentTime As Long)
Public Event Status(ByVal Status As Long)
Private mvarVelocity As Integer
Private mvarMusicLength As Long
Private mvarMusicTimeLength As String
Private mvarGetMusicVolume As Long
Private mvarTextLyric As String


Private Type MCI_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
End Type
Private OpenParms As MCI_OPEN_PARMS

Private Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Private PlayParms As MCI_PLAY_PARMS

Private Type MCI_STATUS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwItem As Long
        dwTrack As Integer
End Type
Private StatusParms As MCI_STATUS_PARMS

Private Type MCI_SEQ_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        dwTempo As Long
        dwPort As Long
        dwSlave As Long
        dwMaster As Long
        dwOffset As Long
End Type
Private SeqSetParms As MCI_SEQ_SET_PARMS

Private Type MCI_SEEK_PARMS
        dwCallback As Long
        dwTo As Long
End Type
Private SeekParms As MCI_SEEK_PARMS

Private Type MCI_GENERIC_PARMS
        dwCallback As Long
End Type
Private GenParms As MCI_GENERIC_PARMS

Private Type MCI_INFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
End Type
Private InfoParms As MCI_INFO_PARMS


Private Const MCI_OPEN = &H803
Private Const MCI_PLAY = &H806
Private Const MCI_STATUS = &H814
Private Const MCI_SET = &H80D
Private Const MCI_WAIT = &H2&
Private Const MCI_NOTIFY = &H1&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_CLOSE = &H804
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_MODE = &H4&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_SEQ_FORMAT_SONGPTR = &H4001
Private Const MCI_SEQ_SET_TEMPO = &H10000
Private Const MCI_SEQ_STATUS_TEMPO = &H4002&
Private Const MCI_FORMAT_MILLISECONDS = 0
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_FROM = &H4&
Private Const MCI_PAUSE = &H809
Private Const MCI_SEEK = &H807
Private Const MCI_SEEK_TO_END = &H200&
Private Const MCI_SEEK_TO_START = &H100&
Private Const MCI_STOP = &H808
Private Const MCI_STRING_OFFSET = 512
Private Const MCI_MODE_NOT_READY = (MCI_STRING_OFFSET + 12)
Private Const MCI_MODE_OPEN = (MCI_STRING_OFFSET + 18)
Private Const MCI_MODE_PAUSE = (MCI_STRING_OFFSET + 17)
Private Const MCI_MODE_PLAY = (MCI_STRING_OFFSET + 14)
Private Const MCI_MODE_RECORD = (MCI_STRING_OFFSET + 15)
Private Const MCI_MODE_SEEK = (MCI_STRING_OFFSET + 16)
Private Const MCI_MODE_STOP = (MCI_STRING_OFFSET + 13)

Private Const MCIERR_BASE = 256
Private Const MCIERR_BAD_CONSTANT = (MCIERR_BASE + 34)
Private Const MCIERR_BAD_INTEGER = (MCIERR_BASE + 14)
Private Const MCIERR_BAD_TIME_FORMAT = (MCIERR_BASE + 37)
Private Const MCIERR_CANNOT_LOAD_DRIVER = (MCIERR_BASE + 10)
Private Const MCIERR_CANNOT_USE_ALL = (MCIERR_BASE + 23)
Private Const MCIERR_CREATEWINDOW = (MCIERR_BASE + 91)
Private Const MCIERR_CUSTOM_DRIVER_BASE = (MCIERR_BASE + 256)
Private Const MCIERR_DEVICE_LENGTH = (MCIERR_BASE + 54)
Private Const MCIERR_DEVICE_LOCKED = (MCIERR_BASE + 32)
Private Const MCIERR_DEVICE_NOT_INSTALLED = (MCIERR_BASE + 50)
Private Const MCIERR_DEVICE_NOT_READY = (MCIERR_BASE + 20)
Private Const MCIERR_DEVICE_OPEN = (MCIERR_BASE + 9)
Private Const MCIERR_DEVICE_ORD_LENGTH = (MCIERR_BASE + 55)
Private Const MCIERR_DEVICE_TYPE_REQUIRED = (MCIERR_BASE + 31)
Private Const MCIERR_DRIVER = (MCIERR_BASE + 22)
Private Const MCIERR_DRIVER_INTERNAL = (MCIERR_BASE + 16)
Private Const MCIERR_DUPLICATE_ALIAS = (MCIERR_BASE + 33)
Private Const MCIERR_DUPLICATE_FLAGS = (MCIERR_BASE + 39)
Private Const MCIERR_EXTENSION_NOT_FOUND = (MCIERR_BASE + 25)
Private Const MCIERR_EXTRA_CHARACTERS = (MCIERR_BASE + 49)
Private Const MCIERR_FILE_NOT_FOUND = (MCIERR_BASE + 19)
Private Const MCIERR_FILE_READ = (MCIERR_BASE + 92)
Private Const MCIERR_FILE_NOT_SAVED = (MCIERR_BASE + 30)
Private Const MCIERR_FILE_WRITE = (MCIERR_BASE + 93)
Private Const MCIERR_FILENAME_REQUIRED = (MCIERR_BASE + 48)
Private Const MCIERR_FLAGS_NOT_COMPATIBLE = (MCIERR_BASE + 28)
Private Const MCIERR_GET_CD = (MCIERR_BASE + 51)
Private Const MCIERR_HARDWARE = (MCIERR_BASE + 6)
Private Const MCIERR_ILLEGAL_FOR_AUTO_OPEN = (MCIERR_BASE + 47)
Private Const MCIERR_INTERNAL = (MCIERR_BASE + 21)
Private Const MCIERR_INVALID_DEVICE_ID = (MCIERR_BASE + 1)
Private Const MCIERR_INVALID_DEVICE_NAME = (MCIERR_BASE + 7)
Private Const MCIERR_INVALID_FILE = (MCIERR_BASE + 40)
Private Const MCIERR_MISSING_COMMAND_STRING = (MCIERR_BASE + 11)
Private Const MCIERR_MISSING_DEVICE_NAME = (MCIERR_BASE + 36)
Private Const MCIERR_MISSING_PARAMETER = (MCIERR_BASE + 17)
Private Const MCIERR_MISSING_STRING_ARGUMENT = (MCIERR_BASE + 13)
Private Const MCIERR_MULTIPLE = (MCIERR_BASE + 24)
Private Const MCIERR_MUST_USE_SHAREABLE = (MCIERR_BASE + 35)
Private Const MCIERR_NEW_REQUIRES_ALIAS = (MCIERR_BASE + 43)
Private Const MCIERR_NO_CLOSING_QUOTE = (MCIERR_BASE + 38)
Private Const MCIERR_NO_ELEMENT_ALLOWED = (MCIERR_BASE + 45)
Private Const MCIERR_NO_INTEGER = (MCIERR_BASE + 56)
Private Const MCIERR_NO_WINDOW = (MCIERR_BASE + 90)
Private Const MCIERR_NONAPPLICABLE_FUNCTION = (MCIERR_BASE + 46)
Private Const MCIERR_NOTIFY_ON_AUTO_OPEN = (MCIERR_BASE + 44)
Private Const MCIERR_NULL_PARAMETER_BLOCK = (MCIERR_BASE + 41)
Private Const MCIERR_OUT_OF_MEMORY = (MCIERR_BASE + 8)
Private Const MCIERR_OUTOFRANGE = (MCIERR_BASE + 26)
Private Const MCIERR_PARAM_OVERFLOW = (MCIERR_BASE + 12)
Private Const MCIERR_PARSER_INTERNAL = (MCIERR_BASE + 15)
Private Const MCIERR_SEQ_DIV_INCOMPATIBLE = (MCIERR_BASE + 80)
Private Const MCIERR_SEQ_NOMIDIPRESENT = (MCIERR_BASE + 87)
Private Const MCIERR_SEQ_PORT_INUSE = (MCIERR_BASE + 81)
Private Const MCIERR_SEQ_PORT_MAPNODEVICE = (MCIERR_BASE + 83)
Private Const MCIERR_SEQ_PORT_MISCERROR = (MCIERR_BASE + 84)
Private Const MCIERR_SEQ_PORT_NONEXISTENT = (MCIERR_BASE + 82)
Private Const MCIERR_SEQ_PORTUNSPECIFIED = (MCIERR_BASE + 86)
Private Const MCIERR_SEQ_TIMER = (MCIERR_BASE + 85)
Private Const MCIERR_SET_CD = (MCIERR_BASE + 52)
Private Const MCIERR_SET_DRIVE = (MCIERR_BASE + 53)
Private Const MCIERR_UNNAMED_RESOURCE = (MCIERR_BASE + 42)
Private Const MCIERR_UNRECOGNIZED_COMMAND = (MCIERR_BASE + 5)
Private Const MCIERR_UNRECOGNIZED_KEYWORD = (MCIERR_BASE + 3)
Private Const MCIERR_UNSUPPORTED_FUNCTION = (MCIERR_BASE + 18)
Private Const MMSYSERR_BASE = 0
Private Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)
Private Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Private Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)
Private Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)
Private Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)
Private Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)
Private Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)
Private Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Private Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)
Private Const MMSYSERR_NOERROR = 0
Private Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Private Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)
Private Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)
Private Const MCI_DEVTYPE_SEQUENCER = 523



Dim TxtString As String, Bpm As Single, ppqn, Clk, Den, Quarter As Double
Dim Xini As Long, Ret As Long, IsPlaying As Boolean
Dim Lyr() As Lyric, Frase() As String


Private Type Lyric
    TempoTotal As Double
    TempoAtual As Double
    TxtString As String
    TxtStringLen As Integer
    TextStart As Long
    FraseIndex As Integer
End Type



Private Property Let TextLyric(ByVal vData As String)
    mvarTextLyric = vData
End Property


Public Property Get TextLyric() As String
    GetTextLyric
    TextLyric = mvarTextLyric
End Property





Public Property Get GetMusicVolume() As Long
    midiOutGetVolume -1, Volum&
    GetMusicVolume = Volum& And &HFFFF&
End Property

Public Sub SetMusicVolume(Value As Long)
    
    If Value > 65535 Then Value = 65535
    If Value < 5000 Then Value = 5000
    Vlm = Val("&H" & Hex(Value) & Hex(Value))
    Call midiOutSetVolume(-1, Vlm)

End Sub

Public Property Let MusicTimeLength(ByVal vData As String)
    mvarMusicTimeLength = vData
End Property


Public Property Get MusicTimeLength() As String
    MusicTimeLength = mvarMusicTimeLength
End Property



Public Property Let MusicLength(ByVal vData As Long)
    mvarMusicLength = vData
End Property


Public Property Get MusicLength() As Long
    MusicLength = mvarMusicLength
End Property



Public Property Let Velocity(ByVal vData As Integer)
    mvarVelocity = vData
    SetTempo
End Property


Public Property Get Velocity() As Integer
    Velocity = mvarVelocity
End Property





Private Sub NotifyPlay()
    
    On Error Resume Next
    If Not IsLyric Then Exit Sub
    Do
        If Not IsPlaying Then Exit Sub
        DoEvents
        StatusParms.dwItem = MCI_STATUS_POSITION
        Call mciSendCommand(OpenParms.wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_WAIT, StatusParms)
        Tempo = StatusParms.dwReturn
        If Xini = 0 Then Xini = 1
        DoEvents
        If Xini + 1 <= UBound(Lyr) Then
            If Tempo >= Lyr(Xini).TempoTotal Then
                Xini = Xini + 1
                OldTextLyr = ""
                NextTextLyr = ""
                OldTextLyr = Frase(Lyr(Xini).FraseIndex - 1)
                NextTextLyr = Frase(Lyr(Xini).FraseIndex + 1)
                RaiseEvent Playing(Lyr(Xini).TxtString, Lyr(Xini).TxtStringLen, Lyr(Xini).TextStart, Frase(Lyr(Xini).FraseIndex), OldTextLyr, NextTextLyr)
            End If
        End If
        RaiseEvent TimePosition(Tempo)
        DoEvents
        StatusParms.dwItem = MCI_STATUS_MODE
        Call mciSendCommand(OpenParms.wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_WAIT, StatusParms)
        If StatusParms.dwReturn = MCI_MODE_STOP Then
            IsPlaying = False
        End If
        RaiseEvent Status(StatusParms.dwReturn)
    Loop

End Sub

Private Property Let QuarterNote(ByVal vData As String)
    mvarQuarterNote = vData
End Property


Public Property Get QuarterNote() As String
    QuarterNote = mvarQuarterNote
End Property



Private Property Let IsLyric(ByVal vData As Boolean)
    mvarIsLyric = vData
End Property


Public Property Get IsLyric() As Boolean
    IsLyric = mvarIsLyric
End Property



Public Property Set Titles(ByVal vData As Object)
    Set mvarTitles = vData
End Property


Public Property Get Titles() As Collection
    Set Titles = mvarTitles
End Property




Sub CloseDevice()
    ReDim Lyr(0)
    ReDim Frase(0)
    Close #1
    Reset
    Xini = 0
    Call mciSendCommand(OpenParms.wDeviceID, MCI_CLOSE, &H0, vbNull)
    RaiseEvent Status(MCI_MODE_STOP)
End Sub
Private Sub CreateType()
    If Not IsLyric Then Exit Sub
    Dim N As Integer, OldPause As Double, C As Integer, K As Integer, W As Integer
    K = 1: W = 0: Nx = 0: Ni = 0
    Set Titles = New Collection
    Do
        Nx = Nx + 1
        NewText = ""
        K = InStr(K, TxtString, Chr(255))
        If K = 0 Then Exit Sub
        CompText = Val(Asc(Mid(TxtString, K + 2, 1)))
        Select Case Val(Asc(Mid(TxtString, K + 1, 1)))
            Case 1
                If Mid(TxtString, K + 3, 1) <> "@" Then
                    NewText = Mid(TxtString, K + 3, CompText)
                Else
                    If Mid(TxtString, K + 3, 2) = "@T" Then
                        Titles.Add StrConv(Mid(TxtString, K + 5, CompText - 2), vbProperCase)
                    End If
                End If
            Case 47
                    Exit Sub
        End Select
        Pausa1 = Val(Asc(Mid(TxtString, (K + 3) + CompText, 1)))
        Pausa2 = Val(Asc(Mid(TxtString, (K + 4) + CompText, 1)))
        Pausa3 = Val(Asc(Mid(TxtString, (K + 5) + CompText, 1)))
        Pausa4 = Val(Asc(Mid(TxtString, (K + 6) + CompText, 1)))
        If Pausa1 = 255 Then Pausa1 = 0
        If Pausa2 = 255 Then Pausa2 = 0
        If Pausa3 = 255 Then Pausa3 = -1
        If Pausa4 = 255 Then Pausa4 = -1
        K = K + 1
        Select Case Left(NewText, 1)
            Case Is = "/", "\"
                C = 0
                NewText = Right(NewText, Len(NewText) - 1)
                CompText = CompText - 1
                W = W + 1
            Case "@"
                NewText = ""
            Case Else
        End Select
        ReDim Preserve Lyr(Nx)

        Valor = 0
        If Pausa1 >= 127 Then
            If Pausa3 >= 0 Then
                If Pausa4 >= 0 Then
                    Valor = Valor And &H7F
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa1 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa2 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa3 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa4 And &H7F)
                    Pausa = Valor * Quarter
                Else
                    Valor = Valor And &H7F
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa1 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa2 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa3 And &H7F)
                    Pausa = Valor * Quarter
                End If
            Else
                Pausa = (((Pausa1 And &H7F) * 128) Or (Pausa2 And &H7F)) * Quarter
            End If
        Else
            Pausa = Pausa1 * Quarter
        End If
        OldPause = OldPause + Pausa
        Lyr(Nx).TxtString = ""
        Lyr(Nx).TxtStringLen = 0
        Lyr(Nx).TextStart = 0
        Lyr(Nx).TxtString = NewText
        Lyr(Nx).TxtStringLen = Len(NewText)
        Lyr(Nx).TextStart = C
        Lyr(Nx).TempoAtual = Pausa
        Lyr(Nx).TempoTotal = OldPause
        Lyr(Nx).FraseIndex = W
        C = C + CompText
    Loop

End Sub
Private Sub GetTextLyric()
    If Not IsLyric Then Exit Sub
    j = 1
    N = 0
    ReDim Frase(0)
    Do
        K = InStr(j, TxtString, Chr(255) & Chr(1))
        If K = 0 Then Exit Do
        j = K + 1
        CompText = Val(Asc(Mid(TxtString, K + 2, 1)))
        NewText = Mid(TxtString, K + 3, CompText)
        NovaFrase = False
        Select Case Left(NewText, 1)
            Case Is = "/", "\"
                NewText = Chr(13) & Chr(10) & Right(NewText, Len(NewText) - 1)
                N = N + 1
                ReDim Preserve Frase(N)
            Case "@"
                NewText = " "
        End Select
        Tx = Tx & NewText
        If Left(NewText, 2) = Chr(13) & Chr(10) Then
            NewText = Right(NewText, Len(NewText) - 2)
        End If
        If Trim(NewText) <> "" Then Frase(N) = Frase(N) & NewText
    Loop
    TextLyric = Tx
End Sub
Private Sub OpenFile(FileName As String)
        
    On Error Resume Next
    Close #1
    Open FileName For Binary As #1
    TxtString = Space(FileLen(FileName))
    Get #1, , TxtString
    Track = TxtString
    
    K = InStr(1, TxtString, Chr(255) & Chr(88))
    If K = 0 Then Exit Sub
    Num = Asc(Mid(TxtString, K + 3, 1))
    Den = Asc(Mid(TxtString, K + 4, 1))
    Select Case Den
        Case 1: Den = 2
        Case 2: Den = 4
        Case 3: Den = 8
        Case 4: Den = 16
        Case 5: Den = 32
        Case 6: Den = 64
        Case 7: Den = 128
        Case 8: Den = 256
        Case 9: Den = 512
        Case 10: Den = 1024
    End Select
    Clk = Asc(Mid(TxtString, K + 5, 1)) / Den
    
    K = InStr(1, TxtString, Chr(255) & Chr(81))
    If K = 0 Then Exit Sub
    
    n1 = Format(Hex(Asc(Mid(TxtString, K + 3, 1))), "00")
    n2 = Format(Hex(Asc(Mid(TxtString, K + 4, 1))), "00")
    N3 = Format(Hex(Asc(Mid(TxtString, K + 5, 1))), "00")
    If Len(n1) = 1 Then n1 = "0" & n1
    If Len(n2) = 1 Then n2 = "0" & n2
    If Len(N3) = 1 Then N3 = "0" & N3
    
    
    T1 = Format(Hex(Asc(Mid(TxtString, 13, 1))), "00")
    T2 = Format(Hex(Asc(Mid(TxtString, 14, 1))), "00")
    If Len(T1) = 1 Then T1 = "0" & T1
    If Len(T2) = 1 Then T2 = "0" & T2
    
    TimeBase = CDec("&H" & T1 & T2) / 4
    
    Bpm = Format(60000000 / CDec("&H" & n1 & n2 & N3), "0.00")
    ppqn = CDec("&H" & n1 & n2 & N3)
    Quarter = (ppqn / TimeBase) / 4000
    QuarterNote = Num & "/" & Den
    
    K = InStr(1, TxtString, Chr(255) & Chr(3) & Chr(5) & "W")
    If K = 0 Then
        K = InStr(1, TxtString, "@T")
        If K = 0 Then
            IsLyric = False
            Exit Sub
        End If
    End If
    IsLyric = True
    
    For T = K To 1 Step -1
        If UCase(Mid(TxtString, T, 4)) = "MTRK" Then
            IniTrack = T
            Exit For
        End If
    Next

    V1 = Format(Hex(Asc(Mid(TxtString, IniTrack + 4, 1))), "00")
    V2 = Format(Hex(Asc(Mid(TxtString, IniTrack + 5, 1))), "00")
    V3 = Format(Hex(Asc(Mid(TxtString, IniTrack + 6, 1))), "00")
    V4 = Format(Hex(Asc(Mid(TxtString, IniTrack + 7, 1))), "00")
    If Len(V1) = 1 Then V1 = "0" & V1
    If Len(V2) = 1 Then V2 = "0" & V2
    If Len(V3) = 1 Then V3 = "0" & V3
    If Len(V4) = 1 Then V4 = "0" & V4


    TxtString = Space(CDec("&H" & Va & V2 & V3 & V4) + 100)
    Track = Right(Track, Len(Track) - (IniTrack + CDec("&H" & Va & V2 & V3 & V4)))
    
    Seek #1, IniTrack + 8
    Get #1, , TxtString
    TxtString = Right(TxtString, Len(TxtString) - 0)
    Close #1
    CreateType
    GetTextLyric

End Sub

Public Function OpenDevice(FileName As String) As Boolean
        
    CloseDevice
    ReDim Lyr(0)
    ReDim Frase(0)
    Set Titles = Nothing
    OpenFile FileName
    DoEvents
    
    OpenParms.lpstrDeviceType = "sequencer"
    OpenParms.lpstrElementName = FileName
    Ret = mciSendCommand(1, MCI_OPEN, MCI_OPEN_ELEMENT Or MCI_OPEN_TYPE Or MCI_NOTIFY, OpenParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        OpenDevice = False
        Exit Function
    End If
    
    SeqSetParms.dwTimeFormat = MCI_FORMAT_MILLISECONDS
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_SET, MCI_SET_TIME_FORMAT Or MCI_SEQ_FORMAT_SONGPTR Or MCI_WAIT, SeqSetParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        OpenDevice = False
        Exit Function
    End If
    
    StatusParms.dwItem = MCI_STATUS_LENGTH
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_WAIT, StatusParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        OpenDevice = False
        Exit Function
    End If
    MusicLength = StatusParms.dwReturn
    
    Minutos = Int((MusicLength / 1000) / 60)
    Segundo = Int(MusicLength / 1000)
    If Segundo > 59 Then
        Segundo = Segundo - (Minutos * 60)
    End If
    MusicTimeLength = "00:" & Format(Minutos, "00") & ":" & Format(Segundo, "00")

    
    StatusParms.dwItem = MCI_SEQ_STATUS_TEMPO
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_STATUS, MCI_STATUS_ITEM, StatusParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        OpenDevice = False
        Exit Function
    End If
    Velocity = StatusParms.dwReturn
    If Velocity > 200 Then Velocity = 200
    
    OpenDevice = True
    RaiseEvent Status(MCI_MODE_OPEN)
End Function




Private Function GetErr(Erro) As String

    'Put the error message here
    Select Case Erro
        Case MCIERR_BAD_CONSTANT
            Msg = "Constante errada"
        Case MCIERR_BAD_INTEGER
            Msg = "Erro de número inteiro"
        Case MCIERR_BAD_TIME_FORMAT
            Msg = "Tempo inválido."
        Case MCIERR_CANNOT_LOAD_DRIVER
            Msg = "Não foi possível carregar o driver de som."
        Case MCIERR_CANNOT_USE_ALL
            Msg = "Não é possível o uso de todos os dispositivos."
        Case MCIERR_CREATEWINDOW
            Msg = "Erro na criação da janela."
        Case MCIERR_CUSTOM_DRIVER_BASE
            Msg = "Erro no driver do usuário."
        Case MCIERR_DEVICE_LENGTH
            Msg = "Erro na especificação do dispositivo de som."
        Case MCIERR_DEVICE_LOCKED
            Msg = "Dispositivo de som indisponível."
        Case MCIERR_DEVICE_NOT_INSTALLED
            Msg = "Dispositivo de som não instalado."
        Case MCIERR_DEVICE_NOT_READY
            Msg = "O dispositivo de som não está respondendo."
        Case MCIERR_DEVICE_OPEN
            Msg = "O dispositivo de som está aberto por outro aplicativo."
        Case MCIERR_DEVICE_ORD_LENGTH
            Msg = "Erro ao conectar ao dispositivo de som."
        Case MCIERR_DEVICE_TYPE_REQUIRED
            Msg = "Dispositivo de som requerido."
        Case MCIERR_DRIVER
            Msg = "Erro no driver do dispositivo de som."
        Case MCIERR_DRIVER_INTERNAL
            Msg = "Erro no driver interno do dispositivo de som."
        Case MCIERR_DUPLICATE_ALIAS
            Msg = "Erro desconhecido."
        Case MCIERR_DUPLICATE_FLAGS
            Msg = "Erro desconhecido."
        Case MCIERR_EXTENSION_NOT_FOUND
            Msg = "Extensão de arquivo desconhecido."
        Case MCIERR_EXTRA_CHARACTERS
            Msg = "Extra carateres."
        Case MCIERR_FILE_NOT_FOUND
            Msg = "Arquivo não encontrado."
        Case MCIERR_FILE_READ
            Msg = "Erro de leitura no arquivo."
        Case MCIERR_FILE_NOT_SAVED
            Msg = "Arquivo ainda não gravado."
        Case MCIERR_FILE_WRITE
            Msg = "Erro de gravação no arquivo."
        Case MCIERR_FILENAME_REQUIRED
            Msg = "Nome do arquivo requerido."
        Case MCIERR_FLAGS_NOT_COMPATIBLE
            Msg = "Sinalizadores não compatíveis."
        Case MCIERR_GET_CD
            Msg = "Erro ao acessar o CD."
        Case MCIERR_HARDWARE
            Msg = "Erro de hardware."
        Case MCIERR_ILLEGAL_FOR_AUTO_OPEN
            Msg = "Erro ao abrir o arquivo."
        Case MCIERR_INTERNAL
            Msg = "Erro interno."
        Case MCIERR_INVALID_DEVICE_ID
            Msg = "Dispositivo de som inválido."
        Case MCIERR_INVALID_DEVICE_NAME
            Msg = "Dispositivo de som inválido."
        Case MCIERR_INVALID_FILE
            Msg = "Arquivo inválido."
        Case MCIERR_MISSING_COMMAND_STRING
            Msg = "Erro de comando."
        Case MCIERR_MISSING_DEVICE_NAME
            Msg = "Dispositivo de som inválido."
        Case MCIERR_MISSING_PARAMETER
            Msg = "Parâmetro inválido."
        Case MCIERR_MISSING_STRING_ARGUMENT
            Msg = "Argumento inválido."
        Case MCIERR_MULTIPLE
            Msg = "Erros mútiplos."
        Case MCIERR_MUST_USE_SHAREABLE
            Msg = "Dispositivo não compartilhado."
        Case MCIERR_NEW_REQUIRES_ALIAS
            Msg = "Erro interno."
        Case MCIERR_NO_CLOSING_QUOTE
            Msg = "Erro interno."
        Case MCIERR_NO_ELEMENT_ALLOWED
            Msg = "Erro ao ler o arquivo de som."
        Case MCIERR_NO_INTEGER
            Msg = "Erro interno."
        Case MCIERR_NO_WINDOW
            Msg = "Erro interno."
        Case MCIERR_NONAPPLICABLE_FUNCTION
            Msg = "Erro interno."
        Case MCIERR_NOTIFY_ON_AUTO_OPEN
            Msg = "Erro interno."
        Case MCIERR_NULL_PARAMETER_BLOCK
            Msg = "Erro interno."
        Case MCIERR_OUT_OF_MEMORY
            Msg = "Sem memória"
        Case MCIERR_OUTOFRANGE
            Msg = "Erro interno."
        Case MCIERR_PARAM_OVERFLOW
            Msg = "Erro interno."
        Case MCIERR_PARSER_INTERNAL
            Msg = "Erro interno."
        Case MCIERR_SEQ_DIV_INCOMPATIBLE
            Msg = "Divisão de tempo incompatível."
        Case MCIERR_SEQ_NOMIDIPRESENT
            Msg = "Este dispositivo não pode executar sons MIDI."
        Case MCIERR_SEQ_PORT_INUSE
            Msg = "Dispositivo em uso."
        Case MCIERR_SEQ_PORT_MAPNODEVICE
            Msg = "Este dispositivo não pode executar sons MIDI."
        Case MCIERR_SEQ_PORT_MISCERROR
            Msg = "Erro no dispositivo de som"
        Case MCIERR_SEQ_PORT_NONEXISTENT
            Msg = "Este dispositivo não pode executar sons MIDI."
        Case MCIERR_SEQ_PORTUNSPECIFIED
            Msg = "Este dispositivo não pode executar sons MIDI."
        Case MCIERR_SEQ_TIMER
            Msg = "Erro interno."
        Case MCIERR_SET_CD
            Msg = "Erro interno."
        Case MCIERR_SET_DRIVE
            Msg = "Erro interno."
        Case MCIERR_UNNAMED_RESOURCE
            Msg = "Erro interno."
        Case MCIERR_UNRECOGNIZED_COMMAND
            Msg = "Comando não reconhecido."
        Case MCIERR_UNRECOGNIZED_KEYWORD
            Msg = "Comando não reconhecido."
        Case MCIERR_UNSUPPORTED_FUNCTION
            Msg = "Função não suportada."
        Case MMSYSERR_ALLOCATED
            Msg = "Erro interno."
        Case MMSYSERR_BADDEVICEID
            Msg = "Erro no dispositivo de som."
        Case MMSYSERR_BADERRNUM
            Msg = "Erro interno."
        Case MMSYSERR_ERROR
            Msg = "Erro interno."
        Case MMSYSERR_HANDLEBUSY
            Msg = "Erro interno."
        Case MMSYSERR_INVALFLAG
            Msg = "Erro interno."
        Case MMSYSERR_INVALHANDLE
            Msg = "Erro interno."
        Case MMSYSERR_INVALIDALIAS
            Msg = "Erro interno."
        Case MMSYSERR_INVALPARAM
            Msg = "Par6ametro inválido."
        Case MMSYSERR_LASTERROR
            Msg = "Erro interno."
        Case MMSYSERR_NODRIVER
            Msg = "Erro interno."
        Case MMSYSERR_NOMEM
            Msg = "Sem memória."
        Case MMSYSERR_NOTENABLED
            Msg = "Erro no dispositivo."
        Case MMSYSERR_NOTSUPPORTED
            Msg = "Erro interno."
        Case Else
            Msg = "Erro desconhecido."
    End Select
    GetErr = Msg

End Function



Public Sub Pause()
    
    If Not IsPlaying Then Exit Sub
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_PAUSE, MCI_WAIT, GenParms)
    RaiseEvent Status(MCI_MODE_PAUSE)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        Exit Sub
    End If

End Sub


Public Sub Play()
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_PLAY, &H0, PlayParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        Exit Sub
    End If
    IsPlaying = True
    NotifyPlay
End Sub


Public Sub SeekEnd()
    
    On Error Resume Next
    Call mciSendCommand(OpenParms.wDeviceID, MCI_SEEK, MCI_SEEK_TO_END, SeekParms)
    RaiseEvent Status(MCI_MODE_SEEK)
    Xini = UBound(Lyr)
    Play
    
End Sub
Public Sub SeekIni()
    
    Call mciSendCommand(OpenParms.wDeviceID, MCI_SEEK, MCI_SEEK_TO_START, SeekParms)
    RaiseEvent Status(MCI_MODE_SEEK)
    Xini = 0

End Sub

Public Sub SetPosition(ByVal nMilliseconds As Long)
    
    On Error Resume Next
    PlayParms.dwFrom = nMilliseconds
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_PLAY, MCI_FROM, PlayParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        Exit Sub
    End If
    For T = 1 To UBound(Lyr) - 1
        If Lyr(T).TempoTotal >= (nMilliseconds - 100) Then
            Xini = T
            Exit For
        End If
    Next
    Play
    

End Sub


Private Sub SetTempo()
    
    SeqSetParms.dwTempo = Velocity
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_SET, MCI_SEQ_SET_TEMPO, SeqSetParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        Exit Sub
    End If

End Sub


Public Sub StopPlay()
    Ret = mciSendCommand(OpenParms.wDeviceID, MCI_STOP, &H0, GenParms)
    If Ret <> 0 Then
        MsgBox GetErr(Ret), vbInformation, Caption
        Exit Sub
    End If
    IsPlaying = False
    RaiseEvent Status(MCI_MODE_STOP)
End Sub




