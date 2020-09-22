VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Karaoke Class Demo"
   ClientHeight    =   5220
   ClientLeft      =   1920
   ClientTop       =   1965
   ClientWidth     =   8520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8520
   Begin VB.CommandButton Command10 
      Caption         =   "About..."
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   3780
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "End"
      Height          =   375
      Left            =   7470
      TabIndex        =   23
      Top             =   3780
      Width           =   960
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   240
      LargeChange     =   50
      Left            =   6930
      Max             =   250
      Min             =   10
      SmallChange     =   5
      TabIndex        =   20
      Top             =   4635
      Value           =   10
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8190
      Top             =   5265
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   8460
      TabIndex        =   13
      Top             =   4980
      Width           =   8520
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3690
      TabIndex        =   12
      Top             =   3780
      Width           =   870
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4590
      TabIndex        =   11
      Top             =   3780
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Music Titles"
      Height          =   1185
      Left            =   90
      TabIndex        =   8
      Top             =   225
      Width           =   5325
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   5010
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Text"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5490
      TabIndex        =   7
      Top             =   3780
      Width           =   960
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   1000
      Left            =   6930
      SmallChange     =   100
      TabIndex        =   6
      Top             =   4320
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   2175
      Left            =   90
      ScaleHeight     =   2115
      ScaleWidth      =   8280
      TabIndex        =   5
      Top             =   1530
      Width           =   8340
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   2025
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2790
      TabIndex        =   4
      Top             =   3780
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1890
      TabIndex        =   3
      Top             =   3780
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   3780
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   3780
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   4455
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Velocity"
      Height          =   195
      Left            =   5940
      TabIndex        =   21
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Length:"
      Height          =   240
      Left            =   5490
      TabIndex        =   19
      Top             =   585
      Width           =   2985
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Quarter Notes:"
      Height          =   240
      Left            =   5490
      TabIndex        =   18
      Top             =   1035
      Width           =   2985
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IsLyric"
      Height          =   240
      Left            =   5490
      TabIndex        =   17
      Top             =   810
      Width           =   2985
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Music Length"
      Height          =   240
      Left            =   5490
      TabIndex        =   16
      Top             =   360
      Width           =   2985
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   195
      Left            =   5940
      TabIndex        =   10
      Top             =   4365
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   945
      TabIndex        =   1
      Top             =   4455
      Width           =   45
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Dim WithEvents Cl As Karaoke
Attribute Cl.VB_VarHelpID = -1
Dim X1 As Integer

Sub UpdateBar(Time)
    
    On Error Resume Next
    Command9.Left = ((Time * 100) / Cl.MusicLength) * (((Picture2.Width - Command9.Width) - 50) / 100)

End Sub

Private Sub Cl_Playing(ByVal CurrentWord As String, ByVal WordLen As Long, ByVal WordStart As Long, ByVal CurrentText As String, ByVal PreviousText As String, ByVal NextText As String)
    
    
    Picture1.Cls
    Picture1.ForeColor = &H0
    
    If CurrentWord <> "" Then
        Label6.Top = 300
        Label6.Visible = True
        Label6.Left = (Picture1.TextWidth(Left(CurrentText, WordStart)) + (Picture1.Width / 2) - (Picture1.TextWidth(CurrentText) / 2)) + ((Picture1.TextWidth(CurrentWord) / 2) - (Label6.Width / 2))
    Else
        Label6.Visible = False
    End If
    Picture1.CurrentY = 600
    Picture1.CurrentX = (Picture1.Width / 2) - (Picture1.TextWidth(CurrentText) / 2)
    Picture1.Print CurrentText
    
    Picture1.ForeColor = &HFF
    Picture1.CurrentY = 600
    Picture1.CurrentX = Picture1.TextWidth(Left(CurrentText, WordStart)) + (Picture1.Width / 2) - (Picture1.TextWidth(CurrentText) / 2)
    Picture1.Print CurrentWord

    Picture1.ForeColor = &H0
    Picture1.CurrentY = 1100
    Picture1.CurrentX = (Picture1.Width / 2) - (Picture1.TextWidth(NextText) / 2)
    Picture1.Print NextText
End Sub

Private Sub Cl_Status(ByVal Status As Long)

    If Status = 526 Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
        Label6.Visible = False
    End If

End Sub


Private Sub Cl_TimePosition(ByVal CurrentTime As Long)
    
    UpdateBar CurrentTime

End Sub


Private Sub Command1_Click()

    On Error Resume Next
    Picture1.Cls
    Command9.Left = 0
    
    X = ShowOpen(Me, "Karaoke Files (*.kar)|*.kar", Caption, App.Path)
    If Trim(X) = "" Then Exit Sub
    Label1 = X

    If Cl.OpenDevice(Label1) Then
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command6.Enabled = True
        Command7.Enabled = True
        Command8.Enabled = True
    Else
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command6.Enabled = False
        Command7.Enabled = False
        Command8.Enabled = False
        Exit Sub
    End If
    
    For T = 1 To Cl.Titles.Count
        Tit = Tit & Cl.Titles(T) & Chr(10)
    Next
    Label4 = Tit
    
    
    
    Vle = (Cl.GetMusicVolume / 2) - 10
    If Vle < 0 Then Vle = 0
    HScroll1 = Vle

    Label7 = "Music Length: " & Cl.MusicLength & " millisseconds"
    Label8 = "IsLyric = " & Cl.IsLyric
    Label9 = "Quarter Notes = " & Cl.QuarterNote
    Label10 = "Time Length: " & Cl.MusicTimeLength
        
    HScroll2 = Cl.Velocity
    
End Sub


Private Sub Command10_Click()

    frmAbout.Show
    Set frmAbout = Nothing
    

End Sub


Private Sub Command2_Click()

    Cl.Play

End Sub

Private Sub Command3_Click()

    Cl.Pause

End Sub

Private Sub Command4_Click()

    Cl.StopPlay

End Sub

Private Sub Command5_Click()
    
    Cl.CloseDevice
    DoEvents
    End
    
End Sub

Private Sub Command6_Click()

    MsgBox Cl.TextLyric

End Sub

Private Sub Command7_Click()

    Cl.SeekEnd

End Sub

Private Sub Command8_Click()

    Cl.SeekIni

End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Cl.StopPlay
    Picture1.Cls
    X1 = X

End Sub


Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Picture2_MouseMove Button, Shift, X + Command9.Left, Y
    

End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Cl.SetPosition (Cl.MusicLength / 100) * (Command9.Left / (((Picture2.Width - Command9.Width) - 50) / 100))
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Set Cl = New Karaoke

End Sub

Private Function ShowOpen(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        a = GetOpenFileName(ofn)

        If (a) Then
            ShowOpen = Trim$(ofn.lpstrFile)
        Else
            ShowOpen = ""
        End If

End Function

Private Sub HScroll1_Change()
    Vlr& = HScroll1.Value
    Cl.SetMusicVolume Vlr& * 2

End Sub

Private Sub HScroll1_Scroll()

    Vlr& = HScroll1.Value
    Cl.SetMusicVolume CLng(HScroll1.Value) * 2

End Sub


Private Sub HScroll2_Change()

    Cl.Velocity = HScroll2.Value

End Sub

Private Sub HScroll2_Scroll()

    Cl.Velocity = HScroll2.Value

End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        X = IIf((X - X1) < 0, 0, X - X1)
        X = IIf((X - X1) > ((Picture2.Width - Command9.Width) - 90), (Picture2.Width - Command9.Width) - 50, X)
        Command9.Left = X
    End If

End Sub


Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
a = B
End Sub


Private Sub Timer1_Timer()
        
    Static Tp
    Select Case Tp
        Case 300
            Tp = 282
        Case 282
            Tp = 248
        Case 248
            Tp = 214
        Case 214
            Tp = 180
        Case 180
            Tp = 213
        Case 213
            Tp = 247
        Case 247
            Tp = 281
        Case 281
            Tp = 300
        Case Else
            Tp = 300
    End Select
    
    Label6.Top = Tp

End Sub


