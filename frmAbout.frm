VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   2970
   ClientLeft      =   3750
   ClientTop       =   3375
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -90
      TabIndex        =   3
      Top             =   2385
      Width           =   5460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4005
      TabIndex        =   2
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Terms of Use"
      Height          =   1185
      Left            =   135
      TabIndex        =   1
      Top             =   1170
      Width           =   4785
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Karaoke Class"
      Height          =   870
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4785
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Msg1 = "Karaoke Class" & Chr(10)
    Msg1 = Msg1 & "Developed by Fausto C. Arruda " & Chr(10)
    Msg1 = Msg1 & "e-mail: arruda@sinainet.com.br" & Chr(10)
    Msg1 = Msg1 & "Londrina - Brazil" & Chr(10)
    Label1 = Msg1

    
    Msg2 = "You may freely use and modify the source code contained in this class.  You may also freely distribute any application that uses this sample code or derivations thereof.  However, you may not redistribute any part of this archive or in any way derive financial gain from this sample without the express permission of it's author "
    Label2 = Msg2

End Sub


