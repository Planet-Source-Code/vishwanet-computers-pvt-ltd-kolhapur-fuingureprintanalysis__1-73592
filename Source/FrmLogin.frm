VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2415
   ClientLeft      =   6525
   ClientTop       =   5820
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmLogin.frx":0000
   ScaleHeight     =   1426.861
   ScaleMode       =   0  'User
   ScaleWidth      =   6520.978
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   5280
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2925
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      URL             =   "D:\SANDY\AA054591.gif"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6165
      _cy             =   4048
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnt As Integer

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    End
End Sub

Private Sub Cmdok_Click()
    If txtPassword = "finger" Then
        MsgBox "Congrats ", , "Valid User"
        LoginSucceeded = True
        Me.Hide
    Else
        Cnt = Cnt + 1
        If Cnt < 4 Then
            If Cnt = 3 Then
                MsgBox "Your last try", vbCritical, "Login"
            Else
                MsgBox "Invalid Password, try again!", , "Login"
            End If
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        Else
            End
        End If
    End If
End Sub

Private Sub Form_Load()
    LoginSucceeded = False
    Cnt = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If LoginSucceeded = False Then
        End
    End If
End Sub

