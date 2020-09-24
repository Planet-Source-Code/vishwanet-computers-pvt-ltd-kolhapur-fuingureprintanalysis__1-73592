VERSION 5.00
Begin VB.Form frmpassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   2610
   ClientLeft      =   2190
   ClientTop       =   4860
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   9195
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Password"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdexit_Click()
    End
End Sub

Private Sub Cmdok_Click()
'    If txtpass.Text = "finger" Or txtpass.Text = "FINGER" Then
        MDIFp.Show
        frmpassword.Hide
'    ElseIf txtpass.Text = "" Then
'        MsgBox "Enter the Password"
'        txtpass.SetFocus
'    Else
'        MsgBox "Enter Correct Password"
'        txtpass.Text = ""
'        txtpass.SetFocus
'    End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdok.SetFocus
    End If
End Sub
