VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   3600
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   1200
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsP As New ADODB.Recordset
Private Sub Form_Load()
    RsP.ActiveConnection = ModFp.Cn
    RsP.LockType = adLockOptimistic
    RsP.CursorLocation = adUseClient
    RsP.CursorType = adOpenDynamic

    RsP.Source = "Select * from Personal"
    RsP.Open
    Text1.Text = RsP.Fields(0)
    Text2.Text = RsP.Fields(1)
    Image1.Picture = LoadPicture(RsP.Fields(4))
End Sub
