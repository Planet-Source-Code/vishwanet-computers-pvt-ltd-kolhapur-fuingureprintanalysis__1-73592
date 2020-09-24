VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm FrmMdiFp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Finger Print Analysis"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMdiFp.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=FP"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=FP"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuPer 
      Caption         =   "&Personal Detail"
      Begin VB.Menu mnuNewPer 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuComp 
         Caption         =   "&Compare"
         Begin VB.Menu mnuCompFile 
            Caption         =   "Complete File"
         End
         Begin VB.Menu mnuCompCent 
            Caption         =   "Center Protion"
         End
         Begin VB.Menu mnuCompRnd 
            Caption         =   "Random Bytes"
         End
      End
   End
   Begin VB.Menu mnuSrch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSrchPer 
         Caption         =   "&Person"
      End
      Begin VB.Menu dtls 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Information"
      End
   End
   Begin VB.Menu mnuFpInfo 
      Caption         =   "&Finger Print"
      Begin VB.Menu mnuFpType 
         Caption         =   "Finger Print &Types"
      End
      Begin VB.Menu mnuFpPro 
         Caption         =   "Finger Print &Properties"
      End
   End
   Begin VB.Menu mnuExt 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMdiFp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HideAll()
    objFrmComp.Hide
    objFrmPr.Hide
    objFrmSrch.Hide
    objFrmPro.Hide
    objFrmType.Hide
End Sub

Private Sub dtls_Click()
    If LoginSucceeded = True Then
        HideAll
        rptdtls.Show
    End If
End Sub

Private Sub extra_Click()
    If LoginSucceeded = True Then
        HideAll
        Form1.Show
    End If
End Sub

Private Sub MDIForm_Load()
    objFrmlogin.Show
End Sub

Private Sub mnuCompCent_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmComp.Show
        objFrmComp.lblMnu.Caption = "2"
    End If
End Sub

Private Sub mnuCompFile_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmComp.Show
        objFrmComp.lblMnu.Caption = "1"
    End If
End Sub

Private Sub mnuCompRnd_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmComp.Show
        objFrmComp.lblMnu.Caption = "3"
    End If
End Sub

Private Sub mnuExt_Click()
    End
End Sub

Private Sub mnuFpPro_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmPro.Show
    End If
End Sub

Private Sub mnuFpType_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmType.Show
    End If
End Sub

Private Sub mnuInfo_Click()
    If LoginSucceeded = True Then
        HideAll
        RptPertInfo.Show
    End If
End Sub

Private Sub mnuNewPer_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmPr.Show
    End If
End Sub

Private Sub mnuSrchPer_Click()
    If LoginSucceeded = True Then
        HideAll
        objFrmSrch.Show
    End If
End Sub
