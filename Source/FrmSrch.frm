VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSrch 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Search a Person"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmSrch.frx":0000
   ScaleHeight     =   12555
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdclr 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdRd 
      Caption         =   "Reduce"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdEn 
      Caption         =   "Enlarge"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Image"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6480
      Width           =   2175
   End
   Begin VB.ListBox lstFing 
      BackColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   3360
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.ListBox lstHand 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DgSrch 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   14737632
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   3
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
   Begin VB.Image imgF 
      Height          =   3000
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmSrch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pth As String
Dim Str As String

Dim RsSrch As New ADODB.Recordset
Dim RsSh As New ADODB.Recordset

Private Sub cmdclr_Click()
    txtName.Text = ""
    txtName.SetFocus
    imgF.Picture = LoadPicture()
End Sub

Private Sub cmdEn_Click()
    If Pth = "" Then
        MsgBox "Please select show button", vbCritical, "Must Select"
        txtName.SetFocus
        Exit Sub
    Else
        If imgF.Width < 8000 And imgF.Height < 9000 Then
            imgF.Width = imgF.Width + 100
            imgF.Height = imgF.Height + 100
        End If
    End If
End Sub

Private Sub cmdRd_Click()
    If Pth = "" Then
        MsgBox "Please select show button", vbCritical, "Must Select"
        txtName.SetFocus
        Exit Sub
    Else
        If imgF.Width > 2500 And imgF.Height > 3000 Then
            imgF.Width = imgF.Width - 100
            imgF.Height = imgF.Height - 100
        End If
    End If
End Sub


Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    If imgF.Height > 3000 And imgF.Width > 2500 Then
        imgF.Height = 3000
        imgF.Width = 2500
        End If
    Lst1 = lstHand.ListIndex
    Lst2 = lstFing.ListIndex
        
    Str = ""
    If Lst1 = 0 Then
        Str = "Right"
    End If
    
    If Lst1 = 1 Then
        Str = "Left"
    End If
    
    Select Case Lst2
        Case 0
            Str = Str & "1"
        Case 1
            Str = Str & "2"
        Case 2
            Str = Str & "3"
        Case 3
            Str = Str & "4"
        Case 4
            Str = Str & "5"
        Case 5
            Str = Str & "All"
    End Select
    
    RsSh.Source = "Select " & Str & " from Personal where ID = " & DgSrch.Columns(0)
    RsSh.Open
    Pth = RsSh.Fields(0)
    RsSh.Close
    imgF.Picture = LoadPicture(Pth)
    cmdEn.SetFocus
End Sub

Private Sub DgSrch_Click()
    lstHand.SetFocus
End Sub

Private Sub DgSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         lstHand.SetFocus
        lstHand.Text = lstHand.List(0)
        txtName.Text = DgSrch.Columns(1)
        DgSrch_LostFocus
    End If
End Sub

Private Sub DgSrch_LostFocus()
    lstHand.SetFocus
End Sub

Private Sub Form_Load()
    RsSrch.ActiveConnection = ModFp.Cn
    RsSrch.LockType = adLockOptimistic
    RsSrch.CursorLocation = adUseClient
    RsSrch.CursorType = adOpenDynamic
        
    RsSh.ActiveConnection = ModFp.Cn
    RsSh.LockType = adLockOptimistic
    RsSh.CursorLocation = adUseClient
    RsSh.CursorType = adOpenDynamic
        
    RsSrch.Source = "Select ID,Name,Address from Personal"
    RsSrch.Open
    If RsSrch.EOF = True Then
        MsgBox "Not a single record present", vbInformation, "Information"
    Else
        Set DgSrch.DataSource = RsSrch
    End If
    
    Pth = ""
    lstHand.AddItem "Right"
    lstHand.AddItem "Left"
    
    lstFing.AddItem "Thumb"
    lstFing.AddItem "First"
    lstFing.AddItem "Second"
    lstFing.AddItem "Third"
    lstFing.AddItem "Fourth"
    lstFing.AddItem "All"
    
    txtName.Text = DgSrch.Columns(1)
    lstHand.Text = lstHand.List(0)
    lstFing.Text = lstFing.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Save = 0
    RsSrch.Close
End Sub

Private Sub lstFing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdShow.SetFocus
    End If
End Sub

Private Sub lstHand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstFing.SetFocus
        lstFing.Text = lstFing.List(0)
    End If
End Sub

Private Sub txtName_Change()
    Dim L As Integer
    L = Len(txtName.Text)
    
    RsSrch.Close
    RsSrch.Source = "Select Id,Name,Address from Personal where left(Name," & _
                    L & ") = '" & UCase(txtName.Text) & "'"
    RsSrch.Open
    Set DgSrch.DataSource = RsSrch
End Sub

Private Sub txtName_GotFocus()
    txtName.Text = ""
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DgSrch.SetFocus
    End If
End Sub
