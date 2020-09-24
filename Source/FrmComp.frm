VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmComp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Compare fingerprints"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmComp.frx":0000
   ScaleHeight     =   12555
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton CMDRTN 
      Caption         =   "RETURN"
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
      Left            =   4920
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txtpath 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Cmdclear 
      Caption         =   "CLEAR"
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
      Left            =   7170
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdcomp 
      Caption         =   "COMPARE"
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
      Left            =   4920
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Cmdload 
      Caption         =   "LOAD"
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
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblMnu 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Img2 
      Height          =   3000
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2505
   End
   Begin VB.Image Img1 
      Height          =   3000
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2505
   End
End
Attribute VB_Name = "FrmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pth As String
Dim Save As Integer

Dim RsP As New ADODB.Recordset
Dim RsP1 As New ADODB.Recordset

Dim pic As String

Private Sub Cmdclear_Click()
    Txtpath.Text = ""
    txtName.Text = ""
    Img1.Picture = LoadPicture()
    Img2.Picture = LoadPicture()
    cmdcomp.Enabled = False
End Sub

Private Sub cmdcomp_Click()
    Dim i As Integer
    Dim Cnt As Integer                     ' for file count
    Dim Str As String
    
    Dim F1 As String                       ' for first file name
    Dim F2 As String
    
    Dim F1Size As Double                   ' for file 1 size
    Dim F2Size As Double
    
    Dim FCnt As Double                     ' as count for byte of file
    
    Dim Flag As Integer
    
    If RsP.State = 1 Then
        RsP.Close
    End If
    
    Open Txtpath.Text For Binary Access Read As 1
    ReDim F1Arr(1 To LOF(1)) As Byte        ' define array of bites of sourcs file
    
    F1Size = FileLen(Txtpath.Text)
    
    RsP.Source = "Select * from Personal order by ID"
    RsP.Open
    
    If RsP.EOF = True Then
        MsgBox "There is Not a single Finger Print Saved", vbCritical, "Not a Single Record"
        RsP.Close
        Exit Sub
    Else
        RsP.MoveFirst
        While RsP.EOF = False
            For Cnt = 3 To 14
                Flag = 0
                Str = RsP.Fields(Cnt)
                F2Size = FileLen(Str)
                If F1Size = F2Size Then
                    FCnt = 0
                    Open RsP.Fields(Cnt) For Binary Access Read As 2
                    ReDim F2Arr(1 To LOF(2)) As Byte        ' define array of bites of target file
                    
                    Select Case Val(lblMnu.Caption)
                        Case 1
                                ' actual all bytes comparision
                                 For FCnt = 1 To LOF(2)
                                     If F1Arr(FCnt) <> F2Arr(FCnt) Then
                                         Flag = 1
                                         GoTo GotIt
                                     End If
                                 Next FCnt
                        Case 2
                                Dim CFst As Double
                                Dim CLst As Double
                                Dim RFst As Double
                                Dim RLst As Double
                                Dim FCnt1 As Integer

                                FCnt1 = 0
                                FCnt = 0
                                CFst = F2Size / 3
                                RFst = F2Size / 3
                                CLst = F2Size / 3 + CFst
                                RLst = F2Size / 3 + RFst
                                For FCnt1 = RFst To RLst Step 10
                                    For FCnt = CFst To CLst Step 10
                                        If F1Arr(FCnt) <> F2Arr(FCnt) Then
                                            Flag = 1
                                            GoTo GotIt
                                        End If
                                    Next FCnt
                                Next FCnt1
                        
                        Case 3
                                Dim Count As Single
                                Count = InputBox("Enter number of points to be compared", "Pixel Input")
                                ReDim FArr(1 To Count) As Integer
                                
                                Dim C As Integer
                                For C = 1 To Count Step 1
                                    FArr(C) = Rnd(F2Size)
                                Next C
                                
                                For C = 1 To Count Step 1
                                    If F1Arr(C) <> F2Arr(C) Then
                                        Flag = 1
                                        GoTo GotIt
                                    End If
                                 Next C
                    End Select
                    
                    Close 2
                    GoTo GotIt
                Else
                    Flag = 1
                End If
            Next Cnt
            RsP.MoveNext
'            Str = "e:\dipproj\fingerp\imgcomp.exe " & Txtpath.Text & " " & RsP.Fields(Cnt)
'            i = Shell(Str)
'            MsgBox Str & " " & i
        Wend
        Close 1
    End If
        
GotIt:
        If Flag = 1 Then
            MsgBox "There is no mathing Image"
            Img2.Picture = LoadPicture()
        End If
        If Flag = 0 Then
            Img2.Picture = LoadPicture(RsP.Fields(Cnt))
            txtName.Text = RsP.Fields(1)
            Close (1)
            Close (2)
        End If
End Sub

Private Sub Cmdload_Click()
    comDia.ShowOpen
    pic = comDia.FileName
    Txtpath.Text = pic
    Img2.Picture = LoadPicture()
    Img1.Picture = LoadPicture(pic)
    If Img1.Picture = 0 Then
        MsgBox "Load the Picture"
    Else
        cmdcomp.Enabled = True
    End If
End Sub

Private Sub CMDRTN_Click()
    Unload objFrmComp
    objFrmMdi.Show
End Sub

Private Sub Form_Load()
    
    RsP1.ActiveConnection = ModFp.Cn
    RsP1.LockType = adLockOptimistic
    RsP1.CursorLocation = adUseClient
    RsP1.CursorType = adOpenDynamic

    RsP1.Source = "Select * from Personal"
    RsP1.Open

    RsP.ActiveConnection = ModFp.Cn
    RsP.LockType = adLockOptimistic
    RsP.CursorLocation = adUseClient
    RsP.CursorType = adOpenDynamic

    cmdcomp.Enabled = False
    Txtpath.Enabled = False
    Pth = "E:\DipProj\FingurP\Images"
    comDia.InitDir = Pth
    comDia.Filter = "Image (*.jpg;*.jpeg)|*.jpg;*.jpeg"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RsP1.State = 1 Then
        RsP1.Close
    End If
    If RsP.State = 1 Then
        RsP.Close
    End If
End Sub
