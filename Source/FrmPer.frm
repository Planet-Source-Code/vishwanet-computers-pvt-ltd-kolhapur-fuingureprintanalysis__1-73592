VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPerson 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Personal Details"
   ClientHeight    =   9750
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   12315
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
   Picture         =   "FrmPer.frx":0000
   ScaleHeight     =   12555
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   495
      Left            =   5400
      TabIndex        =   45
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   14737632
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59637763
      CurrentDate     =   39085
   End
   Begin VB.Frame framepass 
      BorderStyle     =   0  'None
      Caption         =   "Password"
      Height          =   2655
      Left            =   3000
      TabIndex        =   40
      Top             =   3480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Cmdcancel 
         Caption         =   "CANCLE"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtpassword 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   42
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
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
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
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
      Height          =   495
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
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
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
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
      Left            =   4680
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "Previous"
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
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
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
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   4680
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdFL 
      Caption         =   "Load All Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdFR 
      Caption         =   "Load All Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF24 
      Caption         =   "Load 4th Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF23 
      Caption         =   "Load 3rd Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF22 
      Caption         =   "Load 2nd Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF21 
      Caption         =   "Load 1st Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF20 
      Caption         =   "Load  Thumb"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF14 
      Caption         =   "Load 4th Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF13 
      Caption         =   "Load 3rd Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF12 
      Caption         =   "Load 2nd Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF11 
      Caption         =   "Load 1st Fingure"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdF10 
      Caption         =   "Load  Thumb"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtF11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   30
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtF10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   29
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtF14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   28
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtF13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   27
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtF12 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   26
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtFR 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10920
      TabIndex        =   25
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtFL 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11040
      TabIndex        =   24
      Top             =   8520
      Width           =   3975
   End
   Begin VB.TextBox txtF24 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   23
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtF23 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      TabIndex        =   22
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtF22 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   21
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtF21 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtF20 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   19
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtNo 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtAdd 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
   Begin MSComDlg.CommonDialog comDia 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   4560
      TabIndex        =   44
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Left Hand"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Right Hand"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Image F24 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Image F22 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Image F11 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Image F14 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Image F13 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Image F12 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Image F10 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Image FR 
      Enabled         =   0   'False
      Height          =   3045
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3975
   End
   Begin VB.Image FL 
      Enabled         =   0   'False
      Height          =   3045
      Left            =   11040
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Image F23 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Image F21 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Image F20 
      Enabled         =   0   'False
      Height          =   1365
      Left            =   360
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
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
      Left            =   1440
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   1440
      TabIndex        =   16
      Top             =   1560
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
      Left            =   1440
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pth As String
Dim Save As Integer

Dim RsP As New ADODB.Recordset
Dim RsP1 As New ADODB.Recordset

Public Sub BtnDisable()
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
End Sub

Public Sub BtnEnable()
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
End Sub

Public Sub GetFile(Ctrl1 As Control, Ctrl2 As Control)
    Dim i As Integer
    Dim xFnm As String

    comDia.ShowOpen
    xFnm = comDia.FileName
    If txtF10.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF11.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF12.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF13.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF14.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtFR.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF20.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF21.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF22.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF23.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtF24.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    ElseIf txtFL.Text = xFnm Then
        MsgBox "The file is allready selected", vbCritical, "Already Selected"
        Exit Sub
    End If
    
    RsP1.Close
    RsP1.Source = "Select * from Personal"
    RsP1.Open
    If RsP1.EOF = False Then
        RsP1.MoveFirst
        While RsP1.EOF = False
            For i = 3 To 14
                If RsP1.Fields(i) = UCase(comDia.FileName) Then
                    MsgBox "The file is already set", vbCritical, "Image Assigned"
                    Exit Sub
                End If
            Next i
            RsP1.MoveNext
        Wend
    End If
    Ctrl1.Text = xFnm
    Ctrl2.Picture = LoadPicture(xFnm)
End Sub

Private Sub cmdCancel_Click()
    framepass.Visible = False
End Sub

Private Sub cmdF10_Click()
    GetFile txtF10, F10
End Sub

Private Sub cmdF10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        cmdF11.SetFocus
    End If
End Sub

Private Sub cmdF11_Click()
    GetFile txtF11, F11
End Sub

Private Sub cmdF12_Click()
    GetFile txtF12, F12
End Sub

Private Sub cmdF13_Click()
    GetFile txtF13, F13
End Sub

Private Sub cmdF14_Click()
    GetFile txtF14, F14
End Sub

Private Sub cmdFR_Click()
    GetFile txtFR, FR
End Sub

Private Sub cmdF20_Click()
    GetFile txtF20, F20
End Sub
Private Sub cmdF21_Click()
    GetFile txtF21, F21
End Sub

Private Sub cmdF22_Click()
    GetFile txtF22, F22
End Sub

Private Sub cmdF23_Click()
    GetFile txtF23, F23
End Sub

Private Sub cmdF24_Click()
    GetFile txtF24, F24
End Sub

Private Sub cmdFL_Click()
    GetFile txtFL, FL
End Sub

Public Sub SetLdButEnable()
    txtName.Enabled = True
    txtAdd.Enabled = True
    dtpdate.Enabled = True
    
    cmdF10.Enabled = True
    cmdF11.Enabled = True
    cmdF12.Enabled = True
    cmdF13.Enabled = True
    cmdF14.Enabled = True
    cmdFR.Enabled = True
    
    cmdF20.Enabled = True
    cmdF21.Enabled = True
    cmdF22.Enabled = True
    cmdF23.Enabled = True
    cmdF24.Enabled = True
    cmdFL.Enabled = True
End Sub

Public Sub SetLdButDisable()
    Save = 0
    txtName.Enabled = False
    txtAdd.Enabled = False
    dtpdate.Enabled = False
    
    cmdF10.Enabled = False
    cmdF11.Enabled = False
    cmdF12.Enabled = False
    cmdF13.Enabled = False
    cmdF14.Enabled = False
    cmdFR.Enabled = False
    
    cmdF20.Enabled = False
    cmdF21.Enabled = False
    cmdF22.Enabled = False
    cmdF23.Enabled = False
    cmdF24.Enabled = False
    cmdFL.Enabled = False
End Sub

Public Sub ClearData()
    F10.Picture = LoadPicture()
    F11.Picture = LoadPicture()
    F12.Picture = LoadPicture()
    F13.Picture = LoadPicture()
    F14.Picture = LoadPicture()
    FR.Picture = LoadPicture()
    
    F20.Picture = LoadPicture()
    F21.Picture = LoadPicture()
    F22.Picture = LoadPicture()
    F23.Picture = LoadPicture()
    F24.Picture = LoadPicture()
    FL.Picture = LoadPicture()
    
    txtNo.Text = ""
    txtName.Text = ""
    txtAdd.Text = ""
    txtF10.Text = ""
    txtF11.Text = ""
    txtF12.Text = ""
    txtF13.Text = ""
    txtF14.Text = ""
    txtFR.Text = ""
        
    txtF20.Text = ""
    txtF21.Text = ""
    txtF22.Text = ""
    txtF23.Text = ""
    txtF24.Text = ""
    txtFL.Text = ""
    
    dtpdate.Value = Date
End Sub

Public Sub SetData()
    F10.Picture = LoadPicture(RsP.Fields(3))
    F11.Picture = LoadPicture(RsP.Fields(4))
    F12.Picture = LoadPicture(RsP.Fields(5))
    F13.Picture = LoadPicture(RsP.Fields(6))
    F14.Picture = LoadPicture(RsP.Fields(7))
    FR.Picture = LoadPicture(RsP.Fields(8))
    
    F20.Picture = LoadPicture(RsP.Fields(9))
    F21.Picture = LoadPicture(RsP.Fields(10))
    F22.Picture = LoadPicture(RsP.Fields(11))
    F23.Picture = LoadPicture(RsP.Fields(12))
    F24.Picture = LoadPicture(RsP.Fields(13))
    FL.Picture = LoadPicture(RsP.Fields(14))
    
    txtNo.Text = RsP.Fields(0)
    txtName.Text = RsP.Fields(1)
    txtAdd.Text = RsP.Fields(2)
    
    txtF10.Text = RsP.Fields(3)
    txtF11.Text = RsP.Fields(4)
    txtF12.Text = RsP.Fields(5)
    txtF13.Text = RsP.Fields(6)
    txtF14.Text = RsP.Fields(7)
    txtFR.Text = RsP.Fields(8)
        
    txtF20.Text = RsP.Fields(9)
    txtF21.Text = RsP.Fields(10)
    txtF22.Text = RsP.Fields(11)
    txtF23.Text = RsP.Fields(12)
    txtF24.Text = RsP.Fields(13)
    txtFL.Text = RsP.Fields(14)
    dtpdate.Value = RsP.Fields(15)
End Sub

Private Sub cmdNext_Click()
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdNew.Enabled = True
    SetLdButDisable
    If RsP.EOF = False Then
        RsP.MoveNext
        If RsP.EOF = True Then
            RsP.MoveLast
        End If
        SetData
    End If
        framepass.Visible = False

End Sub

Private Sub cmdPre_Click()
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdNew.Enabled = True
    SetLdButDisable
    If RsP.EOF = False Then
        RsP.MovePrevious
        If RsP.BOF = True Then
            RsP.MoveFirst
        End If
        SetData
    End If
    framepass.Visible = False

End Sub

Private Sub cmdLast_Click()
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdNew.Enabled = True
    SetLdButDisable
    If RsP.EOF = False Then
        RsP.MoveLast
        SetData
    End If
        framepass.Visible = False

End Sub

Private Sub cmdFirst_Click()
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdNew.Enabled = True
    SetLdButDisable
    If RsP.EOF = False Then
        RsP.MoveFirst
        SetData
    End If
    framepass.Visible = False

End Sub

Private Sub cmdNew_Click()
    Save = 1
    SetLdButEnable
    ClearData
    BtnDisable
    txtName.SetFocus
    framepass.Visible = False
End Sub

Private Sub cmdEdit_Click()
    framepass.Visible = True
    txtPassword.Text = ""
    txtPassword.SetFocus
End Sub

Private Sub cmdReturn_Click()
If RsP.State = 0 Then
    RsP.Open
End If
    
    framepass.Visible = False
    Unload objFrmPr

End Sub

Private Sub cmdSave_Click()
    If txtName.Text = "" Then
        MsgBox "You must have to fill Name ", vbCritical, "Must Enter"
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtAdd.Text = "" Then
        MsgBox "You must have to fill Address ", vbCritical, "Must Enter"
        txtAdd.SetFocus
        Exit Sub
    End If
    
    If txtF10.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF11.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF12.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF13.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF14.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtFR.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF20.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF21.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF22.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF23.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtF24.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    ElseIf txtFL.Text = "" Then
        MsgBox "You must have to select Image File", vbCritical, "Must Select"
        Exit Sub
    End If
    If RsP.State = 1 Then
        RsP.Close
    End If
    If Save = 1 Then
        RsP.Source = "insert into Personal (Name,Address,Right1,Right2,Right3,Right4,Right5," & _
                "RightAll,Left1,Left2,Left3,Left4,Left5,LeftAll,TrialDate) values ('" & _
                UCase(txtName.Text) & "','" & UCase(txtAdd.Text) & "','" & UCase(txtF10.Text) & _
                "','" & UCase(txtF11.Text) & "','" & UCase(txtF12.Text) & "','" & _
                UCase(txtF13.Text) & "','" & UCase(txtF14.Text) & "','" & UCase(txtFR.Text) & _
                "','" & UCase(txtF20.Text) & "','" & UCase(txtF21.Text) & "','" & UCase(txtF22.Text) & "','" & _
                UCase(txtF23.Text) & "','" & UCase(txtF24.Text) & "','" & UCase(txtFL.Text) & "','" & Format(dtpdate.Value, "dd/MMM/yyyy") & "' )"
    End If
    If Save = 2 Then
        RsP.Source = "Update Personal set Name = '" & UCase(txtName.Text) & "'," & _
                    " Address = '" & UCase(txtAdd.Text) & "'," & _
                    " Right1 = '" & UCase(txtF10.Text) & "'," & _
                    " Right2 = '" & UCase(txtF11.Text) & "'," & _
                    " Right3 = '" & UCase(txtF12.Text) & "'," & _
                    " Right4 = '" & UCase(txtF13.Text) & "'," & _
                    " Right5 = '" & UCase(txtF14.Text) & "'," & _
                    " RightAll = '" & UCase(txtFR.Text) & "'," & _
                    " Left1 = '" & UCase(txtF20.Text) & "'," & _
                    " Left2 = '" & UCase(txtF21.Text) & "'," & _
                    " Left3 = '" & UCase(txtF22.Text) & "'," & _
                    " Left4 = '" & UCase(txtF23.Text) & "'," & _
                    " Left5 = '" & UCase(txtF24.Text) & "'," & _
                    " LeftAll = '" & UCase(txtFL.Text) & "', " & _
                    " TrialDate = '" & Format(dtpdate.Value, "dd/MMM/yyyy") & "' " & _
                    " where ID = " & Val(txtNo.Text)
    End If
    Save = 0
    SetLdButDisable
    BtnEnable
    RsP.Open
    RsP.Source = "Select * from Personal"
    RsP.Open
    RsP1.Close
    RsP1.Source = "Select * from Personal"
    RsP1.Open
    RsP.MoveFirst
    SetData
End Sub


Private Sub Cmdok_Click()
If txtPassword.Text = "FINGER" Or txtPassword.Text = "finger" Then
    Save = 2
    SetLdButEnable
    BtnDisable
    txtName.SetFocus
    framepass.Visible = False
Else
    MsgBox "Enter Correct Password"
    txtPassword.Text = ""
    txtPassword.SetFocus
End If
End Sub

Private Sub Form_Load()

    Save = 0
    BtnEnable
    
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
        
    RsP.Source = "Select * from Personal"
    RsP.Open
    If RsP.EOF = True Then
        MsgBox "Not a single record present", vbInformation, "Information"
    Else
        RsP.MoveFirst
        SetData
    End If
    SetLdButDisable
    
    Pth = "D:\DipProj\FingurP\Images"
    comDia.InitDir = Pth
    comDia.Filter = "Image (*.jpg;*.jpeg)|*.jpg;*.jpeg"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Save = 0
    RsP1.Close
If RsP.State = 0 Then
    RsP.Open
End If
RsP.Close
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        cmdF10.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAdd.SetFocus
    End If
End Sub

Private Sub txtName_LostFocus()
If Save = 1 Then
    If txtName <> "" Then
        If RsP.State = 1 Then
            RsP.Close
        End If
        RsP1.Close
        RsP1.Source = "Select Name from Personal where Name = '" & UCase(txtName.Text) & "'"
        RsP1.Open
        If RsP1.EOF = False Then
            MsgBox "The name allready present", vbCritical, "Already Present"
            txtName.Text = ""
            txtName.SetFocus
            RsP.Source = "Select * from Personal"
            RsP.Open
            Exit Sub
        End If
    End If
End If
End Sub


Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOK.SetFocus
End If
End Sub
