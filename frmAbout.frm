VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Credits"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   1440
      ScaleHeight     =   1275
      ScaleWidth      =   3315
      TabIndex        =   23
      Top             =   1600
      Width           =   3375
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "S/N: 21EC2020-3AEA-1069-A2DD"
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   840
         Width           =   2430
      End
      Begin VB.Label regCompany 
         AutoSize        =   -1  'True
         Caption         =   "Who are you ?"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label regName 
         AutoSize        =   -1  'True
         Caption         =   "Yahoo!"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2850
      Left            =   120
      ScaleHeight     =   2790
      ScaleWidth      =   1035
      TabIndex        =   22
      Top             =   120
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAbout.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1440
      ScaleHeight     =   1335
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   1600
      Width           =   3375
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ehasan@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1780
         MouseIcon       =   "frmAbout.frx":27AE
         MousePointer    =   99  'Custom
         TabIndex        =   34
         ToolTipText     =   "Mail to the author"
         Top             =   1000
         Width           =   1470
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "admin@ehasan.net,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   320
         MouseIcon       =   "frmAbout.frx":2AB8
         MousePointer    =   99  'Custom
         TabIndex        =   33
         ToolTipText     =   "Mail to the ERA Developers Group."
         Top             =   1005
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3600
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   -1680
         X2              =   3360
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   18
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1830
         TabIndex        =   17
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1725
         TabIndex        =   16
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1590
         TabIndex        =   15
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1515
         TabIndex        =   14
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   13
         Top             =   120
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   2520
         TabIndex        =   10
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   2400
         TabIndex        =   9
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   2205
         TabIndex        =   7
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2085
         TabIndex        =   6
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Complex,Shahbag,Dhaka,Bangladesh."
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2760
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Flat #3-A Aziz Co-Operative Housing"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   400
         Width           =   2655
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   3960
      Top             =   4560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "http://www16.brinkster.com/eragroup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1440
      MouseIcon       =   "frmAbout.frx":2DC2
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "Visit ERA Groups home page on the net"
      Top             =   975
      Width           =   2700
   End
   Begin VB.Label Label12 
      Caption         =   $"frmAbout.frx":30CC
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "This product is licensed to:"
      Height          =   195
      Left            =   1440
      TabIndex        =   24
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Make a web page in minutes."
      Height          =   195
      Left            =   1440
      TabIndex        =   21
      Top             =   360
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Web Wizard v3.26"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Copyright ©  2000-01, ERA Developers Group."
      Height          =   195
      Left            =   1440
      TabIndex        =   19
      Top             =   650
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim n, m

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Label4.Caption = "This product is developed by..."
Picture3.Visible = False
End Sub

Private Sub Form_Load()
n = 0: m = 0
regName.Caption = modOthers.RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
regCompany.Caption = modOthers.RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
End Sub

Private Sub Label10_Click()
    Call ShellExecute(&O0, vbNullString, "mailto:ehasan@yahoo.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label8_Click()
    Call ShellExecute(&O0, vbNullString, "http://www16.brinkster.com/eragroup", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label9_Click()
    Call ShellExecute(&O0, vbNullString, "mailto:admin@ehasan.net", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Timer1_Timer()
Label5(n).ForeColor = vbRed
If n > 0 And n < 15 Then
Label5(n - 1).ForeColor = vbBlack
ElseIf n = 15 Then
Label5(n).ForeColor = vbBlack
Label5(n - 1).ForeColor = vbBlack
End If
If n = 15 Then
n = 0
Else
n = n + 1
End If
End Sub

