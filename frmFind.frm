VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   200
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Find What:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
Dim textFound As Integer
cmdFindNext.Enabled = True
frmMain.rtbText.Find (Text1.Text)
frmMain.rtbText.SetFocus

textFound = frmMain.rtbText.Find(Text1.Text)
If textFound = -1 Then
MsgBox vbCr & "Text could not be found.", vbInformation, "Find"
Unload Me
End If
End Sub

Private Sub cmdFindNext_Click()
frmMain.rtbText.SetFocus
  
frmMain.rtbText.Find (Text1.Text), frmMain.rtbText.SelStart + 1

End Sub
