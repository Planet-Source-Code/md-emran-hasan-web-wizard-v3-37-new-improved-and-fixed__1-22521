VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4380
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmProgress.frx":0000
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   960
   End
   Begin VB.PictureBox picStatus 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Generating Web Page..."
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   3675
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
Call DoFinal
Timer1.Enabled = True
End Sub

Private Sub Status(StartPos As Integer, EndPos As Integer, Caption As String)
Label1.Caption = Caption
For i = StartPos To EndPos Step 0.1
    ProgressBar1.Value = i
Next i
End Sub

Private Sub DoWork()

Call Status(1, 25, "Generating Web Page...")
frmMain.rtbText.SaveFile "C:\tmp.001"
Call Status(25, 50, "Setting up page header...")
Call LoadText(frmMain.txtTemp, "C:\tmp.001")
Call Status(50, 75, "Setting up page body...")
frmMain.txtFinal.Text = rtf2HTML.rtf2HTML(frmMain.txtTemp.Text)
frmMain.txtFinal.Text = frmMain.txtStarting.Text & frmMain.txtFinal.Text & Me.Text1.Text
If frmMain.Preview = False Then
Call Status(75, 100, "Saving HTML document...")
Call FileSave(frmMain.txtFinal.Text, frmMain.txtLocation.Text)
Else
Call Status(75, 100, "Making document ready for preview...")
Call FileSave(frmMain.txtFinal.Text, "C:\tempDoc.htm")
End If
End Sub

Private Sub Timer1_Timer()
Call DoWork
Timer1.Enabled = False
If frmMain.Preview = False Then
Unload Me
Else
Call ShellExecute(&O0, vbNullString, "C:\tempDoc.htm", vbNullString, vbNullString, vbNormalFocus)
Unload Me
End If
End Sub

Private Sub DoFinal()
Text1.Text = Text1.Text & "<br><br><b>"
Text1.Text = Text1.Text & "<center><font face=Verdana size=2>"

If Len(frmMain.txtCompanyName.Text) > 0 Then
Text1.Text = Text1.Text & "Copyright © " & frmMain.txtCompanyName.Text & ". All rights reserved.<br>"
End If

If frmMain.chkContact.Value <> 1 Then
Text1.Text = Text1.Text & frmMain.txtAddress.Text

If Len(frmMain.txtPhone.Text) > 0 Then
    Text1.Text = Text1.Text & "Phone :" & frmMain.txtPhone.Text
End If

If Len(frmMain.txtFax.Text) > 0 Then
    Text1.Text = Text1.Text & ", Fax :" & frmMain.txtFax.Text
End If
End If

If Len(frmMain.txtEmail.Text) > 0 Then
Text1.Text = Text1.Text & "<a href=""mailto:" & frmMain.txtEmail.Text & """>" & frmMain.txtEmail.Text
End If

Text1.Text = Text1.Text & "</font></center><b></body></html>"
End Sub
