VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Wizard - Start"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   7
      Left            =   2400
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   85
      Top             =   150
      Width           =   4815
      Begin VB.CommandButton Command8 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   1680
         TabIndex        =   89
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label36 
         Caption         =   $"frmStart.frx":27A2
         Height          =   855
         Left            =   120
         TabIndex        =   88
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label35 
         Caption         =   $"frmStart.frx":285F
         Height          =   975
         Left            =   120
         TabIndex        =   87
         Top             =   600
         Width           =   4550
      End
      Begin VB.Label Label34 
         Caption         =   "Finished !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   2400
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   9
      Top             =   150
      Width           =   4815
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "http://www16.brinkster.com/eragroup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmStart.frx":2922
         MousePointer    =   99  'Custom
         TabIndex        =   102
         ToolTipText     =   "Visit ERA Groups home page on the net"
         Top             =   3480
         Width           =   3270
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Copyright ©  2000-01, ERA Developers Group."
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
         Left            =   120
         TabIndex        =   101
         Top             =   3150
         Width           =   3840
      End
      Begin VB.Label Label7 
         Caption         =   "Click on Next to continue..."
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Follow the on screen instruction to set up your web page."
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   $"frmStart.frx":2C2C
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   $"frmStart.frx":2CDF
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Welcome to Web Wizard !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   6
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   73
      Top             =   150
      Width           =   4815
      Begin VB.CheckBox chkContact 
         Caption         =   "&Don't include contact details."
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   3520
         Width           =   2655
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   120
         TabIndex        =   75
         Top             =   680
         Width           =   4335
      End
      Begin VB.TextBox txtAddress 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   1360
         Width           =   4335
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   120
         TabIndex        =   77
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   2400
         TabIndex        =   78
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   120
         TabIndex        =   79
         Top             =   3080
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "&Company Name"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "&Street Address"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "&Telephone"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Fa&ximile"
         Height          =   255
         Left            =   2400
         TabIndex        =   82
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "&E-mail"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2840
         Width           =   1215
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Specify the contact information for your web page:"
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
         Left            =   120
         TabIndex        =   74
         Top             =   75
         Width           =   4290
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   5
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   70
      Top             =   200
      Width           =   4815
      Begin RichTextLib.RichTextBox rtbText 
         Height          =   3375
         Left            =   0
         TabIndex        =   71
         Top             =   405
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5953
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmStart.frx":2D78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Left"
               Object.ToolTipText     =   "Align Left"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Center"
               Object.ToolTipText     =   "Align Center"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Align Right"
               Object.ToolTipText     =   "Align Right"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Font"
               Object.ToolTipText     =   "Font"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Text Color"
               Object.ToolTipText     =   "Text Color"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullet"
               Object.ToolTipText     =   "Bullet"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Date/Time"
               Object.ToolTipText     =   "Date/Time"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert File"
               Object.ToolTipText     =   "Insert File"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   4
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   60
      Top             =   150
      Width           =   4815
      Begin MSComctlLib.ListView lv 
         Height          =   1215
         Left            =   120
         TabIndex        =   66
         Top             =   2280
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Event"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Transition Effect"
            Object.Width           =   3511
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Duration"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.ListBox lstEffects 
         Height          =   1425
         Left            =   2400
         TabIndex        =   65
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1620
         TabIndex        =   64
         Text            =   "1.0"
         Top             =   1530
         Width           =   495
      End
      Begin VB.ComboBox cboEvent 
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Text            =   "Combo1"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "N.B. Applicable only in Internet Explorer 5 or above."
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
         Left            =   120
         TabIndex        =   69
         Top             =   3600
         Width           =   4320
      End
      Begin VB.Label Label27 
         Caption         =   "&Transation Effects:"
         Height          =   255
         Left            =   2400
         TabIndex        =   68
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "&Durations(seconds):"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "&Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Choose the transition effect for your web page:"
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
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   3
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   28
      Top             =   150
      Width           =   4815
      Begin VB.TextBox txtMetaAuthor 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtMetaDescription 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   3120
         Width           =   4455
      End
      Begin VB.TextBox txtMetaKeyword 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtMetaCopyright 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "&Author"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Description"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Keyword"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Specify the following META tags for your web page:"
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
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   4290
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "&Copyright © "
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   2400
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   20
      Top             =   150
      Width           =   4815
      Begin VB.CommandButton Command6 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   4455
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label11 
         Caption         =   "Web page &location"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Where do you want save your web page ?"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   3510
      End
      Begin VB.Label Label9 
         Caption         =   "Web page &title"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "What is the title of your web page ?"
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
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2985
      End
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   2
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   38
      Top             =   150
      Width           =   4815
      Begin VB.CheckBox chkWatermark 
         Caption         =   "Watermark (fixed) - Internet Explorer Only"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Width           =   3855
      End
      Begin VB.TextBox txtMarginHeight 
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtMarginWidth 
         Height          =   285
         Left            =   3360
         TabIndex        =   50
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtTopMargin 
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtLeftMargin 
         Height          =   285
         Left            =   1080
         TabIndex        =   48
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtVLinkColor 
         Height          =   285
         Left            =   3360
         TabIndex        =   47
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtLinkColor 
         Height          =   285
         Left            =   3360
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtTextColor 
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         Text            =   "#66CCFF"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtBGColor 
         Height          =   285
         Left            =   1080
         TabIndex        =   44
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdBrowseImg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "B&rowse"
         Height          =   350
         Left            =   3660
         TabIndex        =   55
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtBgImage 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblVLink 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   93
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblALink 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   92
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   91
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblBG 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   90
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Background &Image:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label vf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin &Height:"
         Height          =   195
         Left            =   2235
         TabIndex        =   58
         Top             =   2295
         Width           =   1050
      End
      Begin VB.Label ff 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin &Width:"
         Height          =   195
         Left            =   2280
         TabIndex        =   57
         Top             =   1695
         Width           =   1005
      End
      Begin VB.Label Labelfdf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To&p Margin:"
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Left Margin:"
         Height          =   195
         Left            =   135
         TabIndex        =   52
         Top             =   1695
         Width           =   870
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Visited Link: "
         Height          =   195
         Left            =   2475
         TabIndex        =   43
         Top             =   1095
         Width           =   885
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Active Link: "
         Height          =   195
         Left            =   2475
         TabIndex        =   42
         Top             =   495
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Text : "
         Height          =   195
         Left            =   480
         TabIndex        =   41
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back&ground:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Specify the Body Elements for your web page:"
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
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.TextBox txtTrans2 
      Height          =   285
      Left            =   9000
      TabIndex        =   99
      Text            =   "<meta http-equiv=""Effect"" content=""revealTrans(Duration=Emran,Transition=Hasan)"">"
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtTrans 
      Height          =   285
      Left            =   8000
      TabIndex        =   98
      Text            =   "Text1"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtStarting 
      Height          =   1455
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   97
      Text            =   "frmStart.frx":2E4E
      Top             =   8000
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1560
      Top             =   4080
   End
   Begin VB.TextBox txtFinal 
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   96
      Top             =   9000
      Width           =   2415
   End
   Begin VB.TextBox txtTemp 
      Height          =   1575
      Left            =   9000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   95
      Text            =   "frmStart.frx":2E54
      Top             =   2040
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox rtbTemp 
      Height          =   735
      Left            =   2880
      TabIndex        =   94
      Top             =   5000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmStart.frx":2E5A
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3480
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12345
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":2F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3098
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":31F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":34AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3608
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3764
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":38C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":3F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":4094
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":4B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":59B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   350
      Left            =   120
      TabIndex        =   14
      Top             =   4120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   2520
      TabIndex        =   13
      Top             =   4120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3720
      TabIndex        =   12
      Top             =   4120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   4120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFinished 
      Caption         =   "&Finish"
      Height          =   350
      Left            =   6120
      TabIndex        =   10
      Top             =   4120
      Width           =   1095
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   8
         Top             =   2880
         Width           =   915
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   360
         Top             =   2880
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   480
         X2              =   360
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   120
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   615
         TabIndex        =   7
         Top             =   3240
         Width           =   285
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page Content"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page Transition"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   5
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "META Tag Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body Elements"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title && Location"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   360
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   360
         Top             =   1920
         Width           =   255
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   360
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   360
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape shpStart 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   360
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   480
         X2              =   480
         Y1              =   240
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   480
         X2              =   360
         Y1              =   240
         Y2              =   240
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Preview As Boolean
Private Sub cmdAbout_Click()
frmAbout.Show 1
End Sub

Private Sub cmdBack_Click()
'Check which label is currently selected
For i = 0 To 7
If Label1(i).FontBold = True Then bcnt = i - 1
Next i
'Handle all possible errors and show the content
If bcnt = 0 Then
ShowContent (bcnt)
ChangeColor (bcnt)
If Label1(bcnt).Caption <> "Title && Location" Then
Me.Caption = "Web Wizard - " + Label1(bcnt).Caption
Else
Me.Caption = "Web Wizard - " + "Title & Location"
End If
cmdBack.Enabled = False
cmdNext.Enabled = True
Else
ShowContent (bcnt)
ChangeColor (bcnt)
If Label1(bcnt).Caption <> "Title && Location" Then
Me.Caption = "Web Wizard - " + Label1(bcnt).Caption
Else
Me.Caption = "Web Wizard - " + "Title & Location"
End If

cmdNext.Enabled = True
End If
End Sub

Private Sub cmdBrowseImg_Click()
CD1.DialogTitle = "Choose the background image..."
CD1.Filter = "GIF Image(*.gif)|*.gif|JPEG Image(*.jpg)|*.jpg|Bitmap Image(*.bmp)|*.bmp|All Files (*.*)|*.*|"
CD1.InitDir = App.Path
CD1.ShowOpen
If Len(CD1.FileName) > 0 Then
txtBgImage.Text = CD1.FileName
Else
txtBgImage.Text = ""
End If
End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdFinished_Click()
Preview = False
frmProgress.Show 1
b = MsgBox("Do you want to create another web page ?", vbYesNo + vbQuestion, "Web Wizard")
If b = vbNo Then
    End
Else
    Call ShellExecute(&O0, vbNullString, App.Path & "\" & App.EXEName, vbNullString, vbNullString, vbNormalFocus)
    End
End If
End Sub

Private Sub cmdNext_Click()
'Check which label is currently selected
For i = 0 To 7
If Label1(i).FontBold = True Then cnt = i + 1
Next i
'Handle all possible errors and show the content
If cnt = 7 Then
ShowContent (cnt)
ChangeColor (cnt)
If Label1(cnt).Caption <> "Title && Location" Then
Me.Caption = "Web Wizard - " + Label1(cnt).Caption
Else
Me.Caption = "Web Wizard - " + "Title & Location"
End If
cmdNext.Enabled = False
Else
ShowContent (cnt)
ChangeColor (cnt)
If Label1(cnt).Caption <> "Title && Location" Then
Me.Caption = "Web Wizard - " + Label1(cnt).Caption
Else
Me.Caption = "Web Wizard - " + "Title & Location"
End If
End If
cmdBack.Enabled = True
End Sub

Private Sub Command6_Click()
CD1.DialogTitle = "Choose the filename for your web page..."
CD1.Filter = "HTML (*.html)|*.html|All Files (*.*)|*.*|"
CD1.InitDir = App.Path
CD1.ShowSave
If Len(CD1.FileName) > 0 Then
txtLocation.Text = CD1.FileName
Else
txtLocation.Text = App.Path & "\" & "NewPage1.htm"
End If
End Sub

Private Sub Command8_Click()
Preview = True
frmProgress.Show 1
End Sub

Private Sub Form_Load()
Label1_Click (0)
Call SetCombo
Call SetTextBox
Call SetListBox
End Sub

Private Sub Label1_Click(Index As Integer)
ShowContent (Index)
ChangeColor (Index)
If Label1(Index).Caption <> "Title && Location" Then
Me.Caption = "Web Wizard - " + Label1(Index).Caption
Else
Me.Caption = "Web Wizard - " + "Title & Location"
End If
If Index = 0 Then
cmdBack.Enabled = False
cmdNext.Enabled = True
ElseIf Index = 7 Then
cmdNext.Enabled = False
cmdBack.Enabled = True
Else
cmdNext.Enabled = True
cmdBack.Enabled = True
End If
End Sub

Private Sub Label38_Click()
    Call ShellExecute(&O0, vbNullString, "http://www16.brinkster.com/eragroup", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblALink_Click()
CD1.ShowColor
txtLinkColor.Text = "#" + RGBHexColor(CD1.Color)
lblALink.BackColor = CD1.Color
End Sub

Private Sub lblBG_Click()
CD1.ShowColor
txtBGColor.Text = "#" + RGBHexColor(CD1.Color)
lblBG.BackColor = CD1.Color
End Sub

Private Sub lblText_Click()
CD1.ShowColor
txtTextColor.Text = "#" + RGBHexColor(CD1.Color)
lblText.BackColor = CD1.Color
End Sub

Private Sub lblVLink_Click()
CD1.ShowColor
txtVLinkColor.Text = "#" + RGBHexColor(CD1.Color)
lblVLink.BackColor = CD1.Color
End Sub

Private Sub lstEffects_Click()
Dim EventList(4)
EventList(1) = "Page Enter"
EventList(2) = "Page Exit"
EventList(3) = "Site Enter"
EventList(4) = "Site Exit"
Select Case cboEvent.ListIndex

Case 0
lv.ListItems.Remove (1)
Set q = lv.ListItems.Add(1, , EventList(1))

q.SubItems(1) = lstEffects.List(lstEffects.ListIndex)
q.SubItems(2) = txtDuration.Text

Case 1
lv.ListItems.Remove (2)
Set q = lv.ListItems.Add(2, , EventList(2))

q.SubItems(1) = lstEffects.List(lstEffects.ListIndex)
q.SubItems(2) = txtDuration.Text

Case 2
lv.ListItems.Remove (3)
Set q = lv.ListItems.Add(3, , EventList(3))

q.SubItems(1) = lstEffects.List(lstEffects.ListIndex)
q.SubItems(2) = txtDuration.Text

Case 3
lv.ListItems.Remove (4)
Set q = lv.ListItems.Add(4, , EventList(4))

q.SubItems(1) = lstEffects.List(lstEffects.ListIndex)
q.SubItems(2) = txtDuration.Text

End Select
End Sub

Private Sub lv_Click()
Set q = lv.ListItems.Item(lv.SelectedItem.Index)
Select Case q.Text

Case "Page Enter"
cboEvent.ListIndex = 0

Case "Page Exit"
cboEvent.ListIndex = 1

Case "Site Enter"
cboEvent.ListIndex = 2

Case "Site Exit"
cboEvent.ListIndex = 3

End Select
End Sub

Private Sub Timer1_Timer()
txtStarting.Text = "<HTML><HEAD>" & "<TITLE>" & txtTitle.Text & "</TITLE>"

If Len(txtMetaCopyright.Text) > 0 Then
txtStarting.Text = txtStarting.Text & "<META content = """ & txtMetaCopyright.Text & """ name=Copyright>"
End If

If Len(txtMetaAuthor.Text) > 0 Then
txtStarting.Text = txtStarting.Text & "<META content = """ & txtMetaAuthor.Text & """ name=Author>"
End If

txtStarting.Text = txtStarting.Text & "<META content = ""Web Wizard 3"" name=Generator>"

If Len(txtMetaDescription.Text) > 0 Then
txtStarting.Text = txtStarting.Text & "<META content = """ & txtMetaDescription.Text & """ name=Description>"
End If

If Len(txtMetaKeyword.Text) > 0 Then
txtStarting.Text = txtStarting.Text & "<META content = """ & txtMetaKeyword.Text & """ name=keywords>"
End If

sd$ = TransWhich

txtTrans.Text = txtTrans2.Text
If InStr(1, sd$, "1") > 0 Then
Set q = lv.ListItems.Item(1)
txtTrans.Text = Replace(txtTrans.Text, "Effect", "Page-Enter")
txtTrans.Text = Replace(txtTrans.Text, "Emran", q.SubItems(2))
txtTrans.Text = Replace(txtTrans.Text, "Hasan", GetEffectNum(q.SubItems(1)))
txtStarting.Text = txtStarting.Text & txtTrans.Text
End If

txtTrans.Text = txtTrans2.Text
If InStr(1, sd$, "2") > 0 Then
Set q = lv.ListItems.Item(2)
txtTrans.Text = Replace(txtTrans.Text, "Effect", "Page-Exit")
txtTrans.Text = Replace(txtTrans.Text, "Emran", q.SubItems(2))
txtTrans.Text = Replace(txtTrans.Text, "Hasan", GetEffectNum(q.SubItems(1)))
txtStarting.Text = txtStarting.Text & txtTrans.Text
End If

txtTrans.Text = txtTrans2.Text
If InStr(1, sd$, "3") > 0 Then
Set q = lv.ListItems.Item(3)
txtTrans.Text = Replace(txtTrans.Text, "Effect", "Site-Enter")
txtTrans.Text = Replace(txtTrans.Text, "Emran", q.SubItems(2))
txtTrans.Text = Replace(txtTrans.Text, "Hasan", GetEffectNum(q.SubItems(1)))
txtStarting.Text = txtStarting.Text & txtTrans.Text
End If

txtTrans.Text = txtTrans2.Text
If InStr(1, sd$, "4") > 0 Then
Set q = lv.ListItems.Item(4)
txtTrans.Text = Replace(txtTrans.Text, "Effect", "Site-Exit")
txtTrans.Text = Replace(txtTrans.Text, "Emran", q.SubItems(2))
txtTrans.Text = Replace(txtTrans.Text, "Hasan", GetEffectNum(q.SubItems(1)))
txtStarting.Text = txtStarting.Text & txtTrans.Text
End If

txtStarting.Text = txtStarting.Text & "</HEAD>"

If Len(txtBgImage.Text) > 0 Then
txtStarting.Text = txtStarting.Text & "<BODY background=""" & txtBgImage.Text & """ text=" & txtTextColor.Text & " link=" & txtLinkColor.Text & " vlink=" & txtVLinkColor.Text
Else
txtStarting.Text = txtStarting.Text & "<BODY bgcolor=" & txtBGColor.Text & " text=" & txtTextColor.Text & " link=" & txtLinkColor.Text & " vlink=" & txtVLinkColor.Text
End If

If Len(txtLeftMargin.Text) > 0 Then
txtStarting.Text = txtStarting.Text & " leftmargin=""" & txtLeftMargin.Text & """"
End If

If Len(txtTopMargin.Text) > 0 Then
txtStarting.Text = txtStarting.Text & " topmargin=""" & txtTopMargin.Text & """"
End If

If Len(txtMarginHeight.Text) > 0 Then
txtStarting.Text = txtStarting.Text & " marginheight=""" & txtMarginHeight.Text & """"
End If

If Len(txtMarginWidth.Text) > 0 Then
txtStarting.Text = txtStarting.Text & " marginwidth=""" & txtMarginWidth.Text & """"
End If

If chkWatermark.Value = 1 Then
txtStarting.Text = txtStarting.Text & " bgproperties = fixed"
End If

txtStarting.Text = txtStarting.Text & ">"

End Sub

Private Function GetEffectNum(EffectName As String) As String
Select Case EffectName

 Case "None"
 GetEffectNum = None
 
 Case "Box In"
 GetEffectNum = 0
 
 Case "Box Out"
 GetEffectNum = 1
 
 Case "Circle In"
 GetEffectNum = 2
 
 Case "Circle Out"
 GetEffectNum = 3
 
 Case "Wipe Up"
 GetEffectNum = 4
 
 Case "Wipe Down"
 GetEffectNum = 5
 
 Case "Wipe Right"
 GetEffectNum = 6
 
 Case "Wipe Left"
 GetEffectNum = 7
 
 Case "Vertical Blinds"
 GetEffectNum = 8
 
 Case "Horizontal Blinds"
 GetEffectNum = 9
 
 Case "Checkerboard Across"
 GetEffectNum = 10
 
 Case "Checkerboard Down"
 GetEffectNum = 11
 
 Case "Random Dissolve"
 GetEffectNum = 12
 
 Case "Split Vertical In"
 GetEffectNum = 13
 
 Case "Split Vertical Out"
 GetEffectNum = 14
 
 Case "Split Horizontal In"
 GetEffectNum = 15
 
 Case "Split Horizontal Out"
 GetEffectNum = 16
 
 Case "Strips Left Down"
 GetEffectNum = 17
 
 Case "Strips Left Up"
 GetEffectNum = 18
 
 Case "Strips Right Down"
 GetEffectNum = 19
 
 Case "Strips Right Up"
 GetEffectNum = 20
 
 Case "Random Bars Horizontal"
 GetEffectNum = 21
 
 Case "Random Bars Vertical"
 GetEffectNum = 22
 
 Case "Random"
 GetEffectNum = 23
End Select
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "Bold"
If rtbText.SelBold = True Then
    rtbText.SelBold = False
Else
    rtbText.SelBold = True
End If

Case "Italic"
If rtbText.SelItalic = True Then
    rtbText.SelItalic = False
Else
    rtbText.SelItalic = True
End If

Case "Underline"
If rtbText.SelUnderline = True Then
    rtbText.SelUnderline = False
Else
    rtbText.SelUnderline = True
End If

Case "Align Left"
rtbText.SelAlignment = rtfLeft

Case "Align Right"
rtbText.SelAlignment = rtfRight

Case "Align Center"
rtbText.SelAlignment = rtfCenter

Case "Font"
CD1.Flags = cdlCFBoth Or cdlCFEffects
CD1.ShowFont
With rtbText
.SelFontName = CD1.FontName
.SelFontSize = CD1.FontSize
.SelStrikeThru = CD1.FontStrikethru
.SelUnderline = CD1.FontUnderline
.SelBold = CD1.FontBold
.SelItalic = CD1.FontItalic
End With

Case "Text Color"
CD1.ShowColor
With rtbText
.SelColor = CD1.Color
End With

Case "Find"
frmFind.Show , Me

Case "Bullet"
rtbText.SelBullet = True

Case "Date/Time"
rtbText.SelText = Date & " " & Time

Case "Insert File"
CD1.DialogTitle = "Choose file to insert..."
CD1.Filter = "Text File(*.txt)|*.txt|All Files (*.*)|*.*|"
CD1.InitDir = App.Path
CD1.ShowOpen
If Len(CD1.FileName) > 0 Then
rtbTemp.LoadFile CD1.FileName
rtbTemp.SelStart = 0 'Set selStart to 0
rtbTemp.SelLength = Len(rtbTemp.Text) - 1 'Select all text
SendMessage rtbTemp.hwnd, WM_CUT, 0, 0&
rtbText.SelText = SendMessage(rtbText.hwnd, WM_PASTE, 0, 0&)
End If
End Select

End Sub


Private Sub SetListBox()
lstEffects.AddItem "None"
lstEffects.AddItem "Box In"
lstEffects.AddItem "Box Out"
lstEffects.AddItem "Circle In"
lstEffects.AddItem "Circle Out"
lstEffects.AddItem "Wipe Up"
lstEffects.AddItem "Wipe Down"
lstEffects.AddItem "Wipe Right"
lstEffects.AddItem "Wipe Left"
lstEffects.AddItem "Vertical Blinds"
lstEffects.AddItem "Horizontal Blinds"
lstEffects.AddItem "Checkerboard Across"
lstEffects.AddItem "Checkerboard Down"
lstEffects.AddItem "Random Dissolve"
lstEffects.AddItem "Split Vertical In"
lstEffects.AddItem "Split Vertical Out"
lstEffects.AddItem "Split Horizontal In"
lstEffects.AddItem "Split Horizontal Out"
lstEffects.AddItem "Strips Left Down"
lstEffects.AddItem "Strips Left Up"
lstEffects.AddItem "Strips Right Down"
lstEffects.AddItem "Strips Right Up"
lstEffects.AddItem "Random Bars Horizontal"
lstEffects.AddItem "Random Bars Vertical"
lstEffects.AddItem "Random"
End Sub

Private Sub SetCombo()
Dim EventList(4), TransEffect(4), Duration(4)
EventList(1) = "Page Enter"
EventList(2) = "Page Exit"
EventList(3) = "Site Enter"
EventList(4) = "Site Exit"
For i = 1 To 4
cboEvent.AddItem EventList(i)
Next i
cboEvent.ListIndex = 0
For i = 1 To 4
Set q = lv.ListItems.Add(i, , EventList(i))
q.SubItems(1) = "None"
q.SubItems(2) = "0.0"

Next i

End Sub

Private Sub SetTextBox()
txtLocation.Text = App.Path & "\" & "NewPage1.htm"
txtTitle.Text = "Your Title Here"
txtBGColor.Text = "#" + RGBHexColor(lblBG.BackColor)
txtTextColor.Text = "#" + RGBHexColor(lblText.BackColor)
txtLinkColor.Text = "#" + RGBHexColor(lblALink.BackColor)
txtVLinkColor.Text = "#" + RGBHexColor(lblVLink.BackColor)
End Sub

Private Function TransWhich() As String
For i = 1 To 4
Set q = lv.ListItems.Item(i)
If q.SubItems(1) <> "None" Then
transyes$ = transyes$ + Str(i)
End If
Next i
TransWhich = transyes$
End Function
