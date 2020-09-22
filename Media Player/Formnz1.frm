VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multimedia Player by A s i F"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "Formnz1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "V -"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "V +"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton g 
      Caption         =   "Sound"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton f 
      Caption         =   "Mute"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton p1 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton p 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Look In"
      BeginProperty Font 
         Name            =   "ChasThirdSH"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8175
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      Begin VB.TextBox t2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Americana BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   7440
         Width           =   2295
      End
      Begin VB.DirListBox l1 
         Height          =   1665
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.FileListBox l 
         Height          =   2040
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   2415
      End
      Begin VB.DriveListBox d 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         X1              =   0
         X2              =   2880
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label t 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   360
         TabIndex        =   11
         Top             =   5400
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label l6 
         Alignment       =   2  'Center
         Caption         =   "File Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   6960
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Full Screen"
      BeginProperty Font 
         Name            =   "Busorama Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   7320
      Width           =   1575
   End
   Begin MediaPlayerCtl.MediaPlayer a 
      Height          =   6495
      Left            =   3600
      TabIndex        =   7
      Top             =   360
      Width           =   6735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a.DisplaySize = mpFullScreen
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
a.SetFocus
SendKeys "{up}"
SendKeys "{up}"
SendKeys "{up}"
End Sub
Private Sub Command4_Click()
a.SetFocus
SendKeys "{down}"
SendKeys "{down}"
SendKeys "{down}"
End Sub
Private Sub f_Click()
a.Mute = 1
g.Visible = 1
End Sub
Private Sub d_Change()
l1.Path = d.Drive
End Sub
Private Sub Form_Load()
p1.Visible = 0
g.Visible = 0
a.EnablePositionControls = 1
End Sub
Private Sub g_Click()
g.Visible = 0
a.Mute = False
End Sub
Private Sub l_Click()
a.FileName = l.Path + "\" + l.FileName
t = a.FileName
t2 = l.FileName
End Sub
Private Sub l1_Change()
l.Path = l1.Path
End Sub
Private Sub p_Click()
On Error Resume Next
a.Pause
p1.Visible = 1
End Sub
Private Sub p1_Click()
On Error Resume Next
p1.Visible = 0
a.Play
End Sub
