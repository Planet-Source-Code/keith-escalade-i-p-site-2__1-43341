VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "I.P. Site v1.01"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmIpSite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   8
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "HTML"
      TabPicture(0)   =   "FrmIpSite.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Visitors"
      TabPicture(1)   =   "FrmIpSite.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check3"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "List1"
      Tab(1).Control(4)=   "List2"
      Tab(1).Control(5)=   "List3"
      Tab(1).Control(6)=   "List4"
      Tab(1).Control(7)=   "hits"
      Tab(1).Control(8)=   "Label5"
      Tab(1).Control(9)=   "Label6"
      Tab(1).Control(10)=   "Label7"
      Tab(1).Control(11)=   "Label8"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Blocked"
      TabPicture(2)   =   "FrmIpSite.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command16"
      Tab(2).Control(1)=   "Command17"
      Tab(2).Control(2)=   "Text12"
      Tab(2).Control(3)=   "List9"
      Tab(2).Control(4)=   "Check5"
      Tab(2).Control(5)=   "Text4"
      Tab(2).Control(6)=   "Command6"
      Tab(2).Control(7)=   "Command5"
      Tab(2).Control(8)=   "Text3"
      Tab(2).Control(9)=   "List5"
      Tab(2).Control(10)=   "Label20"
      Tab(2).Control(11)=   "Label10"
      Tab(2).Control(12)=   "Label9"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Lists"
      TabPicture(3)   =   "FrmIpSite.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command14"
      Tab(3).Control(1)=   "Command15"
      Tab(3).Control(2)=   "Text11"
      Tab(3).Control(3)=   "List8"
      Tab(3).Control(4)=   "Command12"
      Tab(3).Control(5)=   "Command13"
      Tab(3).Control(6)=   "Text10"
      Tab(3).Control(7)=   "List7"
      Tab(3).Control(8)=   "Command8"
      Tab(3).Control(9)=   "Command7"
      Tab(3).Control(10)=   "Text5"
      Tab(3).Control(11)=   "List6"
      Tab(3).Control(12)=   "Label17"
      Tab(3).Control(13)=   "Label16"
      Tab(3).Control(14)=   "Label11"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "AIM"
      TabPicture(4)   =   "FrmIpSite.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Check6"
      Tab(4).Control(1)=   "Text8"
      Tab(4).Control(2)=   "Command9"
      Tab(4).Control(3)=   "Text7"
      Tab(4).Control(4)=   "Text6"
      Tab(4).Control(5)=   "CD1"
      Tab(4).Control(6)=   "Label13"
      Tab(4).Control(7)=   "Label12"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Help"
      TabPicture(5)   =   "FrmIpSite.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text13"
      Tab(5).Control(1)=   "WB2"
      Tab(5).Control(2)=   "Label14"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Broadcast"
      TabPicture(6)   =   "FrmIpSite.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command11"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Command10"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Text9"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Command1"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Command2"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Check1"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Text2"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "Label18"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Label4"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).ControlCount=   9
      Begin VB.TextBox Text13 
         Height          =   1245
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         Text            =   "FrmIpSite.frx":0506
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-"
         Height          =   375
         Left            =   -72000
         TabIndex        =   64
         Top             =   5880
         Width           =   375
      End
      Begin VB.CommandButton Command17 
         Caption         =   "+"
         Height          =   375
         Left            =   -72360
         TabIndex        =   65
         Top             =   5880
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   405
         Left            =   -74880
         TabIndex        =   63
         Top             =   5880
         Width           =   2415
      End
      Begin VB.ListBox List9 
         Height          =   2010
         Left            =   -74880
         TabIndex        =   62
         Top             =   3720
         Width           =   3255
      End
      Begin VB.CommandButton Command14 
         Caption         =   "-"
         Height          =   375
         Left            =   -65880
         TabIndex        =   57
         Top             =   5880
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+"
         Height          =   375
         Left            =   -66240
         TabIndex        =   58
         Top             =   5880
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   405
         Left            =   -72120
         TabIndex        =   56
         Top             =   5880
         Width           =   5775
      End
      Begin VB.ListBox List8 
         Height          =   2205
         ItemData        =   "FrmIpSite.frx":0916
         Left            =   -72120
         List            =   "FrmIpSite.frx":0918
         TabIndex        =   55
         Top             =   3600
         Width           =   6615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-"
         Height          =   375
         Left            =   -72720
         TabIndex        =   53
         Top             =   5880
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Caption         =   "+"
         Height          =   375
         Left            =   -73080
         TabIndex        =   54
         Top             =   5880
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   -74880
         TabIndex        =   51
         Top             =   5880
         Width           =   1695
      End
      Begin VB.ListBox List7 
         Height          =   2205
         ItemData        =   "FrmIpSite.frx":091A
         Left            =   -74880
         List            =   "FrmIpSite.frx":091C
         TabIndex        =   50
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Change on startup"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "View"
         Height          =   375
         Left            =   -70800
         TabIndex        =   47
         Top             =   5880
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Don't send them anything (page cannot be displayed)"
         Height          =   255
         Left            =   -71520
         TabIndex        =   46
         Top             =   6000
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Copy"
         Height          =   375
         Left            =   -71760
         TabIndex        =   45
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5880
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Beep on visit (1 beep per visit)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   6000
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start broadcasting"
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop broadcasting"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Broadcast hits"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   1455
         Left            =   -72240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "FrmIpSite.frx":091E
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   960
         Width           =   8415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "info.htm"
         Height          =   375
         Left            =   -74880
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -71880
         TabIndex        =   31
         Top             =   480
         Width           =   6375
      End
      Begin VB.TextBox Text6 
         Height          =   4455
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   1800
         Width           =   9375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-"
         Height          =   375
         Left            =   -65880
         TabIndex        =   27
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   375
         Left            =   -66240
         TabIndex        =   26
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   -74880
         TabIndex        =   25
         Top             =   2880
         Width           =   8535
      End
      Begin VB.ListBox List6 
         Height          =   2010
         ItemData        =   "FrmIpSite.frx":0955
         Left            =   -74880
         List            =   "FrmIpSite.frx":0957
         TabIndex        =   24
         Top             =   720
         Width           =   9375
      End
      Begin VB.TextBox Text4 
         Height          =   5175
         Left            =   -71520
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   375
         Left            =   -72000
         TabIndex        =   20
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   375
         Left            =   -72360
         TabIndex        =   19
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   -74880
         TabIndex        =   18
         Top             =   2880
         Width           =   2415
      End
      Begin VB.ListBox List5 
         Height          =   2010
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear names/times"
         Height          =   375
         Left            =   -70200
         TabIndex        =   16
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Ip's/times"
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   5520
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   4740
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin VB.ListBox List2 
         Height          =   4740
         Left            =   -70200
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.ListBox List3 
         Height          =   4740
         Left            =   -71520
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   4740
         Left            =   -66480
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5775
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   10186
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   8
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "HTML"
         TabPicture(0)   =   "FrmIpSite.frx":0959
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Preview"
         TabPicture(1)   =   "FrmIpSite.frx":0975
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "WB1"
         Tab(1).Control(1)=   "Label19"
         Tab(1).Control(2)=   "Label15"
         Tab(1).ControlCount=   3
         Begin VB.TextBox Text1 
            Height          =   5175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   9
            Top             =   480
            Width           =   9135
         End
         Begin SHDocVwCtl.WebBrowser WB1 
            Height          =   4455
            Left            =   -74880
            TabIndex        =   10
            Top             =   840
            Width           =   9015
            ExtentX         =   15901
            ExtentY         =   7858
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "res://C:\WINDOWS\system32\shdoclc.dll/offcancl.htm#http:///"
         End
         Begin VB.Label Label19 
            Height          =   255
            Left            =   -74880
            TabIndex        =   61
            Top             =   480
            Width           =   9015
         End
         Begin VB.Label Label15 
            Height          =   255
            Left            =   -74880
            TabIndex        =   42
            Top             =   5400
            Width           =   9015
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   -75000
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin SHDocVwCtl.WebBrowser WB2 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   9375
         ExtentX         =   16536
         ExtentY         =   9763
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINDOWS\system32\shdoclc.dll/offcancl.htm#http:///"
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Blocked AIM names"
         Height          =   255
         Left            =   -74880
         TabIndex        =   66
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Link (your page)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "With"
         Height          =   255
         Left            =   -72120
         TabIndex        =   59
         Top             =   3360
         Width           =   6615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Replace"
         Height          =   255
         Left            =   -74880
         TabIndex        =   52
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label hits 
         Alignment       =   1  'Right Justify
         Caption         =   "Hits 0"
         Height          =   255
         Left            =   -68640
         TabIndex        =   48
         Top             =   6000
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Format :"
         Height          =   255
         Left            =   -73080
         TabIndex        =   41
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Help"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   480
         Width           =   9375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Link text"
         Height          =   255
         Left            =   -72720
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Your AIM profile (what AIM users will see when they view your profile)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1560
         Width           =   9375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Random list, %RAND% replaces itself with a random item from the list below"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   9375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Blocked Ip text"
         Height          =   255
         Left            =   -71520
         TabIndex        =   22
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Blocked Ip's"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Visitor's Ip"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "AIM visitors screen name"
         Height          =   255
         Left            =   -69960
         TabIndex        =   13
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Times"
         Height          =   255
         Left            =   -71520
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Times"
         Height          =   255
         Left            =   -66480
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Line Line7 
      X1              =   9480
      X2              =   9480
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line6 
      X1              =   9120
      X2              =   9120
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   9840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   9840
      X2              =   9840
      Y1              =   6840
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9840
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   I.P. Site v1.01... Port 80 must not be in use..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sAimName As Boolean
Dim WinData$
Dim TheFile$
Dim RiD$
Dim NoSName$
Private Sub Check1_Click()
If Check1.Value = 1 Then Text2.Visible = True: Label4.Visible = True: Exit Sub
If Check1.Value = 0 Then Text2.Visible = False: Label4.Visible = False: Exit Sub
End Sub
Private Sub Check6_Click()
Close #1
If Check6.Value = 0 Then Kill App.Path & "/AimInfo.ini": Exit Sub
If Check6.Value = 1 Then Open App.Path & "/AimInfo.ini" For Output As #1
Select Case Text8.text
Case Is = ""
Close #1
Exit Sub
Case Is <> ""
Print #1, Text8.text
Close #1
End Select
End Sub
Private Sub Command1_Click()
On Error GoTo dieerror
If Text1.text = "" Then: Exit Sub
Command1.Enabled = False
Command2.Enabled = True
Winsock1(0).LocalPort = 80
Winsock1(0).Listen
Exit Sub
dieerror:
Winsock1(0).Close
Command1.Enabled = True
Command2.Enabled = False
MsgBox "Port 80 might be in use, please free it and try again", vbInformation
End Sub
Private Sub Command10_Click()
Clipboard.SetText Text9.text
End Sub
Private Sub Command11_Click()
Shell "Explorer http://" & Winsock1(0).LocalIP
End Sub
Private Sub Command12_Click()
On Error Resume Next
List7.RemoveItem List7.ListIndex
End Sub
Private Sub Command13_Click()
If Text10.text = "%RAND%" Then: Exit Sub
If Text10.text = "%TIME%" Then Exit Sub
If Text10.text = "%DATE%" Then Exit Sub
If Text10.text = "%NOW%" Then Exit Sub
If Text10.text = "%HITS%" Then Exit Sub
If Text10.text = "%IP%" Then Exit Sub
List7.AddItem Text10.text
Text10.text = ""
End Sub
Private Sub Command14_Click()
On Error Resume Next
List8.RemoveItem List8.ListIndex
End Sub
Private Sub Command15_Click()
List8.AddItem Text11.text
Text11.text = ""
End Sub
Private Sub Command16_Click()
On Error Resume Next
List9.RemoveItem List9.ListIndex
End Sub
Private Sub Command17_Click()
List9.AddItem Text12.text
Text12.text = ""
End Sub
Private Sub Command2_Click()
Winsock1(0).Close
Command2.Enabled = False
Command1.Enabled = True
End Sub
Private Sub Command3_Click()
On Error Resume Next
List1.Clear
Kill App.Path & "/VisitorsIp.ini"
List3.Clear
Kill App.Path & "/IpTimes.ini"
End Sub
Private Sub Command4_Click()
On Error Resume Next
List2.Clear
Kill App.Path & "/VisitorSN.ini"
List4.Clear
Kill App.Path & "/AIMTimes.ini"
End Sub
Private Sub Command5_Click()
List5.AddItem Text3.text
Text3.text = ""
End Sub
Private Sub Command6_Click()
On Error Resume Next
List5.RemoveItem List5.ListIndex
End Sub
Private Sub Command7_Click()
List6.AddItem Text5.text
Text5.text = ""
End Sub
Private Sub Command8_Click()
On Error Resume Next
List6.RemoveItem List6.ListIndex
End Sub
Private Sub Command9_Click()
On Error GoTo dieerror
CD1.CancelError = True
CD1.Filter = "AIM HTML files|info.htm"
CD1.FileName = "C:\WINDOWS\aim95\info.htm"
CD1.ShowOpen
Open CD1.FileName For Output As #1
Print #1, "<a href=""http://" & Winsock1(0).LocalIP & "/AimSn=%n"" target=""_self"">" & Text7.text & "</a>"
Close #1
Text8.text = CD1.FileName
If Check6.Value = 1 Then Open App.Path & "/AimInfo.ini" For Output As #1
Print #1, CD1.FileName
Close #1
MsgBox "If you have AIM already opened, you must restart it in order for changes to take place!", vbInformation
Exit Sub
dieerror:
Exit Sub
End Sub
Private Sub Form_Load()
Text9.text = "http://" & Winsock1(0).LocalIP
On Error GoTo ErrorFix
Form1.Show
TheFile = App.Path & "/AimInfo.ini"
Open App.Path & "/AimInfo.ini" For Input As #1
Input #1, sText$
DoEvents
Text8.text = sText$
If sText$ <> "" Then Check6.Value = 1
Close #1
TheFile = App.Path & "/Page.txt"
Open App.Path & "/Page.txt" For Input As #1
While Not EOF(1)
Input #1, sText1$
If EOF(1) = True Then Text1.text = Text1.text & RepComma(sText1$)
If EOF(1) = False Then Text1.text = Text1.text & RepComma(sText1$) & vbCrLf
DoEvents
Wend
Close #1
TheFile = App.Path & "/VisitorsIp.ini"
Open App.Path & "/VisitorsIp.ini" For Input As #1
While Not EOF(1)
Input #1, sText2$
DoEvents
List1.AddItem sText2$
Wend
Close #1
TheFile = App.Path & "/VisitorSN.ini"
Open App.Path & "/VisitorSN.ini" For Input As #1
While Not EOF(1)
Input #1, sText800$
DoEvents
List2.AddItem sText800$
Wend
Close #1
TheFile = App.Path & "/BlockedIps.ini"
Open App.Path & "/BlockedIps.ini" For Input As #1
While Not EOF(1)
Input #1, sText6$
DoEvents
List5.AddItem sText6$
Wend
Close #1
TheFile = App.Path & "/BlockedNames.ini"
Open App.Path & "/BlockedNames.ini" For Input As #1
While Not EOF(1)
Input #1, sText7$
DoEvents
List9.AddItem sText7$
Wend
Close #1
TheFile = App.Path & "/BlockedIpText.ini"
Open App.Path & "/BlockedIpText.ini" For Input As #1
While Not EOF(1)
Input #1, sText8$
If EOF(1) = True Then Text4.text = Text4.text & RepComma(sText8$)
If EOF(1) = False Then Text4.text = Text4.text & RepComma(sText8$) & vbCrLf
DoEvents
Wend
Close #1
TheFile = App.Path & "/RandomList.ini"
Open App.Path & "/RandomList.ini" For Input As #1
While Not EOF(1)
Input #1, sText9$
DoEvents
List6.AddItem RepComma(sText9$)
Wend
Close #1
TheFile = App.Path & "/Variables.ini"
Open App.Path & "/Variables.ini" For Input As #1
While Not EOF(1)
Input #1, sText10$
List7.AddItem RepComma(sText10$)
Wend
Close #1
TheFile = App.Path & "/WithText.ini"
Open App.Path & "/WithText.ini" For Input As #1
While Not EOF(1)
Input #1, sText11$
DoEvents
List8.AddItem RepComma(sText11$)
Wend
Close #1
TheFile = App.Path & "/LinkText.ini"
Open App.Path & "/LinkText.ini" For Input As #1
Input #1, sText12$
DoEvents
Text7.text = RepComma(sText12$)
Close #1
TheFile = App.Path & "/AIMProfile.ini"
Open App.Path & "/AIMProfile.ini" For Input As #1
While Not EOF(1)
Input #1, sText13$
If EOF(1) = True Then Text6.text = Text6.text & RepComma(sText13$)
If EOF(1) = False Then Text6.text = Text6.text & RepComma(sText13$) & vbCrLf
DoEvents
Wend
Close #1
TheFile = App.Path & "/BroadcastHits.ini"
Open App.Path & "/BroadcastHits.ini" For Input As #1
Input #1, sText14$
DoEvents
If sText14$ = "1" Then Check1.Value = 1: Label4.Visible = True: Text2.Visible = True
Close #1
TheFile = App.Path & "/AIMTimes.ini"
Open App.Path & "/AIMTimes.ini" For Input As #1
While Not EOF(1)
Input #1, sText5$
DoEvents
List4.AddItem sText5$
Wend
Close #1
TheFile = App.Path & "/IPTimes.ini"
Open App.Path & "/IPTimes.ini" For Input As #1
While Not EOF(1)
Input #1, sText69$
DoEvents
List3.AddItem sText69$
Wend
Close #1
 If Check6.Value = 1 Then: Open Text8.text For Output As #1: Print #1, "<a href=""http://" & Winsock1(0).LocalIP & "/AimSn=%n"" target=""_self"">" & Text7.text & "</a>": Close #1
WB1.Navigate "about:blank"
WB2.Navigate "about:blank"
Exit Sub
ErrorFix:
Open TheFile$ For Output As #1
Close #1
End Sub
Private Sub Label1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
Private Sub Label2_Click()
Close #1
Open App.Path & "/Page.txt" For Output As #1
Print #1, RepCom(Text1.text)
Close #1
Open App.Path & "/VisitorsIp.ini" For Output As #1
For X = 0 To List1.ListCount - 1
Print #1, List1.List(X)
Next X
Close #1
Open App.Path & "/VisitorSN.ini" For Output As #1
For X = 0 To List4.ListCount - 1
Print #1, List4.List(X)
Next X
Close #1
Open App.Path & "/IpTimes.ini" For Output As #1
For X = 0 To List3.ListCount - 1
Print #1, List3.List(X)
Next X
Close #1
Open App.Path & "/VisitorSN.ini" For Output As #1
For X = 0 To List2.ListCount - 1
Print #1, List2.List(X)
Next X
Close #1
Open App.Path & "/AIMTimes.ini" For Output As #1
For X = 0 To List4.ListCount - 1
Print #1, List4.List(X)
Next X
Close #1
Open App.Path & "/BlockedIps.ini" For Output As #1
For X = 0 To List5.ListCount - 1
Print #1, List5.List(X)
Next X
Close #1
Open App.Path & "/BlockedNames.ini" For Output As #1
For X = 0 To List9.ListCount - 1
Print #1, List9.List(X)
Next X
Close #1
Open App.Path & "/BlockedIpText.ini" For Output As #1
Print #1, RepCom(Text4.text)
Close #1
Open App.Path & "/RandomList.ini" For Output As #1
For X = 0 To List6.ListCount - 1
Print #1, RepCom(List6.List(X))
Next X
Close #1
Open App.Path & "/Variables.ini" For Output As #1
For X = 0 To List7.ListCount - 1
Print #1, RepCom(List7.List(X))
Next X
Close #1
Open App.Path & "/WithText.ini" For Output As #1
For X = 0 To List8.ListCount - 1
Print #1, RepCom(List8.List(X))
Next X
Close #1
Open App.Path & "/AIMProfile.ini" For Output As #1
Print #1, RepCom(Text6.text)
Close #1
Open App.Path & "/LinkText.ini" For Output As #1
Print #1, RepCom(Text7.text)
Close #1
Open App.Path & "/BroadcastHits.ini" For Output As #1
Print #1, Check1.Value
Close #1
Quit
End Sub
Private Sub Label3_Click()
Minimize Me
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "Help" Then WB2.Navigate "about:" & Text13.text
End Sub
Private Sub SSTab2_Click(PreviousTab As Integer)
If SSTab2.Caption = "Preview" Then Open App.Path & "/PreviewHTML.html" For Output As #1: Print #1, ConvText(Text1.text): Close #1: WB1.Navigate App.Path & "/PreviewHTML.html"
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command13_Click
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command15_Click
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command17_Click
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command5_Click
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command7_Click
End Sub
Private Sub WB1_StatusTextChange(ByVal text As String)
Label15.Caption = text
End Sub
Private Sub WB1_TitleChange(ByVal text As String)
Label19.Caption = text
End Sub
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Winsock1(0).Close: Winsock1(0).Accept requestID
RiD = requestID
AddIp (Winsock1(0).RemoteHostIP)
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data$
Winsock1(0).GetData Data
WinData = Data
If AimName(Data) = "" Then sAimName = False
If AimName(Data) <> "" Then sAimName = True: NoSName = NoSpaces(AimName(Data)): AddName (AimName(Data))
sSenddData
End Sub
Private Sub Winsock1_SendComplete(Index As Integer)
Pause 0.5
Winsock1(0).Close
Winsock1(0).Listen
End Sub
Function AimName(Data As String)
On Error GoTo noerror
AimName_1% = InStr(1, Data$, "GET /AimSn=")
LenAimName_1% = 11
AimName_2% = InStr(AimName_1% + 1, Data$, " HTTP/1.0")
AimName_3$ = Mid(Data$, AimName_1%, AimName_2% - AimName_1%)
AimName = Right(AimName_3$, Len(AimName_3$) - LenAimName_1%)
Exit Function
noerror:
Exit Function
End Function
Function NoSpaces(Data$)
Data = Replace(Data, " ", "")
End Function
Function ConvText(DaText$)
On Error Resume Next
ConvText = Replace(DaText$, "%TIME%", Time)
ConvText = Replace(ConvText, "%RAND%", List6.List(RandomNumber(List6.ListCount - 1, 0)))
ConvText = Replace(ConvText, "%DATE%", Date)
ConvText = Replace(ConvText, "%NOW%", Now)
ConvText = Replace(ConvText, "%HITS%", List1.ListCount)
ConvText = Replace(ConvText, "%IP%", Winsock1(0).RemoteHostIP)
For X = 0 To List7.ListCount - 1
ConvText = Replace(ConvText, List7.List(X), List8.List(X))
Next X
End Function
Function RepCom(DaText$)
RepCom = Replace(DaText, ",", "%CommA")
End Function
Function RepComma(DaText$)
RepComma = Replace(DaText, "%CommA", ",")
End Function
Function IsBan(What$, Whichlist As ListBox) As Boolean
For X = 0 To Whichlist.ListCount - 1
If NoSpaces(LCase(What$)) = NoSpaces(LCase(Whichlist.List(X))) Then IsBanned = True
Next X
End Function
Public Sub AddIp(sIp$)
On Error Resume Next
For X = 0 To List1.ListCount - 1
If sIp = List1.List(X) Then List3.List(X) = List3.List(X) + 1: Exit Sub
Next X
List1.AddItem sIp
List3.AddItem "1"
hits.Caption = "Hits " & List1.ListCount
If Check3.Value = 1 Then Beep
End Sub
Public Sub AddName(sName$)
On Error Resume Next
For X = 0 To List2.ListCount - 1
If sName$ = List2.List(X) Then List4.List(X) = List4.List(X) + 1: Exit Sub
Next X
List2.AddItem sName
List4.AddItem "1"
hits.Caption = "Hits " & List1.ListCount
If Check3.Value = 1 Then Beep
End Sub
Public Sub sSenddData()
For X = 0 To List5.ListCount - 1
If List5.List(X) = Winsock1(0).RemoteHostIP Then If Check5.Value = 1 Then Winsock1(0).Close: Winsock1(0).Listen: Exit Sub
If List5.List(X) = Winsock1(0).RemoteHostIP Then If Check5.Value = 0 Then Winsock1(0).SendData ConvText(Text4.text): Exit Sub
Next X
For X = 0 To List9.ListCount - 1
If NoSpaces(List9.List(X)) = NoSpaces(AimName(WinData)) Then If Check5.Value = 1 Then Winsock1(0).Close: Winsock1(0).Listen: GoTo quitit: Exit Sub
If NoSpaces(List9.List(X)) = NoSpaces(AimName(WinData)) Then If Check5.Value = 0 Then Winsock1(0).SendData ConvText(Text4.text): Exit Sub
Next X
If sAimName = True Then Winsock1(0).SendData ConvText(Text6.text): Exit Sub
If sAimName = False Then Winsock1(0).SendData ConvText(Text1.text): Exit Sub
Exit Sub
quitit:
Exit Sub
End Sub
