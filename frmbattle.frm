VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmbattle 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmbattle.frx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer music1 
      Interval        =   60000
      Left            =   480
      Top             =   1920
   End
   Begin VB.Timer music2 
      Enabled         =   0   'False
      Interval        =   6002
      Left            =   960
      Top             =   1920
   End
   Begin TabDlg.SSTab attack 
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16776960
      TabCaption(0)   =   "Attack"
      TabPicture(0)   =   "frmbattle.frx":2D442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Weapon"
      TabPicture(1)   =   "frmbattle.frx":2D45E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Magic"
      TabPicture(2)   =   "frmbattle.frx":2D47A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Command4"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command4 
         Caption         =   "Heal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Beam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Use Weapon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   11
         Top             =   360
         Width           =   7095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Normal Hit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.PictureBox speed 
      BackColor       =   &H00C0C000&
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   7935
      TabIndex        =   8
      Top             =   4440
      Width           =   8000
   End
   Begin VB.Timer goodspeed 
      Interval        =   2
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer stopall 
      Enabled         =   0   'False
      Interval        =   1234
      Left            =   1320
      Top             =   360
   End
   Begin VB.Timer badspeed 
      Interval        =   3000
      Left            =   840
      Top             =   360
   End
   Begin VB.PictureBox badguy5 
      Height          =   855
      Left            =   1920
      Picture         =   "frmbattle.frx":2D496
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox badguy4 
      Height          =   855
      Left            =   1560
      OLEDropMode     =   2  'Automatic
      Picture         =   "frmbattle.frx":2EAE0
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox badguy3 
      Height          =   855
      Left            =   1200
      Picture         =   "frmbattle.frx":3012A
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox badguy2 
      Height          =   855
      Left            =   840
      Picture         =   "frmbattle.frx":309F4
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox badguy1 
      Height          =   855
      Left            =   480
      Picture         =   "frmbattle.frx":312BE
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer showstats 
      Interval        =   3700
      Left            =   360
      Top             =   360
   End
   Begin TabDlg.SSTab fightstats 
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1296
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Your Stats:"
      TabPicture(0)   =   "frmbattle.frx":31B88
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "goodhp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "mana"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.PictureBox mana 
         BackColor       =   &H00FFFF00&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   480
         Width           =   2000
      End
      Begin VB.PictureBox goodhp 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   2445
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
   End
   Begin VB.Label badhit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6840
      TabIndex        =   17
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label goodhit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label gspecial3 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   1560
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape gspecial2 
      BackColor       =   &H00FFFF00&
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   960
      Top             =   3120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Shape gspecial1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Left            =   1680
      Shape           =   2  'Oval
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Shape bspecial3 
      BackColor       =   &H0000FF00&
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   2280
      Top             =   3120
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Shape bspecial2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape bspecial1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H000000FF&
      FillStyle       =   2  'Horizontal Line
      Height          =   615
      Left            =   2280
      Shape           =   2  'Oval
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Image badguy 
      Height          =   480
      Left            =   1920
      Picture         =   "frmbattle.frx":31BA4
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image goodguy 
      Height          =   480
      Left            =   6840
      Picture         =   "frmbattle.frx":3246E
      Top             =   3000
      Width           =   480
   End
End
Attribute VB_Name = "frmbattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim badhp As Integer
Dim badattack As Integer

Private Sub badspeed_Timer()
Dim MyValue, Response
Randomize
   MyValue = Int((7 * Rnd) + 1)
 If MyValue = 1 Then
 badhit.Caption = (Int((badattack * Rnd) + 3) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
 bspecial3.Visible = True
 End If
  If MyValue = 2 Then
 badhit.Caption = (Int((badattack * Rnd) + 3) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial2.Visible = True
 End If
  If MyValue = 3 Then
 badhit.Caption = (Int((badattack * Rnd) + 4) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial1.Visible = True
 End If
  If MyValue = 4 Then
 badhit.Caption = (Int((badattack * Rnd) + 5) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial2.Visible = True
 End If
  If MyValue = 5 Then
 badhit.Caption = (Int((badattack * Rnd) + 6) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial1.Visible = True
 End If
  If MyValue = 6 Then
 badhit.Caption = (Int((badattack * Rnd) + 7) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial3.Visible = True
 End If
  If MyValue = 7 Then
 badhit.Caption = (Int((badattack * Rnd) + 8) - Form1.defense.Text)
 Form1.goodhp.Text = Form1.goodhp.Text - badhit.Caption
  bspecial3.Visible = True
 End If
 stopall.Enabled = True
End Sub

Private Sub Command1_Click()
Dim MyValue, Response
Randomize
   MyValue = Int((4 * Rnd) + 1)
 If MyValue = 1 Then
 goodhit.Caption = Int((Form1.attack.Text * Rnd) + 1)
 badhp = badhp - goodhit.Caption
 gspecial3.Visible = True
 stopall.Enabled = True
 End If
  If MyValue = 2 Then
 goodhit.Caption = Int((Form1.attack.Text * Rnd) + 2)
 badhp = badhp - goodhit.Caption
 gspecial3.Visible = True
 stopall.Enabled = True
 End If
  If MyValue = 3 Then
 goodhit.Caption = Int((Form1.attack.Text * Rnd) + 3)
 badhp = badhp - goodhit.Caption
 gspecial3.Visible = True
 stopall.Enabled = True
 End If
  If MyValue = 4 Then
 goodhit.Caption = Int((Form1.attack.Text * Rnd) + 4)
 badhp = badhp - goodhit.Caption
 gspecial3.Visible = True
 stopall.Enabled = True
 End If
 attack.Visible = False
 speed.Width = 1000
 gspecial3.Visible = True
End Sub

Private Sub Command2_Click()
If Form1.weapon.Text = "Knife" Then
 goodhit.Caption = Int((5 * Rnd) + 2)
 badhp = badhp - goodhit.Caption
 gspecial2.Visible = True
 stopall.Enabled = True
 attack.Visible = False
 End If
 speed.Width = 100
End Sub

Private Sub Form_Load()
Dim MyValue, Response
Randomize
   MyValue = Int((7 * Rnd) + 1)
 If MyValue = 1 Then
 badguy.Picture = badguy1.Picture
 badhp = 7
 badattack = 5
 badspeed.Interval = 4000
 End If
  If MyValue = 2 Then
 badguy.Picture = badguy2.Picture
 badhp = 12
  badattack = 6
   badspeed.Interval = 3800
 End If
  If MyValue = 3 Then
 badguy.Picture = badguy3.Picture
 badhp = 15
  badattack = 4
   badspeed.Interval = 3899
 End If
  If MyValue = 4 Then
 badguy.Picture = badguy4.Picture
 badhp = 18
  badattack = 5
   badspeed.Interval = 5000
 End If
  If MyValue = 5 Then
 badguy.Picture = badguy5.Picture
 badhp = 24
  badattack = 5
   badspeed.Interval = 4700
 End If
 If MyValue = 6 Then
 badguy.Picture = badguy1.Picture
 badhp = 8
  badattack = 4
   badspeed.Interval = 3000
 End If
 If MyValue = 6 Then
 badguy.Picture = badguy1.Picture
 badhp = 6
  badattack = 4
   badspeed.Interval = 4000
 End If
 Form2.MMControl1.Command = "Stop"
 MMControl1.Command = "Seek"
 MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
  'MMControl1.Command = "Seek"
MMControl1.DeviceType = "Sequencer"
 MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
PlayLoop = True
End Sub

Private Sub goodspeed_Timer()
On Error Resume Next
If speed.Width = 8000 Then
attack.Visible = True
Else
speed.Width = speed.Width + 100
End If
If badhp < 0 Then
MsgBox "You won!"
MMControl1.Command = "Stop"
 Form2.MMControl1.Command = "Seek"
  Form2.MMControl1.FileName = App.Path + "\shadowhaven.mid"
  'MMControl1.Command = "Seek"
 Form2.MMControl1.DeviceType = "Sequencer"
  Form2.MMControl1.Command = "Seek"
 Form2.MMControl1.Command = "Open"
 Form2.MMControl1.Command = "Play"
Form1.goodhp.Text = Form1.goodhp.Text + 8
Form1.attack.Text = Form1.attack.Text + 1
Form1.defense.Text = Form1.defense + 1
Form1.mana.Text = Form1.mana.Text + 1
Form2.Enabled = True
Form2.Show
'frmbattle.Hide
Unload Me
goodspeed.Enabled = False
End If
If Form1.goodhp.Text < 0 Then
MsgBox "You got killed!"
goodspeed.Enabled = False
End
End If
If badhit.Caption < 0 Then badhit.Caption = "1"
End Sub

Private Sub music1_Timer()
music2.Enabled = True
music1.Enabled = False
End Sub

Private Sub music2_Timer()
MMControl1.Command = "Stop"
MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
MMControl1.DeviceType = "Sequencer"
MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
PlayLoop = True
music1.Enabled = True
music2.Enabled = False
End Sub

Private Sub showstats_Timer()
goodhp.Width = (Form1.goodhp.Text * 100)
mana.Width = (Form1.mana.Text * 100)
If fightstats.Visible = False Then
fightstats.Visible = True
Else
fightstats.Visible = False
End If

End Sub

Private Sub stopall_Timer()
gspecial1.Visible = False
gspecial2.Visible = False
gspecial3.Visible = False
bspecial1.Visible = False
bspecial2.Visible = False
bspecial3.Visible = False
badhit.Caption = ""
goodhit.Caption = ""
stopall.Enabled = False
End Sub
