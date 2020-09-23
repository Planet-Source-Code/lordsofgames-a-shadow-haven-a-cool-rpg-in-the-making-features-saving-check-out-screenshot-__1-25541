VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shadow Haven--Ancient Duels 2"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer music 
      Interval        =   56000
      Left            =   0
      Top             =   5640
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer map3 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2640
      Top             =   5160
   End
   Begin VB.Timer bounds 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   0
      Top             =   5160
   End
   Begin VB.Timer map2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2280
      Top             =   5160
   End
   Begin VB.PictureBox fighterright 
      Enabled         =   0   'False
      Height          =   615
      Left            =   840
      Picture         =   "mainfrm.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox fighterleft 
      Enabled         =   0   'False
      Height          =   615
      Left            =   720
      Picture         =   "mainfrm.frx":08CA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox fighterdown 
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      Picture         =   "mainfrm.frx":1194
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox fighterup 
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      Picture         =   "mainfrm.frx":1A5E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer map1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1920
      Top             =   5160
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Image fighter 
         Height          =   480
         Left            =   1440
         Picture         =   "mainfrm.frx":2328
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   109
         Left            =   4320
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   108
         Left            =   3840
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   107
         Left            =   3360
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   106
         Left            =   2880
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   105
         Left            =   2400
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   104
         Left            =   1920
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   103
         Left            =   1440
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   102
         Left            =   960
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   101
         Left            =   480
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   100
         Left            =   0
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   99
         Left            =   4320
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   98
         Left            =   3840
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   97
         Left            =   3360
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   96
         Left            =   2880
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   95
         Left            =   2400
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   94
         Left            =   1920
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   93
         Left            =   1440
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   92
         Left            =   960
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   91
         Left            =   480
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   90
         Left            =   0
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   89
         Left            =   4320
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   88
         Left            =   3840
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   87
         Left            =   3360
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   86
         Left            =   2880
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   85
         Left            =   2400
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   84
         Left            =   1920
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   83
         Left            =   1440
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   82
         Left            =   960
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   81
         Left            =   480
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   80
         Left            =   0
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   79
         Left            =   4320
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   78
         Left            =   3840
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   77
         Left            =   3360
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   76
         Left            =   2880
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   75
         Left            =   2400
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   74
         Left            =   1920
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   73
         Left            =   1440
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   72
         Left            =   960
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   71
         Left            =   480
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   70
         Left            =   0
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   69
         Left            =   4320
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   68
         Left            =   3840
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   67
         Left            =   3360
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   66
         Left            =   2880
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   65
         Left            =   2400
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   64
         Left            =   1920
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   63
         Left            =   1440
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   62
         Left            =   960
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   61
         Left            =   480
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   60
         Left            =   0
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   59
         Left            =   4320
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   58
         Left            =   3840
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   57
         Left            =   3360
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   56
         Left            =   2880
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   55
         Left            =   2400
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   54
         Left            =   1920
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   53
         Left            =   1440
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   52
         Left            =   960
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   51
         Left            =   480
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   50
         Left            =   0
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   49
         Left            =   4320
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   48
         Left            =   3840
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   47
         Left            =   3360
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   46
         Left            =   2880
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   45
         Left            =   2400
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   44
         Left            =   1920
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   43
         Left            =   1440
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   42
         Left            =   960
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   41
         Left            =   480
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   40
         Left            =   0
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   29
         Left            =   4320
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   28
         Left            =   3840
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   27
         Left            =   3360
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   26
         Left            =   2880
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   25
         Left            =   2400
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   24
         Left            =   1920
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   23
         Left            =   1440
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   22
         Left            =   960
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   21
         Left            =   480
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   20
         Left            =   0
         Top             =   960
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   19
         Left            =   4320
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   18
         Left            =   3840
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   17
         Left            =   3360
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   16
         Left            =   2880
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   15
         Left            =   2400
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   14
         Left            =   1920
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   13
         Left            =   1440
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   12
         Left            =   960
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   11
         Left            =   480
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   10
         Left            =   0
         Top             =   480
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   9
         Left            =   4320
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   8
         Left            =   3840
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   7
         Left            =   3360
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   6
         Left            =   2880
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   5
         Left            =   2400
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   4
         Left            =   1920
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   3
         Left            =   1440
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   2
         Left            =   960
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   1
         Left            =   480
         Top             =   0
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2520
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub bounds_Timer()
'We don't need this anymore.
'This would cause it to have errors. I cancelled it.
'If fighter.Left > 4320 Then fighter.Left = 0 'Incase player goes out.
'If fighter.Left < -2 Then fighter.Left = 4320
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MyValue, Response
If KeyCode = vbKeyUp Then
fighter.Picture = fighterup.Picture 'go up
fighter.Top = fighter.Top - 480
Randomize
   MyValue = Int((4 * Rnd) + 1)
 If MyValue = 1 Then
 End If
 If MyValue = 2 Then
 frmbattle.Show
 frmbattle.goodspeed.Enabled = True
 Form2.MMControl1.Command = "Stop"
 frmbattle.MMControl1.Command = "Seek"
  frmbattle.MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
  'MMControl1.Command = "Seek"
 frmbattle.MMControl1.DeviceType = "Sequencer"
  frmbattle.MMControl1.Command = "Seek"
 frmbattle.MMControl1.Command = "Open"
 frmbattle.MMControl1.Command = "Play"
 Form2.Enabled = False
 End If
 If MyValue = 3 Then
 End If
 If MyValue = 4 Then
 End If
End If
If KeyCode = vbKeyDown Then
fighter.Picture = fighterdown.Picture  'go down
fighter.Top = fighter.Top + 480
Randomize
   MyValue = Int((4 * Rnd) + 1)
 If MyValue = 1 Then
 End If
 If MyValue = 2 Then
 frmbattle.Show
 frmbattle.goodspeed.Enabled = True
  frmbattle.goodspeed.Enabled = True
 Form2.MMControl1.Command = "Stop"
 frmbattle.MMControl1.Command = "Seek"
  frmbattle.MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
  'MMControl1.Command = "Seek"
 frmbattle.MMControl1.DeviceType = "Sequencer"
  frmbattle.MMControl1.Command = "Seek"
 frmbattle.MMControl1.Command = "Open"
 frmbattle.MMControl1.Command = "Play"
 Form2.Enabled = False
 End If
 If MyValue = 3 Then
 End If
 If MyValue = 4 Then
 End If
End If
If KeyCode = vbKeyLeft Then
fighter.Picture = fighterleft.Picture   'go left
fighter.Left = fighter.Left - 480
Randomize
   MyValue = Int((4 * Rnd) + 1)
 If MyValue = 1 Then
 End If
 If MyValue = 2 Then
 frmbattle.Show
 frmbattle.goodspeed.Enabled = True
  frmbattle.goodspeed.Enabled = True
 Form2.MMControl1.Command = "Stop"
 frmbattle.MMControl1.Command = "Seek"
  frmbattle.MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
  'MMControl1.Command = "Seek"
 frmbattle.MMControl1.DeviceType = "Sequencer"
  frmbattle.MMControl1.Command = "Seek"
 frmbattle.MMControl1.Command = "Open"
 frmbattle.MMControl1.Command = "Play"
 Form2.Enabled = False
 End If
 If MyValue = 3 Then
 End If
 If MyValue = 4 Then
 End If
End If
If KeyCode = vbKeyRight Then
fighter.Picture = fighterright.Picture  'go right
fighter.Left = fighter.Left + 480
Randomize
   MyValue = Int((4 * Rnd) + 1)
 If MyValue = 1 Then
 End If
 If MyValue = 2 Then
 frmbattle.Show
 frmbattle.goodspeed.Enabled = True
  frmbattle.goodspeed.Enabled = True
 Form2.MMControl1.Command = "Stop"
 frmbattle.MMControl1.Command = "Seek"
  frmbattle.MMControl1.FileName = App.Path + "\shadowhavenfight.mid"
  'MMControl1.Command = "Seek"
 frmbattle.MMControl1.DeviceType = "Sequencer"
  frmbattle.MMControl1.Command = "Seek"
 frmbattle.MMControl1.Command = "Open"
 frmbattle.MMControl1.Command = "Play"
 Form2.Enabled = False
 End If
 If MyValue = 3 Then
 End If
 If MyValue = 4 Then
 End If
End If
If KeyCode = vbKeyReturn Then   'open up menu
frmmenu.Show
Form2.Enabled = False
'map1.Enabled = False
'map2.Enabled = False
End If
End Sub

Private Sub Form_Load()
If Form1.position.Text = "map1.map" Then 'load map from saved position
Form2.map1.Enabled = True
End If
If Form1.position.Text = "map2.map" Then
Form2.map2.Enabled = True
End If
If Form1.position.Text = "map3.map" Then
Form2.map3.Enabled = True
End If
 frmbattle.MMControl1.Command = "Stop"
 MMControl1.Command = "Seek"
 MMControl1.FileName = App.Path + "\shadowhaven.mid"
  'MMControl1.Command = "Seek"
MMControl1.DeviceType = "Sequencer"
 MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End Sub

Private Sub map1_Timer()
CommonDialog1.FileName = (App.Path & "\map1.map") 'loads map1
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
'Settings for map1
If fighter.Top = g(73).Top And fighter.Left = g(73).Left Then fighter.Left = (g(73).Left + 480) And fighter.Top = (g(73).Top + 480)
If fighter.Top = g(72).Top And fighter.Left = g(72).Left Then fighter.Left = (g(72).Left + 480) And fighter.Top = (g(72).Top + 480)
If fighter.Left > 4500 Then
fighter.Left = 0
map2.Enabled = True
map1.Enabled = False
End If

End Sub

Private Sub map2_Timer()
CommonDialog1.FileName = (App.Path & "\map2.map") 'loads map2
Picture1.Picture = LoadPicture(CommonDialog1.FileName)

If fighter.Left > 4500 Then
fighter.Left = 0
map3.Enabled = True
map2.Enabled = False
End If
If fighter.Left < -2 Then
fighter.Left = 4000
map1.Enabled = True
map2.Enabled = False
End If
End Sub

Private Sub map3_Timer()
CommonDialog1.FileName = (App.Path & "\map3.map") 'loads map3
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
If fighter.Left < -2 Then
fighter.Left = 4000
map2.Enabled = True
map3.Enabled = False
End If
If fighter.Left = g(5).Left And fighter.Top = g(5).Top Then
fighter.Top = g(15).Top And fighter.Left = g(15).Left
frmtalk.speech.Caption = "Figardo: Hello, the Wind King to the East is waiting for you. Better hurry..."
frmtalk.Show
Form2.Enabled = False
End If
End Sub
Private Sub music_Timer()
 frmbattle.MMControl1.Command = "Stop"
 MMControl1.Command = "Seek"
 MMControl1.FileName = App.Path + "\shadowhaven.mid"
  'MMControl1.Command = "Seek"
MMControl1.DeviceType = "Sequencer"
 MMControl1.Command = "Seek"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End Sub
