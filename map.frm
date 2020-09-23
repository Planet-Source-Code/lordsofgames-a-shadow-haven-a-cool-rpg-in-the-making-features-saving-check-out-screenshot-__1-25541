VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Creator for Shadow Haven RPG--By Darwin Yu"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "map.frx":08CA
   ScaleHeight     =   6225
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   36
      Left            =   8160
      Picture         =   "map.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   240
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   35
      Left            =   8160
      Picture         =   "map.frx":125E
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   34
      Left            =   8160
      Picture         =   "map.frx":1AA0
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   33
      Left            =   8160
      Picture         =   "map.frx":22E2
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1680
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   32
      Left            =   8160
      Picture         =   "map.frx":2B24
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   31
      Left            =   8160
      Picture         =   "map.frx":3366
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2640
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   30
      Left            =   8160
      Picture         =   "map.frx":3BA8
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LoadMap"
      Height          =   375
      Left            =   5640
      TabIndex        =   42
      Top             =   4920
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   41
      Top             =   360
      Width           =   4815
      Begin VB.Image g 
         Height          =   480
         Index           =   99
         Left            =   4320
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   98
         Left            =   3840
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   97
         Left            =   3360
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   96
         Left            =   2880
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   95
         Left            =   2400
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   94
         Left            =   1920
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   93
         Left            =   1440
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   92
         Left            =   960
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   91
         Left            =   480
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   90
         Left            =   0
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   89
         Left            =   4320
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   88
         Left            =   3840
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   87
         Left            =   3360
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   86
         Left            =   2880
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   85
         Left            =   2400
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   84
         Left            =   1920
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   83
         Left            =   1440
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   82
         Left            =   960
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   81
         Left            =   480
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   80
         Left            =   0
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   79
         Left            =   4320
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   78
         Left            =   3840
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   77
         Left            =   3360
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   76
         Left            =   2880
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   75
         Left            =   2400
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   74
         Left            =   1920
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   73
         Left            =   1440
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   72
         Left            =   960
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   71
         Left            =   480
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   70
         Left            =   0
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   69
         Left            =   4320
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   68
         Left            =   3840
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   67
         Left            =   3360
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   66
         Left            =   2880
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   65
         Left            =   2400
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   64
         Left            =   1920
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   63
         Left            =   1440
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   62
         Left            =   960
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   61
         Left            =   480
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   60
         Left            =   0
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   59
         Left            =   4320
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   58
         Left            =   3840
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   57
         Left            =   3360
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   56
         Left            =   2880
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   55
         Left            =   2400
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   54
         Left            =   1920
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   53
         Left            =   1440
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   52
         Left            =   960
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   51
         Left            =   480
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   50
         Left            =   0
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   49
         Left            =   4320
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   48
         Left            =   3840
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   47
         Left            =   3360
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   46
         Left            =   2880
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   45
         Left            =   2400
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   44
         Left            =   1920
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   43
         Left            =   1440
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   42
         Left            =   960
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   41
         Left            =   480
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   40
         Left            =   0
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   39
         Left            =   4320
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   38
         Left            =   3840
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   37
         Left            =   3360
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   36
         Left            =   2880
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   35
         Left            =   2400
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   34
         Left            =   1920
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   33
         Left            =   1440
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   32
         Left            =   960
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   31
         Left            =   480
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image g 
         Height          =   480
         Index           =   30
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
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   29
      Left            =   5760
      Picture         =   "map.frx":43EA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2640
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   28
      Left            =   7680
      Picture         =   "map.frx":4C2C
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   27
      Left            =   7200
      Picture         =   "map.frx":546E
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2640
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   26
      Left            =   6720
      Picture         =   "map.frx":5CB0
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2640
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   25
      Left            =   6240
      Picture         =   "map.frx":64F2
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2640
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8160
      TabIndex        =   30
      Top             =   4920
      Width           =   1095
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   24
      Left            =   6240
      Picture         =   "map.frx":6D34
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   23
      Left            =   6720
      Picture         =   "map.frx":7576
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   22
      Left            =   7200
      Picture         =   "map.frx":7DB8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   20
      Left            =   7680
      Picture         =   "map.frx":85FA
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   21
      Left            =   5760
      Picture         =   "map.frx":8E3C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   19
      Left            =   6720
      Picture         =   "map.frx":967E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1680
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   18
      Left            =   6240
      Picture         =   "map.frx":9F48
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   17
      Left            =   5760
      Picture         =   "map.frx":A812
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1680
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   16
      Left            =   7200
      Picture         =   "map.frx":B0DC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5760
      Picture         =   "map.frx":B9A6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   15
      Left            =   7680
      Picture         =   "map.frx":C1E8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   14
      Left            =   7680
      Picture         =   "map.frx":CAB2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   13
      Left            =   7680
      Picture         =   "map.frx":D37C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   12
      Left            =   7680
      Picture         =   "map.frx":D7BE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   11
      Left            =   7200
      Picture         =   "map.frx":DC00
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   10
      Left            =   6720
      Picture         =   "map.frx":E042
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   9
      Left            =   6240
      Picture         =   "map.frx":E884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   8
      Left            =   5760
      Picture         =   "map.frx":F0C6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   7
      Left            =   7200
      Picture         =   "map.frx":F908
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   6
      Left            =   6720
      Picture         =   "map.frx":1014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "map.frx":1098C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   4
      Left            =   5760
      Picture         =   "map.frx":111CE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   3
      Left            =   7200
      Picture         =   "map.frx":11A10
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   2
      Left            =   6720
      Picture         =   "map.frx":12252
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   1
      Left            =   6240
      Picture         =   "map.frx":12A94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.OptionButton p 
      Height          =   480
      Index           =   0
      Left            =   5760
      Picture         =   "map.frx":132D6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   6000
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   480
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reference:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   6120
      TabIndex        =   34
      ToolTipText     =   "X coordinate"
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
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
      Index           =   1
      Left            =   3480
      TabIndex        =   33
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   32
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow Haven Map Maker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   31
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label4 
      Caption         =   "Draw your Map then you can Screen Print it and save as .map file if you want to."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   29
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Quick Draw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6720
      TabIndex        =   28
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      Height          =   3135
      Left            =   5520
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      ToolTipText     =   "X coordinate"
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      ToolTipText     =   "Y coordinate"
      Top             =   5520
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim X As Integer


'I made this program basically for me to make maps.
'It's not intented for players or anyone else but the people
'who program this thing.
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
MsgBox "HELP:" & vbNewLine & "Thanks for using this program." & vbNewLine & "First, to start off, choose your land type from the options box." & vbNewLine & "You can quickly recognize it. It's on the far right." & vbNewLine & "Make sure you don't leave any emptys, or the 32 x 32 sign." & vbNewLine & "Saving comes later once I figure out how to save all 100 images into one bitmap." & vbNewLine & "" & vbNewLine & "The bottom of the screen is a handy tool. The coordinates help you identify where you're placing the land types." & vbNewLine & "The Reference box will help you a lot. When you're editing the game, you'll have to specify which image when yo step on, will bounce you off, like a wall, or an enemy." & vbNewLine & "That'll help you identify that. Don't worry. You'll figure it out." & vbNewLine & "Well, that's a wrap. Hope you'll like this." & vbNewLine & "_____________________________" & vbNewLine & "Copyright 2001 by Darwin Yu. "
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim mapfile As String
CommonDialog1.CancelError = False
CommonDialog1.ShowOpen
mapfile = CommonDialog1.FileName
If Right(mapfile, 3) = "map" Then
Picture2.Picture = LoadPicture(CommonDialog1.FileName)
End If
End Sub

Private Sub Form_Load()
MsgBox "Hello!" & vbNewLine & "Welcome to Shadow Haven RPG's Map maker. This ain't much, considering it's not blt by blt." & vbNewLine & "But I made it just to I don't have to draw everything in paint. It's the more convinient way to design maps." & vbNewLine & "This is fairly simple. All you do is draw your map and screenprint and save as a map file." & vbNewLine & "I couldn't get it to save all the images in one file so we'll just have to make do." & vbNewLine & "Have fun!"
End Sub

Private Sub g_Click(Index As Integer)
g(Index).Picture = Picture1.Picture


End Sub

Private Sub g_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'Dim Num As Integer
Label1.Caption = (g(Index).Top / 480)
Label2.Caption = (g(Index).Left / 480)
With g
Label6.Caption = "Reference: " & "g" & "(" & g(Index).Index & ")"
End With
If Label3.ForeColor = vbYellow Then
g(Index).Picture = Picture1.Picture
End If
End Sub

Private Sub Label3_Click()
If Label3.BorderStyle = 1 - fixedsingle Then
Label3.BorderStyle = 0 - none
Label3.ForeColor = vbRed
Else
Label3.BorderStyle = 1 - fixedsingle
Label3.ForeColor = vbYellow
End If
End Sub

Private Sub p_Click(Index As Integer)
Picture1.Picture = p(Index).Picture
End Sub

