VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmenu 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   LinkTopic       =   "Form3"
   Picture         =   "frmmenu.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   1050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab menu 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "NewUser"
      TabPicture(0)   =   "frmmenu.frx":0542
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   5  'Downward Diagonal
         Height          =   615
         Left            =   120
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Quit"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Save"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Stats"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDown Then 'All the stuff for the menu
'If Shape1.Top = Label5.Top Then
'Shape1.Top = Label1.Top
'If Shape1.Top = Label5.Top Then
'Else
'Shape1.Top = Shape1.Top + 600
'End If
'End If
'If KeyCode = vbKeyUp Then
'If Shape1.Top = Label1.Top Then
'Shape1.Top = Label5.Top
'Else
'Shape1.Top = Shape1.Top - 600
'End If
'End If
'If KeyCode = vbKeyReturn Then
'If Shape1.Top = Label1.Top Then
'MsgBox ("Stats:" & vbNewLine & "Name: " & Form1.name2.Text & vbNewLine & "HP: " & Form1.goodhp.Text & vbNewLine & "Attack: " & Form1.attack.Text & vbNewLine & "Defense: " & Form1.defense.Text & vbNewLine & "Mana Level: " & Form1.mana.Text & vbNewLine & "Weapon: " & Form1.weapon.Text)
'End If
'If Shape1.Top = Label2.Top Then
'If MsgBox("Are you sure you want to start a New Game?", vbYesNo) = vbYes Then
'Form1.name2.Text = "NewUser"'
'Form1.goodhp.Text = "25'"
'Form1.attack.Text = "5"
'Form1.defense.Text = "7"
'Form1.mana.Text = "2"
'Form1.weapon.Text = "Knife"
'Form1.position.Text = "map1.map"
'Form1.name2.SaveFile (App.Path & "\char.dat")
'Form1.goodhp.SaveFile (App.Path & "\charhp.dat")
'Form1.attack.SaveFile (App.Path & "\charattack.dat")
'Form1.defense.SaveFile (App.Path & "\chardefense.dat")
'Form1.mana.SaveFile (App.Path & "\charmana.dat")
'Form1.weapon.SaveFile (App.Path & "\charweapon.dat")
'Form1.position.SaveFile (App.Path & "\charposition.dat")
'MsgBox "Your stats have been reset."
'Unload Me
'Load Form1
'Form1.Show
'End If
'End If
'If Shape1.Top = Label3.Top Then
'Form1.name2.SaveFile (App.Path & "\char.dat")
'Form1.goodhp.SaveFile (App.Path & "\charhp.dat")
'Form1.attack.SaveFile (App.Path & "\charattack.dat")
'Form1.defense.SaveFile (App.Path & "\chardefense.dat")
'Form1.mana.SaveFile (App.Path & "\charmana.dat")
'Form1.weapon.SaveFile (App.Path & "\charweapon.dat")
'If map1.Enabled = True Then
'Form1.position.Text = "map1.map"
'Form1.position.SaveFile (App.Path & "\charposition.dat")
'MsgBox "Game Saved."
'End If
'If map2.Enabled = True Then
'Form1.position.Text = "map2.map"
'Form1.position.SaveFile (App.Path & "\charposition.dat")
'MsgBox "Game Saved."
'End If
'If map3.Enabled = True Then
'Form1.position.Text = "map3.map"
'Form1.position.SaveFile (App.Path & "\charposition.dat")
'MsgBox "Game Saved."
'End If
'End If
'If Shape1.Top = Label4.Top Then
'frmmenu.Hide
'Form2.Enabled = True
'End If
'If Shape1.Top = Label5.Top Then
'If MsgBox("Are you sure you want to quit? Your stats won't be saved.", vbYesNo) = vbYes Then
'End
'End If
'End If
'End If
'End If   Jeeze!
End Sub

Private Sub menu_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDown Then 'All the stuff for the menu
If Shape1.Top = Label5.Top Then
Shape1.Top = Label1.Top
Else
Shape1.Top = Shape1.Top + 600
End If
End If
If KeyCode = vbKeyUp Then
If Shape1.Top = Label1.Top Then
Shape1.Top = Label5.Top
Else
Shape1.Top = Shape1.Top - 600
End If
End If
If KeyCode = vbKeyReturn Then 'Stats
If Shape1.Top = Label1.Top Then MsgBox ("Stats:" & vbNewLine & "Name: " & Form1.name2.Text & vbNewLine & "HP: " & Form1.goodhp.Text & vbNewLine & "Attack: " & Form1.attack.Text & vbNewLine & "Defense: " & Form1.defense.Text & vbNewLine & "Mana Level: " & Form1.mana.Text & vbNewLine & "Weapon: " & Form1.weapon.Text)
If Shape1.Top = Label2.Top Then 'Restart
If MsgBox("Are you sure you want to start a New Game?", vbYesNo) = vbYes Then
'Form2.map1.Enabled = False
'Form2.map2.Enabled = False
'Form2.map3.Enabled = False
'Form1.name2.Text = "NewUser"
'Form1.goodhp.Text = "25"
'Form1.attack.Text = "5"
'Form1.defense.Text = "7"
'Form1.mana.Text = "2"
'Form1.weapon.Text = "Knife"
'Form1.position.Text = "map1.map"
'Form1.name2.SaveFile (App.Path & "\char.dat")
'Form1.goodhp.SaveFile (App.Path & "\charhp.dat")
'Form1.attack.SaveFile (App.Path & "\charattack.dat")
'Form1.defense.SaveFile (App.Path & "\chardefense.dat")
'Form1.mana.SaveFile (App.Path & "\charmana.dat")
'Form1.weapon.SaveFile (App.Path & "\charweapon.dat")
'Form1.position.SaveFile (App.Path & "\charposition.dat")
'MsgBox "Your stats have been reset."
'Form2.Enabled = True
'Unload Me
'Load Form1
'Form1.Show
'End If
'End If
'-----------------------
'If MsgBox("Are you sure? This will reset all your stats.", vbYesNo) = vbYes Then
Form2.map1.Enabled = False
Form2.map2.Enabled = False
Form2.map3.Enabled = False
Form1.name2.Text = "NewUser" 'resetting stats
Form1.goodhp.Text = "25"
Form1.attack.Text = "5"
Form1.defense.Text = "7"
Form1.mana.Text = "2"
Form1.weapon.Text = "Knife"
Form1.position.Text = "map1.map"
name2.SaveFile (App.Path & "\char.dat")
goodhp.SaveFile (App.Path & "\charhp.dat")
attack.SaveFile (App.Path & "\charattack.dat")
defense.SaveFile (App.Path & "\chardefense.dat")
mana.SaveFile (App.Path & "\charmana.dat")
weapon.SaveFile (App.Path & "\charweapon.dat")
position.SaveFile (App.Path & "\charposition.dat")
MsgBox "Your stats have been reset."
Unload Form2
Form2.Hide
frmmenu.Hide
Form1.Show
End If
End If
'-------------------------
If Shape1.Top = Label3.Top Then 'Saves game
Form1.name2.SaveFile (App.Path & "\char.dat")
Form1.goodhp.SaveFile (App.Path & "\charhp.dat")
Form1.attack.SaveFile (App.Path & "\charattack.dat")
Form1.defense.SaveFile (App.Path & "\chardefense.dat")
Form1.mana.SaveFile (App.Path & "\charmana.dat")
Form1.weapon.SaveFile (App.Path & "\charweapon.dat")
If Form2.map1.Enabled = True Then
Form1.position.Text = "map1.map"
Form1.position.SaveFile (App.Path & "\charposition.dat")
MsgBox "Game Saved."
If Form2.map2.Enabled = True Then
Form1.position.Text = "map2.map"
Form1.position.SaveFile (App.Path & "\charposition.dat")
MsgBox "Game Saved."
If Form2.map3.Enabled = True Then
Form1.position.Text = "map3.map"
Form1.position.SaveFile (App.Path & "\charposition.dat")
MsgBox "Game Saved."
End If
End If
End If
End If
If Shape1.Top = Label4.Top Then 'Exits menu
Form2.Enabled = True
Form2.Show
frmmenu.Hide
End If
If Shape1.Top = Label5.Top Then 'Quits game
If MsgBox("Are you sure you want to quit? Your stats won't be saved.", vbYesNo) = vbYes Then
End
End If
End If
End If


End Sub
