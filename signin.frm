VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stats"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Character"
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go Fight"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox goodhp 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"signin.frx":0000
   End
   Begin RichTextLib.RichTextBox name2 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MaxLength       =   8
      TextRTF         =   $"signin.frx":00AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox attack 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"signin.frx":015E
   End
   Begin RichTextLib.RichTextBox defense 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"signin.frx":020C
   End
   Begin RichTextLib.RichTextBox mana 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"signin.frx":02BA
   End
   Begin RichTextLib.RichTextBox weapon 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"signin.frx":0368
   End
   Begin RichTextLib.RichTextBox position 
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"signin.frx":0416
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Haven 2"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MANA Lv."
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
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Defense"
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
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your HP:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
name2.SaveFile (App.Path & "\char.dat") 'change name
MsgBox "Name has been changed."
End Sub

Private Sub Command2_Click()
frmmenu.menu.Caption = name2.Text 'go!
Form1.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
MsgBox "Your stats were saved." 'save stats
name2.SaveFile (App.Path & "\char.dat")
goodhp.SaveFile (App.Path & "\charhp.dat")
attack.SaveFile (App.Path & "\charattack.dat")
defense.SaveFile (App.Path & "\chardefense.dat")
mana.SaveFile (App.Path & "\charmana.dat")
weapon.SaveFile (App.Path & "\charweapon.dat")
position.SaveFile (App.Path & "\charposition.dat")
End
End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure? This will reset all your stats.", vbYesNo) = vbYes Then
Form2.map1.Enabled = False
Form2.map2.Enabled = False
Form2.map3.Enabled = False
name2.Text = "NewUser" 'resetting stats
goodhp.Text = "25"
attack.Text = "5"
defense.Text = "7"
mana.Text = "2"
weapon.Text = "Knife"
position.Text = "map1.map"
name2.SaveFile (App.Path & "\char.dat")
goodhp.SaveFile (App.Path & "\charhp.dat")
attack.SaveFile (App.Path & "\charattack.dat")
defense.SaveFile (App.Path & "\chardefense.dat")
mana.SaveFile (App.Path & "\charmana.dat")
weapon.SaveFile (App.Path & "\charweapon.dat")
position.SaveFile (App.Path & "\charposition.dat")
MsgBox "Your stats have been reset."
Unload Me
Form1.Show
End If
End Sub

Private Sub Command5_Click()
MsgBox "Your stats were saved."
name2.SaveFile (App.Path & "\char.dat")
goodhp.SaveFile (App.Path & "\charhp.dat")
attack.SaveFile (App.Path & "\charattack.dat")
defense.SaveFile (App.Path & "\chardefense.dat")
mana.SaveFile (App.Path & "\charmana.dat")
weapon.SaveFile (App.Path & "\charweapon.dat")
Form1.position.Text = Form2.CommonDialog1.FileName
Form1.position.SaveFile (App.Path & "\charposition.dat")
End Sub

Private Sub Form_Load()
name2.LoadFile (App.Path & "\char.dat") 'loads the saved stats
goodhp.LoadFile (App.Path & "\charhp.dat")
attack.LoadFile (App.Path & "\charattack.dat")
defense.LoadFile (App.Path & "\chardefense.dat")
mana.LoadFile (App.Path & "\charmana.dat")
weapon.LoadFile (App.Path & "\charweapon.dat")
position.LoadFile (App.Path & "\charposition.dat")
If position.Text = "map1.map" Then 'opens up map for form2
Form2.map1.Enabled = True
End If
If position.Text = "map2.map" Then
Form2.map2.Enabled = True
End If
If position.Text = "map3.map" Then
Form2.map3.Enabled = True
End If
End Sub

