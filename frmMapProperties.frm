VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1320
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "^"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtMusic 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmbCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmbOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtMapDescr 
      Height          =   2055
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtMapPass 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtMapName 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Map Music:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   975
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbCancel_Click()
Unload Me
End Sub

Private Sub cmbOk_Click()
Save = False

MapTyper.Name = txtMapName
MapTyper.pass = txtMapPass
MapTyper.music = txtMusic
MapTyper.description = txtMapDescr

Unload Me
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
CommonDialog.ShowOpen
txtMusic = CommonDialog.Filename
End Sub

Private Sub Form_Load()
CommonDialog.Filter = "Music Files (*.mp3;*.midi;*.mid)|*.mp3;*.midi;*.mid;|MP3 Files (*.mp3)|*.mp3|Midi Files (*.midi;*.mid)|*.midi;*.mid"
txtMapName.Text = MapTyper.Name
txtMapPass.Text = MapTyper.pass
txtMusic.Text = MapTyper.music
txtMapDescr.Text = MapTyper.description


End Sub


