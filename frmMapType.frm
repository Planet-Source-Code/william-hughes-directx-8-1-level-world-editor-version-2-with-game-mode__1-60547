VERSION 5.00
Begin VB.Form frmMapType 
   Caption         =   "Type Of Map"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   Icon            =   "frmMapType.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optMapType 
      Caption         =   "Animation Scene"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton optMapType 
      Caption         =   "Battle Scene"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optMapType 
      Caption         =   "MMORPG"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose the type of map you wish to create."
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
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "frmMapType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

If TheMapType = 3 Then
    MsgBox "Scene Animations are not available yet.", vbCritical, "Map Select Error"
Else
    frmMain.MapType (TheMapType)
    Unload Me
End If
    
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Option1_Click()

End Sub

Private Sub optMapType_Click(Index As Integer)
TheMapType = Index
End Sub


