VERSION 5.00
Begin VB.Form frmAutoGen 
   Caption         =   "Auto Generate"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   Icon            =   "frmAutoGen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOverWrite 
      Caption         =   "OverWrite Tiles (faster)"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ListBox lstMixedTiles 
      Height          =   2205
      ItemData        =   "frmAutoGen.frx":08CA
      Left            =   2280
      List            =   "frmAutoGen.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auto Generate"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Type Of Tiles To Mix:"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Lbl1 
      Caption         =   "Width"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lbl2 
      Caption         =   "Height"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lbl3 
      Caption         =   "Enter number of tiles in width and height you would like to place."
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmAutoGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'if txtheight and txtwidth are not blank
If Len(txtWidth) > 0 And Len(txtHeight) > 0 Then
    If lstMixedTiles.list(lstMixedTiles.ListIndex) <> "" Then
        Call frmMain.AutoGenerate(lstMixedTiles.list(lstMixedTiles.ListIndex), (txtHeight - 1) * 32, (txtWidth - 1) * 32, chkOverWrite.Value)
    Else
        MsgBox "No Tile Set choosen", vbCritical, "Auto Gen Error"
    End If
Else
    MsgBox "No Width and Height choosen", vbCritical, "Auto Gen Error"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call ListLoad(lstMixedTiles, App.path & "\mixed.lst")
End Sub


