VERSION 5.00
Begin VB.Form frmAutoFill 
   Caption         =   "Auto Fill"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   Icon            =   "frmAutoFill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOverWrite 
      Caption         =   "Overwrite Tiles (faster)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auto Fill"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbl3 
      Caption         =   "Enter number of tiles in width and height you would like to fill."
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
      TabIndex        =   7
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lbl2 
      Caption         =   "Height"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Lbl1 
      Caption         =   "Width"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   735
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
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmAutoFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'if txtheight and txtwidth are not blank
If Len(txtHeight) > 0 And Len(txtWidth) > 0 Then
    'autofill
    Call frmMain.AutoFill((txtHeight * 32) - 32, (txtWidth * 32) - 32, chkOverWrite.Value)
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

