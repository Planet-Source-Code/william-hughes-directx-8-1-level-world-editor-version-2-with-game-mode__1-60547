VERSION 5.00
Begin VB.Form frmTriggers 
   Caption         =   "Triggers"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   Icon            =   "frmTriggers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmbSetTrigger 
      Caption         =   "Set Trigger"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtTrigger 
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ListBox lstTriggers 
      Height          =   4155
      ItemData        =   "frmTriggers.frx":08CA
      Left            =   0
      List            =   "frmTriggers.frx":08E6
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trigger:"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax:"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmTriggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TriggerStorage
    x1 As Single
    x2 As Single
    x3 As Single
    x4 As Single
    y1 As Single
    y2 As Single
    y3 As Single
    y4 As Single
    layer As Integer
    Texture As String
    trigger As String
End Type

Dim Triggerer As TriggerStorage

Public Sub StoreTrigger(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture As String, layer As Integer)
    Triggerer.x1 = x1
    Triggerer.x2 = x2
    Triggerer.x3 = x3
    Triggerer.x4 = x4
    Triggerer.y1 = y1
    Triggerer.y2 = y2
    Triggerer.y3 = y3
    Triggerer.y4 = y4
    Triggerer.layer = layer
    Triggerer.Texture = Texture
End Sub


Private Sub cmbSetTrigger_Click()
    Call frmMain.PlaceTile(Triggerer.x1, Triggerer.x2, Triggerer.x3, Triggerer.x4, Triggerer.y1, Triggerer.y2, Triggerer.y3, Triggerer.y4, Triggerer.Texture, Triggerer.layer, lstTriggers.list(lstTriggers.ListIndex) & " " & txtTrigger, RGB(ValRed, ValGreen, ValBlue))
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call frmMain.UnsetVars
End Sub


