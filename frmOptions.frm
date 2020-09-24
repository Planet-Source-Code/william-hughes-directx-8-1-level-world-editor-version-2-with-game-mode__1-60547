VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Editor Options"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Sectors"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtYSectors 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtXSectors 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "# Of Y-Sectors to Load:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "# Of X-Sectors to Load:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Options"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Left Click"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton chkLeftClickOverWrite 
         Caption         =   "Overwrite Tiles"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton otpLeftClickDontOverWrite 
         Caption         =   "Dont Overwrite Tiles"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
LoadXSectors = txtXSectors
LoadYSectors = txtYSectors
Call WriteIni("Options", "XSectors", Str(LoadXSectors), App.path & "/NWEditor.ini")
Call WriteIni("Options", "YSectors", Str(LoadYSectors), App.path & "/NWEditor.ini")

Unload Me

End Sub

Private Sub Form_Load()
txtXSectors = LoadXSectors
txtYSectors = LoadYSectors

End Sub


