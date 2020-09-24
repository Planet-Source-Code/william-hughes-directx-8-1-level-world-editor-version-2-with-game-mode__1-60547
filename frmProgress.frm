VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Progress(perc As Double)
On Error GoTo eih:
Dim test As String
test = perc
    If Left(perc, 3) = "100" Then
        ProgressBar.Value = 100
        Unload Me
    End If
    
    Me.Caption = "Progress " & perc & "%"
    ProgressBar.Value = perc * 10
    
eih:
End Sub




