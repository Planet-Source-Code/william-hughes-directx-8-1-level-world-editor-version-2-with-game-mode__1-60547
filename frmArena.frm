VERSION 5.00
Begin VB.Form frmArena 
   Caption         =   "Arena Scripting"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEffects 
      Height          =   6105
      ItemData        =   "frmArena.frx":08CA
      Left            =   0
      List            =   "frmArena.frx":08E3
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Frame frmWind 
      Caption         =   "Wind Properties"
      Height          =   6135
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.ListBox lstWind 
         Height          =   1035
         ItemData        =   "frmArena.frx":0918
         Left            =   120
         List            =   "frmArena.frx":0928
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
