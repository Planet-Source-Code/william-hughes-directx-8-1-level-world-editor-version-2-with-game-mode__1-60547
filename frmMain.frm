VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Naruto Web World Editor www.narutoweb.net"
   ClientHeight    =   8460
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6840
      Top             =   120
   End
   Begin VB.CheckBox chkGame 
      Caption         =   "Game Mode"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   5880
      Width           =   3135
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8085
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option5 
      Height          =   255
      Left            =   12000
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Height          =   255
      Left            =   9960
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "NarutoWeb Map File (*.NWM)|*.NWM"
   End
   Begin VB.PictureBox PicSizer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CheckBox ChkLayer5 
      Height          =   255
      Left            =   12000
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox ChkLayer4 
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox ChkLayer3 
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox ChkLayer2 
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox ChkLayer1 
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicTilePreview 
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   6240
      Width           =   3135
   End
   Begin VB.PictureBox PicLevel 
      BackColor       =   &H00C00000&
      Height          =   7200
      Left            =   3480
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   600
      Width           =   7200
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   3480
      Max             =   32736
      SmallChange     =   32
      TabIndex        =   2
      Top             =   7800
      Width           =   7215
   End
   Begin VB.VScrollBar VScroll 
      Height          =   8055
      Left            =   3120
      Max             =   32736
      SmallChange     =   32
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      _Version        =   393217
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLoadTiles 
         Caption         =   "Load Tiles"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOView 
      Caption         =   "View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "Map"
      Begin VB.Menu mnuMapProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuMapAutoGen 
         Caption         =   "Auto Generate"
         Begin VB.Menu mnuMapAutoGenScreen 
            Caption         =   "Screen"
         End
         Begin VB.Menu mnuMapAutoGenCustom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu mnuMapAutoFill 
         Caption         =   "Auto Fill"
         Begin VB.Menu mnuMapAutoFillScreen 
            Caption         =   " Screen"
         End
         Begin VB.Menu mnuMapAutoFillCustom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu mnuMapClearLayer 
         Caption         =   "Clear Layer"
         Begin VB.Menu mnuMapClearLayerLayer 
            Caption         =   "All Layers"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuScripting 
      Caption         =   "Scripting"
      Visible         =   0   'False
      Begin VB.Menu mnuScriptingScript 
         Caption         =   "View Script"
      End
      Begin VB.Menu mnuScriptingArena 
         Caption         =   "Arena"
      End
      Begin VB.Menu mnuScriptingDetails 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu mnuNPCS 
      Caption         =   "NPCS"
      Visible         =   0   'False
      Begin VB.Menu mnuNPCSInsertNPC 
         Caption         =   "Insert NPC"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   " Colors"
      Begin VB.Menu mnuColorsResetColors 
         Caption         =   "Reset Colors"
      End
      Begin VB.Menu mnuChangeColors 
         Caption         =   "Change Colors"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Read"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditTrigger 
         Caption         =   "Edit Trigger"
      End
      Begin VB.Menu mnuEditNPC 
         Caption         =   "Edit NPC"
      End
      Begin VB.Menu mnuEditMonster 
         Caption         =   "Edit Monster"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################
'
'           Title: NarutoWeb Level Editor
'           Desc: Level Editor For NarutoWeb (Online Naruto MMORPG)
'           Written by: William Hughes
'           Started: March 24th 2004
'           Contact: Sim@po2.net
'           Website: www.narutoweb.net | www.po2.net
'
'###############################
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type Queue
    word As String
    Owner As String
End Type

Dim Nqueues As Integer
Dim Queues() As Queue
Dim Dx As DirectX8 'The master Object, everything comes from here
Dim D3D As Direct3D8 'This controls all things 3D
Dim D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Dim bRunning As Boolean 'Controls whether the program is running or not...

'This is the Flexible-Vertex-Format description for a 2D vertex (Transformed and Lit)
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'This structure describes a transformed and lit vertex - it's identical to the DirectX7 type "D3DTLVERTEX"
Private Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Private Type Objects_
    X As Single
    Y As Single
    layer As Single
    trigger As String
    TextureRefNum As Long
    TileRefNum As Long
    color As Long
    XSector As Long
    YSector As Long
    RenderOrder As Long
End Type

Private Type TileStorage
    X As Single
    Y As Single
    trigger As String
    Objects() As Objects_
    TileRefNum As Long
    TextureRefNum As Long
    Walk As Boolean
    Used As Boolean
    color As Long
    XSector As Long
    YSector As Long
End Type

Private Type NPCChar
    Name As String
    Dialogue1 As String
    Dialogue2 As String
    Dialogue3 As String
    Dialogue4 As String
    Dialogue5 As String
    Say1 As String
    Say2 As String
    Say3 As String
    Say4 As String
    Say5 As String
    Filename As String
End Type

Private Type MonsterChar
    Name As String
    HP As Integer
    MP As Integer
    ATK As Integer
    Def As Integer
    Spd As Integer
    Accuracy As Integer
    Evade As Integer
    DropChance As Integer
    SpawnRate As Integer
    Item As Integer
    Value As Integer
    Exp As Integer
    Filename As String
End Type

Dim Monster() As MonsterChar
Dim NPC() As NPCChar
Dim NPCCounter As Long
Dim MonsterCounter As Long

Dim Selected As Boolean

Dim TilesLayer1() As TLVERTEX
Dim TilesLayer2() As TLVERTEX
Dim TilesLayer3() As TLVERTEX
Dim TilesLayer4() As TLVERTEX
Dim TilesLayer5() As TLVERTEX

Dim CurrentTile(0 To 3) As TLVERTEX 'square that follows the mouse
Dim LayerTiles1() As TileStorage
Dim LayerTiles2() As TileStorage
Dim LayerTiles3() As TileStorage
Dim LayerTiles4() As TileStorage
Dim LayerTiles5() As TileStorage

Dim Map(0 To 1039, 0 To 1039) As TileStorage
'array for sorting left to right
Dim lArray() As Long

'Counter for # of tiles depending on layer
Dim TileCounter1 As Long
Dim TileCounter2 As Long
Dim TileCounter3 As Long
Dim TileCounter4 As Long
Dim TileCounter5 As Long

'Counter for # of textures depending on layer
Dim TextureCounter As Long
Dim TextureCounter1 As Long
Dim TextureCounter2 As Long
Dim TextureCounter3 As Long
Dim TextureCounter4 As Long
Dim TextureCounter5 As Long

'x and y pixel offset
Dim XPixelDiff As Single
Dim YPixelDiff As Single

Dim MouseIsDown As Boolean
'TEXTURING STUFF
Dim D3DX As D3DX8 '//A helper library

'Store textures by layer
Dim Textures() As Direct3DTexture8
Dim Textures1() As Direct3DTexture8
Dim Textures2() As Direct3DTexture8
Dim Textures3() As Direct3DTexture8
Dim Textures4() As Direct3DTexture8
Dim Textures5() As Direct3DTexture8

'Name of Texture by layer
Dim TheTextures() As String
Dim TheTextures1() As String
Dim TheTextures2() As String
Dim TheTextures3() As String
Dim TheTextures4() As String
Dim TheTextures5() As String

Dim CurrentTexture As Direct3DTexture8
Dim ColorKeyVal As Long 'Transparency color holder

Dim Rendered As Long

'keys for actions
Dim SetTrigger As Boolean
Dim SetNPC As Boolean
Dim SetMonster As Boolean
Dim SetColor As Boolean
Dim Layer1Move As Boolean
Dim Layer2Move As Boolean
Dim DoDelete As Boolean
Dim DoEdit As Boolean
Dim Locked As Boolean

'sectors
Dim SectorXOffset As Long
Dim SectorYOffset As Long
Dim MaxXSector As Long
Dim MaxYSector As Long

'matrix pos
Dim MaxXMatrixPos As Long
Dim MaxYMatrixPos As Long

'Keep track of how long users have worked with the editor
Dim OverallHours As Integer
Dim OverallMinutes As Integer
Dim SessionHours As Integer
Dim SessionMinutes As Integer



'Game Mode Stuff
'Game Mode Varaibles
Dim CharTexture(12) As Direct3DTexture8
Dim CharVertex(0 To 3) As TLVERTEX 'char

Dim ScrollTexture As Direct3DTexture8
Dim ScrollVertex(0 To 3) As TLVERTEX

'character tile position
Dim CharXPos As Single
Dim CharYPos As Single
Dim CharXPixelPos As Long
Dim CharYPixelPos As Long

'character image
Dim CharImage As Integer
Dim CharStep As Integer

'rendering stuff
Dim RenderQueues() As TLVERTEX
Dim RenderQueueCount As Long

'font stuff
Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont

Dim GMode As Integer 'what the game is doing. walking, displaying NPC. ect.

Public Sub AutoFill(a As Integer, b As Integer, OverWrite As Integer)
Dim i As Single
Dim z As Single
Dim Counter As Long
Dim MaxNum1 As Long
Dim rndnum As Integer
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single

'total tiles being placed
MaxNum1 = (a / 32) * (b / 32)
Counter = 0
Unload frmAutoFill
Load frmProgress
frmProgress.Show , frmMain
On Error GoTo eih

'loop threw top left to bottom right of level shown.
For i = YPixelDiff To a + YPixelDiff Step 32
    For z = XPixelDiff To b + XPixelDiff Step 32

        'randomize number
        rndnum = Int((70 - 44 + 1) * Rnd) + 44
        'store in hidden picture box to get size
        PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Tree.SelectedItem.FullPath)
    
    x1 = z
    y1 = i

    x2 = z + 32
    y2 = i
    
    x3 = z
    y3 = i + 32

    x4 = z + 32
    y4 = i + 32


    
    'Get Top Layer we will be working with
    If Option5.Value = True Then
        'dont overwrite tile
        If OverWrite = 0 Then
            'place tile
            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, "", RGB(ValRed, ValGreen, ValBlue))
        'place without overwrite check
        Else
            'place tile
            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, "", RGB(ValRed, ValGreen, ValBlue))
        End If
    ElseIf Option4.Value = True Then
        'dont overwrite tile
        If OverWrite = 0 Then
            'place tile
            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, "", RGB(ValRed, ValGreen, ValBlue))
        'place without overwrite check
        Else
            'place tile
            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, "", RGB(ValRed, ValGreen, ValBlue))
        End If
    ElseIf Option3.Value = True Then
        'dont overwrite tile
        If OverWrite = 0 Then
            'place tile
            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, "", RGB(ValRed, ValGreen, ValBlue))
        'place without overwrite check
        Else
            'place tile
            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, "", RGB(ValRed, ValGreen, ValBlue))
        End If
    ElseIf Option2.Value = True Then
        'dont overwrite tile
        If OverWrite = 0 Then
            'place tile
            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, "", RGB(ValRed, ValGreen, ValBlue))
        'place without overwrite check
        Else
            'place tile
            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, "", RGB(ValRed, ValGreen, ValBlue))
        End If
    ElseIf Option1.Value = True Then
        'dont overwrite tile
        If OverWrite = 0 Then
            'place tile
            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, "", RGB(ValRed, ValGreen, ValBlue))
        'place without overwrite check
        Else
            'place tile
            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, "", RGB(ValRed, ValGreen, ValBlue))
        End If
    End If
        'increase counter
        Counter = Counter + 1
        DoEvents
        
        

        frmProgress.Progress ((Counter / MaxNum1) * 100)

    Next z
Next i

eih:
End Sub

Public Sub AutoGenerate(Name As String, b As Integer, c As Integer, OverWrite As Integer)
Dim a As Integer
Dim i As Single
Dim z As Single
Dim MaxNum1 As Long
Dim newtile As String
Dim rndnum As Integer
Dim Temp1 As Integer
Dim Temp2 As Integer
Dim temp3 As Integer
Dim temp4 As Integer
Dim temp5 As Integer
Dim temp6 As Integer
Dim temp7 As Integer
Dim temp8 As Integer
Dim temp9 As Integer
Dim Counter As Long
Dim Counter1 As Long
Dim Counter2 As Long
Dim TheTile(1000, 1000) As String
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single

Unload frmAutoGen
Load frmProgress
frmProgress.Show , frmMain
Counter1 = 0
Counter = 0
MaxNum1 = (b / 32) * (c / 32)

'On Error GoTo eih
'loop threw top left to bottom right of level shown.
For i = YPixelDiff To b + YPixelDiff Step 32
Counter2 = 0
Counter1 = Counter1 + 1
    For z = XPixelDiff To c + XPixelDiff Step 32
Counter2 = Counter2 + 1

x1 = z
y1 = i

x2 = z + 32
y2 = i
    
x3 = z
y3 = i + 32

x4 = z + 32
y4 = i + 32
    
        'randomize number
        If i = YPixelDiff And z = XPixelDiff Then
            newtile = ""
            For a = 1 To 9
                Randomize
                rndnum = Int((1 - 0 + 1) * Rnd) + 0
                newtile = newtile & rndnum
                TheTile(Counter1, Counter2) = newtile
            Next a
            

        
                'store in hidden picture box to get size
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Name & "\" & newtile & ".bmp")
                

                'get top layer were working with
                If Option5.Value = True Then
                        'dont overwrite tile
                        If OverWrite = 0 Then
                            'place tile
                            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
                        'place without overwrite check
                        Else
                            'place tile
                            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
                        End If
                ElseIf Option4.Value = True Then
                        'dont overwrite tile
                        If OverWrite = 0 Then
                            'place tile
                            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
                        'place without overwrite check
                        Else
                            'place tile
                            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
                        End If
                ElseIf Option3.Value = True Then
                        'dont overwrite tile
                        If OverWrite = 0 Then
                            'place tile
                            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
                        'place without overwrite check
                        Else
                            'place tile
                            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
                        End If
                ElseIf Option2.Value = True Then
                        'dont overwrite tile
                        If OverWrite = 0 Then
                            'place tile
                            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
                        'place without overwrite check
                        Else
                            'place tile
                            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
                        End If
                ElseIf Option1.Value = True Then
                        'dont overwrite tile
                        If OverWrite = 0 Then
                            'place tile
                            Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
                        'place without overwrite check
                        Else
                            'place tile
                            Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
                        End If
                End If
        End If
        
        If Counter1 = 1 And Counter2 > 1 Then
        temp3 = Mid(TheTile(Counter1, Counter2 - 1), 3, 1)
        temp6 = Mid(TheTile(Counter1, Counter2 - 1), 6, 1)
        temp9 = Right(TheTile(Counter1, Counter2 - 1), 1)


        newtile = ""
        
        
        newtile = newtile & temp3
                
                
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum
        
        
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum

        newtile = newtile & temp6
        
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum
        
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum
  
        newtile = newtile & temp9
        
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum
        
        Randomize
        rndnum = Int((1 - 0 + 1) * Rnd) + 0
        newtile = newtile & rndnum
        
        TheTile(Counter1, Counter2) = newtile
        
        
        'store in hidden picture box to get size
        PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Name & "\" & newtile & ".bmp")
                
        'Get Top Layer we will be working with
        If Option5.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
            'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option4.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option3.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option2.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
            'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option1.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        End If
        
        
        End If
        
        If Counter1 > 1 Then
            If z = XPixelDiff Then

            temp7 = Mid(TheTile(Counter1 - 1, Counter2), 7, 1)
            temp8 = Mid(TheTile(Counter1 - 1, Counter2), 8, 1)
            temp9 = Right(TheTile(Counter1 - 1, Counter2), 1)


            newtile = ""

        
            newtile = newtile & temp7
            newtile = newtile & temp8
            newtile = newtile & temp9
            
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
        
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum

        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
  
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
        
            TheTile(Counter1, Counter2) = newtile
        
        
            'store in hidden picture box to get size
            PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Name & "\" & newtile & ".bmp")
                
            'place tile
        'Get Top Layer we will be working with
        If Option5.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option4.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option3.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option2.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        ElseIf Option1.Value = True Then
            'dont overwrite tile
            If OverWrite = 0 Then
                'place tile
                Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
            'place without overwrite check
            Else
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
            End If
        End If
        
            Else

            temp3 = Mid(TheTile(Counter1, Counter2 - 1), 3, 1)
            temp6 = Mid(TheTile(Counter1, Counter2 - 1), 6, 1)
            temp8 = Mid(TheTile(Counter1 - 1, Counter2), 8, 1)
            temp9 = Right(TheTile(Counter1 - 1, Counter2), 1)


            newtile = ""

        
            newtile = newtile & temp3
            newtile = newtile & temp8
            newtile = newtile & temp9
            newtile = newtile & temp6
            

        
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum

        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum

            temp9 = Right(TheTile(Counter1, Counter2 - 1), 1)
            newtile = newtile & temp9
            
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
  
        
            Randomize
            rndnum = Int((1 - 0 + 1) * Rnd) + 0
            newtile = newtile & rndnum
        
        
            TheTile(Counter1, Counter2) = newtile
        
        
            'store in hidden picture box to get size
            PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Name & "\" & newtile & ".bmp")
                
            'Get Top Layer we will be working with
            If Option5.Value = True Then
                'dont overwrite tile
                If OverWrite = 0 Then
                    'place tile
                    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
                'place without overwrite check
                Else
                    'place tile
                    Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 5, "", RGB(ValRed, ValGreen, ValBlue))
                End If
            ElseIf Option4.Value = True Then
                'dont overwrite tile
                If OverWrite = 0 Then
                    'place tile
                    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
                'place without overwrite check
                Else
                    'place tile
                    Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 4, "", RGB(ValRed, ValGreen, ValBlue))
                End If
            ElseIf Option3.Value = True Then
                'dont overwrite tile
                If OverWrite = 0 Then
                    'place tile
                    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
                'place without overwrite check
                Else
                    'place tile
                    Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 3, "", RGB(ValRed, ValGreen, ValBlue))
                End If
            ElseIf Option2.Value = True Then
                'dont overwrite tile
                If OverWrite = 0 Then
                    'place tile
                    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
                'place without overwrite check
                Else
                    'place tile
                    Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 2, "", RGB(ValRed, ValGreen, ValBlue))
                End If
            ElseIf Option1.Value = True Then
                'dont overwrite tile
                If OverWrite = 0 Then
                    'place tile
                    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
                'place without overwrite check
                Else
                    'place tile
                    Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Name & "\" & newtile & ".bmp", 1, "", RGB(ValRed, ValGreen, ValBlue))
                End If
            End If



            End If
        End If
        'increase counter
        Counter = Counter + 1

        DoEvents
        
        'update progress bar
        frmProgress.Progress ((Counter / MaxNum1) * 100)

    Next z
Next i

eih:
End Sub

Public Sub ClearLayer(layer As Integer)

End Sub

Public Sub GetKeys()
Dim Right As Boolean
Dim Left As Boolean
Dim Up As Boolean
Dim Down As Boolean
Right = GetAsyncKeyState(vbKeyRight)
Left = GetAsyncKeyState(vbKeyLeft)
Up = GetAsyncKeyState(vbKeyUp)
Down = GetAsyncKeyState(vbKeyDown)



'if running in walking mode
If GMode = 1 Then
    
    If Down = True Then
        'if check if tile is walkable
        If TileWalkable(CharXPos, CharYPos + 1) = True Then
            'change pixel diff
            YPixelDiff = YPixelDiff + 32
            'change char matrix pos
            CharYPos = CharYPos + 1
            'change char pixel pos
            CharYPixelPos = CharYPixelPos + 32
            
        If CharImage = 0 Or CharImage = 1 Or CharImage = 2 Or CharImage = 3 Or CharImage = 4 Or CharImage = 5 Or CharImage = 7 Or CharImage = 10 Or CharImage = 11 Or CharImage = 12 Then
            If CharStep = 1 Then
                CharImage = 6
                CharStep = 2
            Else
                CharImage = 8
            End If
        ElseIf CharImage = 6 Then
            CharImage = 7
            CharStep = 2
        ElseIf CharImage = 8 Then
            CharImage = 7
            CharStep = 1
        Else
            CharImage = 7
        End If
        
        End If
    End If

    If Up = True Then
        'if check if tile is walkable
        If TileWalkable(CharXPos, CharYPos - 1) = True Then
            'change pixel diff
            YPixelDiff = YPixelDiff - 32
            'change char matrix pos
            CharYPos = CharYPos - 1
            'change char pixel pos
            CharYPixelPos = CharYPixelPos - 32
            
        If CharImage = 0 Or CharImage = 1 Or CharImage = 2 Or CharImage = 3 Or CharImage = 4 Or CharImage = 5 Or CharImage = 7 Or CharImage = 6 Or CharImage = 8 Or CharImage = 10 Then
            If CharStep = 1 Then
                CharImage = 9
                CharStep = 2
            Else
                CharImage = 11
            End If
        ElseIf CharImage = 9 Then
            CharImage = 10
            CharStep = 2
        ElseIf CharImage = 11 Then
            CharImage = 10
            CharStep = 1
        Else
            CharImage = 10
        End If
        
        End If
    End If
    
    If Right = True Then
        'if check if tile is walkable
        If TileWalkable(CharXPos + 1, CharYPos) = True Then
            'change pixel diff
            XPixelDiff = XPixelDiff + 32
            'change char matrix pos
            CharXPos = CharXPos + 1
            'change char pixel pos
            CharXPixelPos = CharXPixelPos + 32
            
        
        If CharImage = 1 Or CharImage = 3 Or CharImage = 4 Or CharImage = 5 Or CharImage = 6 Or CharImage = 7 Or CharImage = 8 Or CharImage = 9 Or CharImage = 10 Or CharImage = 11 Then
            If CharStep = 1 Then
                CharImage = 2
                CharStep = 2
            Else
                CharImage = 0
            End If
        ElseIf CharImage = 2 Then
            CharImage = 1
            CharStep = 2
        ElseIf CharImage = 0 Then
            CharImage = 1
            CharStep = 1
        Else
            CharImage = 1
        End If
        
        End If
    End If

        
    
    If Left = True Then
        'if check if tile is walkable
        If TileWalkable(CharXPos - 1, CharYPos) = True Then
            'change pixel diff
            XPixelDiff = XPixelDiff - 32
            'change char matrix pos
            CharXPos = CharXPos - 1
            'change char pixel pos
            CharXPixelPos = CharXPixelPos - 32

        If CharImage = 0 Or CharImage = 1 Or CharImage = 2 Or CharImage = 4 Or CharImage = 6 Or CharImage = 7 Or CharImage = 8 Or CharImage = 9 Or CharImage = 10 Or CharImage = 11 Then
            If CharStep = 1 Then
                CharImage = 5
                CharStep = 2
            Else
                CharImage = 3
            End If
        ElseIf CharImage = 5 Then
            CharImage = 4
            CharStep = 2
        ElseIf CharImage = 3 Then
            CharImage = 4
            CharStep = 1
        Else
            CharImage = 4
        End If
        
        End If
    End If

'end if of game mode
End If
    

End Sub

Public Function TileWalkable(X, Y) As Boolean

    If Map(X, Y).Walk = True Then
        TileWalkable = True
    Else
        TileWalkable = False
    End If
    
End Function

Public Sub DeleteTile(X As Single, Y As Single, layer As Integer)
Dim TilePlacer As Long
Dim TileLayer As Integer
Dim GoGood As Boolean
Dim TheX As Single
Dim TheY As Single
Dim i As Long
Dim z As Long



    'because I am rounding the pixels to the nearest 32. there
    'are some problems with certain pixel co-ords rounding

    'if the distance between x's are less then 32. increase X
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32))) < 32 Then
        X = X + 1
    End If

    'if the distance between y's are less then 32. increase Y
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32))) < 32 Then
        Y = Y + 1
    End If


    'if the distance between x's are > 64. decrease x
    '1 pixel reduces the rounding to 32
    If (((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32)) >= 64 Then
        X = X - 1
    End If

    'if the distance between y's are > 64. decrease y
    '1 pixel reduces the rounding to 32
    If (((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32)) >= 64 Then
        Y = Y - 1
    End If

TheX = Round(((X)) / 32, 0)
TheY = Round(((Y)) / 32, 0)

    'if layer is 5
    If layer = 5 Then
    
    'loop to check every stored Vertex.
    'Checks to see if the co-ords exist
    'TileCounter * 4 = Total Vertex Count.
    For i = (TileCounter5 * 4) To 0 Step -4

    'first matching X pixel in use
    If X >= (TilesLayer5(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If X <= (TilesLayer5(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If Y >= (TilesLayer5(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If Y <= (TilesLayer5(i + 2).Y) - YPixelDiff Then
                
                '''''''''''''''Start Delete Tile
                    'Remove Tile.
                    If i = TileCounter5 * 4 Then
                        ReDim Preserve TilesLayer5((TileCounter5 * 4))
                        ReDim Preserve LayerTiles5((TileCounter5 * 4))
                        ReDim Preserve Textures5(TextureCounter5)
                        ReDim Preserve TheTextures5(TextureCounter5)
                        TileCounter5 = TileCounter5 - 1
                        TextureCounter5 = TextureCounter5 - 1
                        Exit For
                    Else
                    

                        For z = i To ((TileCounter5 * 4)) '+ 4
                            If z = TileCounter5 * 4 Then
                                Exit For
                            Else
                                TilesLayer5(z) = TilesLayer5(z + 4)
                                LayerTiles5(z) = LayerTiles5(z + 4)
                            End If
                        Next z


                        For z = (i / 4) To TileCounter5
                            If z = TileCounter5 Then
                                Exit For
                            Else
                                'Set Textures5(z) = D3DX.CreateTextureFromFile(D3DDevice, TheTextures5(z + 1))
                                TheTextures5(z) = TheTextures5(z + 1)
                            End If
                        Next z
                        

                        ReDim Preserve Textures5(TextureCounter5)
                        ReDim Preserve TheTextures5(TextureCounter5)

                        ReDim Preserve TilesLayer5((TileCounter5 * 4))
                        ReDim Preserve LayerTiles5((TileCounter5 * 4))
                        TileCounter5 = TileCounter5 - 1
                        TextureCounter5 = TextureCounter5 - 1

                        Exit Sub

                        End If
                '''''''''''''Finish Remove Tile
                End If
                
            End If
        End If
    End If

    Next i
    
    'if layer is 4
    ElseIf layer = 4 Then
    
    
    'loop to check every stored Vertex.
    'Checks to see if the co-ords exist
    'TileCounter * 4 = Total Vertex Count.
    For i = (TileCounter4 * 4) To 0 Step -4

    'first matching X pixel in use
    If X >= (TilesLayer4(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If X <= (TilesLayer4(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If Y >= (TilesLayer4(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If Y <= (TilesLayer4(i + 2).Y) - YPixelDiff Then
                
                '''''''''''''''Start Delete Tile
                    'Remove Tile.
                    If i = TileCounter4 * 4 Then
                        ReDim Preserve TilesLayer4((TileCounter4 * 4))
                        ReDim Preserve LayerTiles4((TileCounter4 * 4))
                        ReDim Preserve Textures4(TextureCounter4)
                        ReDim Preserve TheTextures4(TextureCounter4)
                        TileCounter4 = TileCounter4 - 1
                        TextureCounter4 = TextureCounter4 - 1
                        Exit For
                    Else
                    

                        For z = i To ((TileCounter4 * 4)) '+ 4
                            If z = TileCounter4 * 4 Then
                                Exit For
                            Else
                                TilesLayer4(z) = TilesLayer4(z + 4)
                                LayerTiles4(z) = LayerTiles4(z + 4)
                            End If
                        Next z


                        For z = (i / 4) To TileCounter4
                            If z = TileCounter4 Then
                                Exit For
                            Else
                                'Set Textures4(z) = D3DX.CreateTextureFromFile(D3DDevice, TheTextures4(z + 1))
                                TheTextures4(z) = TheTextures4(z + 1)
                            End If
                        Next z
                        

                        ReDim Preserve Textures4(TextureCounter4)
                        ReDim Preserve TheTextures4(TextureCounter4)

                        ReDim Preserve TilesLayer4((TileCounter4 * 4))
                        ReDim Preserve LayerTiles4((TileCounter4 * 4))
                        TileCounter4 = TileCounter4 - 1
                        TextureCounter4 = TextureCounter4 - 1

                        Exit Sub

                        End If
                '''''''''''''Finish Remove Tile
                End If
                
            End If
        End If
    End If

    Next i
    
    'If layer is 3
    ElseIf layer = 3 Then
    
    
    'loop to check every stored Vertex.
    'Checks to see if the co-ords exist
    'TileCounter * 4 = Total Vertex Count.
    For i = (TileCounter3 * 4) To 0 Step -4

    'first matching X pixel in use
    If X >= (TilesLayer3(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If X <= (TilesLayer3(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If Y >= (TilesLayer3(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If Y <= (TilesLayer3(i + 2).Y) - YPixelDiff Then
                
                '''''''''''''''Start Delete Tile
                    'Remove Tile.
                    If i = TileCounter3 * 4 Then
                        ReDim Preserve TilesLayer3((TileCounter3 * 4))
                        ReDim Preserve LayerTiles3((TileCounter3 * 4))
                        ReDim Preserve Textures3(TextureCounter3)
                        ReDim Preserve TheTextures3(TextureCounter3)
                        TileCounter3 = TileCounter3 - 1
                        TextureCounter3 = TextureCounter3 - 1
                        Exit For
                    Else
                    

                        For z = i To ((TileCounter3 * 4)) '+ 4
                            If z = TileCounter3 * 4 Then
                                Exit For
                            Else
                                TilesLayer3(z) = TilesLayer3(z + 4)
                                LayerTiles3(z) = LayerTiles3(z + 4)
                            End If
                        Next z


                        For z = (i / 4) To TileCounter3
                            If z = TileCounter3 Then
                                Exit For
                            Else
                                'Set Textures3(z) = D3DX.CreateTextureFromFile(D3DDevice, TheTextures3(z + 1))
                                TheTextures3(z) = TheTextures3(z + 1)
                            End If
                        Next z
                        

                        ReDim Preserve Textures3(TextureCounter3)
                        ReDim Preserve TheTextures3(TextureCounter3)

                        ReDim Preserve TilesLayer3((TileCounter3 * 4))
                        ReDim Preserve LayerTiles3((TileCounter3 * 4))
                        TileCounter3 = TileCounter3 - 1
                        TextureCounter3 = TextureCounter3 - 1

                        Exit Sub

                        End If
                '''''''''''''Finish Remove Tile
                End If
                
            End If
        End If
    End If

    Next i
    
    'if layer is 2
    ElseIf layer = 2 Then



    'Checks to see if tile exist
    If Map(TheX, TheY).Used = True Then
        'set tile as no longer being used
        Map(TheX, TheY).Used = False
        

                
                '''''''''''''''Start Delete Tile
 
                    
                        'loop threw all Tiles starting from deleted tile
                        'then move rest of tiles to fill the deleted tiles
                        'position
                        'For z = i To ((TileCounter2 * 4)) '+ 4
                            'if z = TileCounter2 * 4 Then
                                'Exit For
                            'Else
                                'TilesLayer2(z) = TilesLayer2(z + 4)
                            'End If
                        'Next z

                        'loop threw all Tiles starting from deleted tile
                        'then move rest of tiles to fill the deleted tiles
                        'position
                        'For z = (i / 4) To TileCounter2
                            'If z = TileCounter2 Then
                                'Exit For
                            'Else
                                'Set Textures2(z) = D3DX.CreateTextureFromFile(D3DDevice, TheTextures2(z + 1))
                                'TheTextures2(z) = TheTextures2(z + 1)
                            'End If
                        'Next z
                        
                        'ReDim Preserve Textures2(TextureCounter2)
                        'ReDim Preserve TheTextures2(TextureCounter2)

                        'ReDim Preserve TilesLayer2((TileCounter2 * 4))
                        TileCounter2 = TileCounter2 - 1
                        TextureCounter2 = TextureCounter2 - 1

                        Exit Sub

                        
                '''''''''''''Finish Remove Tile
    End If
    
    'if layer is 1
    ElseIf layer = 1 Then
    
    
    'Checks to see if tile exist
    If Map(TheX, TheY).Used = True Then
        'set tile as no longer being used
        Map(TheX, TheY).Used = False
  
                '''''''''''''''Start Delete Tile
                
    
                        'loop threw all Tiles starting from deleted tile
                        'then move rest of tiles to fill the deleted tiles
                        'position
                        'For z = i To ((TileCounter1 * 4)) '+ 4
                            'if not last tile
                            'If z <> TileCounter1 * 4 Then
                                'move new tile
                                'TilesLayer1(z) = TilesLayer1(z + 4)
                            'End If
                        'Next z

                        'loop threw all Tiles starting from deleted tile
                        'then move rest of tiles to fill the deleted tiles
                        'position
                        'For z = (i / 4) To TileCounter1
                            'if not last tile
                            'If z <> TileCounter1 Then
                                'move texture
                                'TheTextures1(z) = TheTextures1(z + 1)
                            'End If
                        'Next z
                        
                        'reserve by reducing by 1
                        'ReDim Preserve Textures1(TextureCounter1)
                        'ReDim Preserve TheTextures1(TextureCounter1)

                        'ReDim Preserve TilesLayer1((TileCounter1 * 4))
                        TileCounter1 = TileCounter1 - 1
                        TextureCounter1 = TextureCounter1 - 1

                        Exit Sub


                       
                '''''''''''''Finish Delete Tile

    End If

    
    'end of layer checking
    End If
End Sub


'// Initialise : This procedure kick starts the whole process.
'// It'll return true for success, false if there was an error.
'// THIS FUNCTION IS NOT IMPORTANT. WONT BE UPDATED ANYMORE. BASIC DIRECTX
'                       STRUCTURE
Public Function Initialise() As Boolean
On Error GoTo ErrHandler:

Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode
Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport


Set Dx = New DirectX8  '//Create our Master Object
Set D3D = Dx.Direct3DCreate() '//Make our Master Object create the Direct3D Interface
Set D3DX = New D3DX8 '//Create our helper library...

'//We're going to use Fullscreen mode because I prefer it to windowed mode :)

'DispMode.Format = D3DFMT_X8R8G8B8
'DispMode.Format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode '//Retrieve the current display Mode

D3DWindow.Windowed = 1 '//Tell it we're using Windowed Mode
D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC '//We'll refresh when the monitor does
D3DWindow.BackBufferFormat = DispMode.Format '//We'll use the format we just retrieved...

D3DWindow.hDeviceWindow = PicLevel.hWnd

'//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
'//See the lesson text for more information on this line...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, PicLevel.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                        D3DWindow)

'//Set the vertex shader to use our vertex format
D3DDevice.SetVertexShader FVF

'//Transformed and lit vertices dont need lighting
'   so we disable it...
D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

'font stuff
fnt.Name = "Verdana"
fnt.Size = 12
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

'//We can only continue if Initialise Geometry succeeds;
'   If it doesn't we'll fail this call as well...
If InitialiseGeometry() = True Then
    Initialise = True '//We succeeded
    Exit Function
End If


ErrHandler:
'//We failed; for now we wont worry about why.
Debug.Print "Error Number Returned: " & Err.Number
Initialise = False
End Function

Public Sub LoadMap(X As String)
Dim ln As String
Dim Counters As Integer
Dim Texture As String
Dim trigger As String
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single
Dim color As Long
Dim XSector As Long
Dim YSector As Long
Dim TempNum As Integer
Dim TempStr As String
Dim MapType1 As Integer
Dim Name As String
Dim Loadedpass As String
Dim Password As String
Dim music As String
Dim Desc As String
Dim XLong As Long
Dim YLong As Long
Dim TheYLong As Long
Dim TheXLong As Long
Dim TheDir1 As String
'if string not empty
If X <> "" Then
    'load map properties
    Open X For Input As #1
    
        Input #1, MapType1
        Input #1, Name
        Input #1, Loadedpass
        Input #1, music
        Input #1, Desc

    Close #1
    'if password on map not blank
    If Loadedpass <> "" Then
    Password = InputBox("This map is password protected. Enter password", "Password Protected Map")
        'if passwords match
        If LCase(Password) <> LCase(Loadedpass) Then
            MsgBox "Incorrect Password", vbCritical, "Incorrect Password"
            Exit Sub
        End If
    End If
    
    TheMapType = MapType1
    MapTyper.Name = Name
    MapTyper.pass = Loadedpass
    MapTyper.music = music
    MapTyper.description = Desc
    'reset string. strip out .NWM
    X = Left(X, Len(X) - 4)
End If


    
    TempStr = X
    'loop and take out path. get map name.
    Do
    DoEvents
        TempNum = InStr(TempStr, "\")
        TempStr = Right(TempStr, Len(TempStr) - TempNum)
    Loop Until InStr(TempStr, "\") = 0
    TheDir1 = Left(X, Len(X) - TempNum + 1)





    'if map type is (1) MMORPG
    If TheMapType = 1 Then

    'set map type. set controls
    Call MapType(TheMapType)
       
    'if the filename isnt blank
    If X <> "" Then
        'get x and y sector
        
        TempStr = X
        'loop and take out path. get map name.
        Do
        DoEvents
            TempNum = InStr(TempStr, "\")
            TempStr = Right(TempStr, Len(TempStr) - TempNum)
        Loop Until InStr(TempStr, "\") = 0

        TempNum = InStr(TempStr, "-")
        XSector = Left(TempStr, TempNum - 1)
        YSector = Right(TempStr, Len(TempStr) - TempNum)

        'set sector offsets
        SectorXOffset = XSector
        SectorYOffset = YSector

        'set Sectors
        'set sectors if new highest loaded sector
        MaxXSector = XSector
        MaxYSector = YSector

        'load Layer1 map files
        'load # of maps depending on LoadSector varaibles.
        For XLong = SectorXOffset To LoadXSectors + SectorXOffset - 1
        For YLong = SectorYOffset To LoadYSectors + SectorYOffset - 1

        'open file if exist

        If Dir(XLong & "-" & YLong & "a.nwm") <> "" Then

        Open XLong & "-" & YLong & "a.nwm" For Input As #1
        Counters = 0
        'X and Y Differences of other maps loading after first map.
        TheXLong = ((XLong - SectorXOffset) * (480))
        TheYLong = ((YLong - SectorYOffset) * (480))
        'set sectors if new highest loaded sector
        If XLong > MaxXSector Then MaxXSector = XLong
        If YLong > MaxYSector Then MaxYSector = YLong

        
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$) + TheXLong

                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$) + TheYLong
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = (x1 + PicSizer.Width)
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile

                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 1, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
 
        End If
        
        Next YLong
        Next XLong


        'Load Layer2 Map Files
        'load # of maps depending on LoadSector varaibles.
        For XLong = SectorXOffset To LoadXSectors + SectorXOffset - 1
        For YLong = SectorYOffset To LoadYSectors + SectorYOffset - 1
        
        'open file if exist
        If Dir(XLong & "-" & YLong & "b.nwm") <> "" Then

        Open XLong & "-" & YLong & "b.nwm" For Input As #1
        Counters = 0
        'X and Y Differences of other maps loading after first map.
        TheXLong = ((XLong - SectorXOffset) * (480))
        TheYLong = ((YLong - SectorYOffset) * (480))
        'set sectors if new highest loaded sector
        If XLong > MaxXSector Then MaxXSector = XLong
        If YLong > MaxYSector Then MaxYSector = YLong

        
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$) + TheXLong
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$) + TheYLong
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 2, trigger, color)
                Counters = 0
            End If

        Wend

        Close #1

        End If
        
        Next YLong
        Next XLong
        
        
        'Load Layer3 map files.
        'load # of maps depending on LoadSector varaibles.
        For XLong = SectorXOffset To LoadXSectors + SectorXOffset - 1
        For YLong = SectorYOffset To LoadYSectors + SectorYOffset - 1
        

        'open file if exist
        If Dir(XLong & "-" & YLong & "c.nwm") <> "" Then

        Open XLong & "-" & YLong & "c.nwm" For Input As #1
        Counters = 0
        'X and Y Differences of other maps loading after first map.
        TheXLong = ((XLong - SectorXOffset) * (480))
        TheYLong = ((YLong - SectorYOffset) * (480))
        'set sectors if new highest loaded sector
        If XLong > MaxXSector Then MaxXSector = XLong
        If YLong > MaxYSector Then MaxYSector = YLong

        
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                If Trim(ln$) <> "" Then
                'set texture
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$) + TheXLong
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$) + TheYLong
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 3, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
        
        End If
        
        Next YLong
        Next XLong
        
        
        
        'Load Layer4 map files.
        'load # of maps depending on LoadSector varaibles.
        For XLong = SectorXOffset To LoadXSectors + SectorXOffset - 1
        For YLong = SectorYOffset To LoadYSectors + SectorYOffset - 1
        

        'open file if exist
        If Dir(XLong & "-" & YLong & "d.nwm") <> "" Then

        Open XLong & "-" & YLong & "d.nwm" For Input As #1
        Counters = 0
        'X and Y Differences of other maps loading after first map.
        TheXLong = ((XLong - SectorXOffset) * (480))
        TheYLong = ((YLong - SectorYOffset) * (480))
        'set sectors if new highest loaded sector
        If XLong > MaxXSector Then MaxXSector = XLong
        If YLong > MaxYSector Then MaxYSector = YLong

        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                If Trim(ln$) <> "" Then
                'set texture
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$) + TheXLong
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$) + TheYLong
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 4, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
        End If
        
        Next YLong
        Next XLong
        
        
        'Load Layer5 map files.
        'load # of maps depending on LoadSector varaibles.
        For XLong = SectorXOffset To LoadXSectors + SectorXOffset - 1
        For YLong = SectorYOffset To LoadYSectors + SectorYOffset - 1
        

        'open file if exist
        If Dir(XLong & "-" & YLong & "e.nwm") <> "" Then

        Open XLong & "-" & YLong & "e.nwm" For Input As #1
        Counters = 0
        'X and Y Differences of other maps loading after first map.
        TheXLong = ((XLong - SectorXOffset) * (480))
        TheYLong = ((YLong - SectorYOffset) * (480))
        'set sectors if new highest loaded sector
        If XLong > MaxXSector Then MaxXSector = XLong
        If YLong > MaxYSector Then MaxYSector = YLong

        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                If Trim(ln$) <> "" Then
                'set texture
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$) + TheXLong
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$) + TheYLong
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 5, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
        End If
        
        Next YLong
        Next XLong
        
        Active = True
        Save = True
    End If
    

    
    'if maptype is (2) Battle Scene
    ElseIf TheMapType = 2 Then
    
    'set map type. set controls
    Call MapType(TheMapType)
    
        'Load Layer1
        'open file if exist
        If Dir(X & "a.nwm") <> "" Then

        Open X & "a.nwm" For Input As #1
        Counters = 0
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$)
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 1, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
 
        End If
    
        'Load Layer2
        'open file if exist
        If Dir(X & "b.nwm") <> "" Then

        Open X & "b.nwm" For Input As #1
        Counters = 0
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$)
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 2, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
 
        End If
        
        
        'Load Layer3
        'open file if exist
        If Dir(X & "c.nwm") <> "" Then
        Open X & "c.nwm" For Input As #1
        
        Counters = 0
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$)
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 3, trigger, color)
                Counters = 0
            End If

        Wend

        Close #1

        End If

    
        'Load Layer1
        'open file if exist
        If Dir(X & "d.nwm") <> "" Then

        Open X & "d.nwm" For Input As #1
        Counters = 0
        'loop until end of file
        While Not (EOF(1))
            'read line
            Input #1, ln$
            If Counters = 0 Then
                'set texture
                If Trim(ln$) <> "" Then
                Texture = App.path & "\tiles\" & Trim(ln$)
                PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Trim(ln$))
                Counters = Counters + 1
                End If
            ElseIf Counters = 1 Then
                'set x1
                x1 = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 2 Then
                'set y1
                y1 = Trim(ln$)
                'set tiles.x and tiles.y
                x2 = x1 + PicSizer.Width
                y2 = y1
                
                x3 = x1
                y3 = y1 + PicSizer.Height
            
                x4 = x1 + PicSizer.Width
                y4 = y1 + PicSizer.Height

                Counters = Counters + 1
            ElseIf Counters = 3 Then
                trigger = Trim(ln$)
                Counters = Counters + 1
            ElseIf Counters = 4 Then
                color = Trim(ln$)
                'place tile
                Call PlaceTileWithoutCheck(x1, x2, x3, x4, y1, y2, y3, y4, Texture, 4, trigger, color)
                Counters = 0
            End If
                
        Wend

        Close #1
 
        End If
        
        Active = True
        Save = True
    End If






End Sub

Public Sub MapType(TheMapType As Integer)
'1 = The Smash

    ChkLayer1.Visible = False
    ChkLayer2.Visible = False
    ChkLayer3.Visible = False
    ChkLayer4.Visible = False
    ChkLayer5.Visible = False

    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    Option5.Visible = False
    
    Select Case TheMapType
        Case 1:
            'fix chklayer1 properties so everything is visible
            ChkLayer1.Caption = "Background (TL)"
            ChkLayer1.ForeColor = vbRed
            ChkLayer1.Width = 105
            

            'fix options1 properties so everything is visible
            Option1.Caption = "Background (TL)"
            Option1.ForeColor = vbRed
            Option1.Width = 105
            
            'fix rest of check box's and option box's
            'so they dont overlap.
            
            ChkLayer2.Caption = "Walkable (TL)"
            ChkLayer2.ForeColor = vbRed
            ChkLayer2.Left = 336
            ChkLayer2.Width = 97
            
            Option2.Caption = "Walkable (TL)"
            Option2.ForeColor = vbRed
            Option2.Left = 336
            Option2.Width = 97
            
            'if menu 2 is loaded, set caption
            
            If mnuMapClearLayerLayer(1) Is Nothing Then
                mnuMapClearLayerLayer(1).Caption = "Layer 2"
            'if menu 2 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(1)
                mnuMapClearLayerLayer(1).Visible = True
                mnuMapClearLayerLayer(1).Caption = "Layer 2"
            End If
            
            ChkLayer3.Caption = "Object Layer (OL)"
            ChkLayer3.ForeColor = &H80FF&
            ChkLayer3.Left = 440
            ChkLayer3.Width = 105
            
            Option3.Caption = "Object Layer (OL)"
            Option3.ForeColor = &H80FF&
            Option3.Left = 440
            Option3.Width = 105
            
            'if menu 2 is loaded, set caption
            If mnuMapClearLayerLayer(2) Is Nothing Then
                mnuMapClearLayerLayer(2).Caption = "Layer 3"
            'if menu 2 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(2)
                mnuMapClearLayerLayer(2).Visible = True
                mnuMapClearLayerLayer(2).Caption = "Layer 3"
            End If
            
            ChkLayer4.Caption = "OverLap (OL)"
            ChkLayer4.ForeColor = &H80FF&
            ChkLayer4.Left = 550
            ChkLayer4.Width = 89
            
            Option4.Caption = "OverLap (OL)"
            Option4.ForeColor = &H80FF&
            Option4.Left = 550
            Option4.Width = 89
              
            'if menu 4 is loaded, set caption
            If mnuMapClearLayerLayer(3) Is Nothing Then
                mnuMapClearLayerLayer(3).Caption = "Layer 4"
            'if menu 4 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(3)
                mnuMapClearLayerLayer(3).Visible = True
                mnuMapClearLayerLayer(3).Caption = "Layer 4"
            End If
            
            ChkLayer5.Caption = "Overlap (OL)"
            ChkLayer5.ForeColor = vbBlue
            ChkLayer5.Left = 636
            ChkLayer5.Width = 97

            Option5.Caption = "Overlap (OL)"
            Option5.ForeColor = vbBlue
            Option5.Left = 636
            Option5.Width = 97


            'show the Check boxes
            ChkLayer1.Visible = True
            ChkLayer2.Visible = True
            ChkLayer3.Visible = True
            ChkLayer4.Visible = True
            ChkLayer5.Visible = False
            
            'show option box's
            Option1.Visible = True
            Option2.Visible = True
            Option3.Visible = True
            Option4.Visible = True
            Option5.Visible = False
            
            
        Case 2:
            'fix chklayer1 properties so everything is visible
            ChkLayer1.Caption = "Background (TL)"
            ChkLayer1.ForeColor = vbRed
            ChkLayer1.Width = 105
            
            'fix options1 properties so everything is visible
            Option1.Caption = "Background (TL)"
            Option1.ForeColor = vbRed
            Option1.Width = 105
            
            'fix rest of check box's and option box's
            'so they dont overlap.
            
            ChkLayer2.Caption = "Walkable (TL)"
            ChkLayer2.ForeColor = vbRed
            ChkLayer2.Left = 336
            ChkLayer2.Width = 97
            
            Option2.Caption = "Walkable (TL)"
            Option2.ForeColor = vbRed
            Option2.Left = 336
            Option2.Width = 97
            
            'if menu 2 is loaded, set caption
            If mnuMapClearLayerLayer(1) Is Nothing Then
                mnuMapClearLayerLayer(1).Caption = "Layer 2"
            'if menu 2 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(1)
                mnuMapClearLayerLayer(1).Visible = True
                mnuMapClearLayerLayer(1).Caption = "Layer 2"
            End If
            
            ChkLayer3.Caption = "Object Layer (OL)"
            ChkLayer3.ForeColor = &H80FF&
            ChkLayer3.Left = 440
            ChkLayer3.Width = 105
            
            Option3.Caption = "Object Layer (OL)"
            Option3.ForeColor = &H80FF&
            Option3.Left = 440
            Option3.Width = 105
            
            'if menu 3 is loaded, set caption
            If mnuMapClearLayerLayer(2) Is Nothing Then
                mnuMapClearLayerLayer(2).Caption = "Layer 3"
            'if menu 3 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(2)
                mnuMapClearLayerLayer(2).Visible = True
                mnuMapClearLayerLayer(2).Caption = "Layer 3"
            End If
            
            
            ChkLayer4.Caption = "OverLap (OL)"
            ChkLayer4.ForeColor = &H80FF&
            ChkLayer4.Left = 550
            ChkLayer4.Width = 89
            
            Option4.Caption = "OverLap (OL)"
            Option4.ForeColor = &H80FF&
            Option4.Left = 550
            Option4.Width = 89
              
            'if menu 4 is loaded, set caption
            If mnuMapClearLayerLayer(3) Is Nothing Then
                mnuMapClearLayerLayer(3).Caption = "Layer 4"
            'if menu 4 loaded. load and set caption
            Else
                Load mnuMapClearLayerLayer(3)
                mnuMapClearLayerLayer(3).Visible = True
                mnuMapClearLayerLayer(3).Caption = "Layer 4"
            End If
            
            
            ChkLayer5.Caption = "Overlap (OL)"
            ChkLayer5.ForeColor = vbBlue
            ChkLayer5.Left = 636
            ChkLayer5.Width = 97

            Option5.Caption = "Overlap (OL)"
            Option5.ForeColor = vbBlue
            Option5.Left = 636
            Option5.Width = 97


            'show the Check boxes
            ChkLayer1.Visible = True
            ChkLayer2.Visible = True
            ChkLayer3.Visible = True
            ChkLayer4.Visible = True
            ChkLayer5.Visible = False
            
            'show option box's
            Option1.Visible = True
            Option2.Visible = True
            Option3.Visible = True
            Option4.Visible = True
            Option5.Visible = False
    End Select
    
    Save = False
    
End Sub


Public Function MonCheck(Name As String) As Boolean
Dim i As Integer
    'loop threw all npcs in memory
    For i = 0 To MonsterCounter
        If LCase(Monster(i).Name) = LCase(Name) Then
            MonCheck = True
            Exit Function
        End If
    Next i

    MonCheck = False
End Function

Public Sub MoveLayer(X As Single, Y As Single, layer As Single)
Dim TheX As Single
Dim TheY As Single

TheX = Round(((X)) / 32, 0)
TheY = Round(((Y)) / 32, 0)

'if moving from layer 2 to layer 1
If layer = 1 Then
    'if tile exist
    If Map(TheX, TheY).Used = True And Map(TheX, TheY).Walk = True Then


    Call PlaceTileWithoutCheck(TilesLayer2(Map(TheX, TheY).TileRefNum).X, TilesLayer2(Map(TheX, TheY).TileRefNum + 1).X, TilesLayer2(Map(TheX, TheY).TileRefNum + 2).X, TilesLayer2(Map(TheX, TheY).TileRefNum + 3).X, TilesLayer2(Map(TheX, TheY).TileRefNum).Y, TilesLayer2(Map(TheX, TheY).TileRefNum + 1).Y, TilesLayer2(Map(TheX, TheY).TileRefNum + 2).Y, TilesLayer2(Map(TheX, TheY).TileRefNum + 3).Y, TheTextures(Map(TheX, TheY).TextureRefNum), 1, Map(TheX, TheY).trigger, Map(TheX, TheY).color)
    Call DeleteTile(TheX, TheY, 2)
    Map(TheX, TheY).Walk = False
    
    End If
'if moving from layer 1 to 2
ElseIf layer = 2 Then

    'if tile exist
    If Map(TheX, TheY).Used = True And Map(TheX, TheY).Walk = False Then
    

    Call PlaceTileWithoutCheck(TilesLayer1(Map(TheX, TheY).TileRefNum).X, TilesLayer1(Map(TheX, TheY).TileRefNum + 1).X, TilesLayer1(Map(TheX, TheY).TileRefNum + 2).X, TilesLayer1(Map(TheX, TheY).TileRefNum + 3).X, TilesLayer1(Map(TheX, TheY).TileRefNum).Y, TilesLayer1(Map(TheX, TheY).TileRefNum + 1).Y, TilesLayer1(Map(TheX, TheY).TileRefNum + 2).Y, TilesLayer1(Map(TheX, TheY).TileRefNum + 3).Y, TheTextures(Map(TheX, TheY).TextureRefNum), 2, Map(TheX, TheY).trigger, Map(TheX, TheY).color)
    Call DeleteTile(TheX, TheY, 1)
    Map(TheX, TheY).Walk = True
    
    End If
End If

End Sub

Public Sub NewMap()

ReDim TilesLayer1(0)
ReDim TilesLayer2(0)
ReDim TilesLayer3(0)
ReDim TilesLayer4(0)
ReDim TilesLayer5(0)

ReDim LayerTiles1(0)
ReDim LayerTiles2(0)
ReDim LayerTiles3(0)
ReDim LayerTiles4(0)
ReDim LayerTiles5(0)


TileCounter1 = -1
TileCounter2 = -1
TileCounter3 = -1
TileCounter4 = -1
TileCounter5 = -1

TextureCounter1 = -1
TextureCounter2 = -1
TextureCounter3 = -1
TextureCounter4 = -1
TextureCounter5 = -1


ReDim Textures1(0)
ReDim Textures2(0)
ReDim Textures3(0)
ReDim Textures4(0)
ReDim Textures5(0)

ReDim TheTextures1(0)
ReDim TheTextures2(0)
ReDim TheTextures3(0)
ReDim TheTextures4(0)
ReDim TheTextures5(0)

frmMapType.Show , Me
Active = True
Save = True
End Sub

Public Function NPCCheck(Name As String) As Boolean
Dim i As Integer
    'loop threw all npcs in memory
    For i = 0 To NPCCounter
        If LCase(NPC(i).Name) = LCase(Name) Then
            NPCCheck = True
            Exit Function
        End If
    Next i
    
    NPCCheck = False
End Function


Public Sub PlaceTile(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture As String, TileLayer As Integer, trigger As String, color As Long)
Dim TilePlacer As Long

Dim GoGood As Boolean
Dim TextureGood As Boolean
Dim SectorX As Long
Dim SectorY As Long
Dim TheX As Single
Dim TheY As Single
Dim i As Long
Dim z As Long
Dim a As Integer
GoGood = True
TextureGood = False
If LCase(Right(Texture, 4)) = ".bmp" Then

'get the sector that the tile is placed in
SectorX = Int((x1 + XPixelDiff) / 480) + SectorXOffset
SectorY = Int((y1 + YPixelDiff) / 480) + SectorYOffset



TheX = Round(((x1)) / 32, 0)
TheY = Round(((y1)) / 32, 0)

'if sector = 0 force to be sector1
If SectorX = 0 Then
    SectorX = 1
End If

'if sector = 0 force to be sector1
If SectorY = 0 Then
    SectorY = 1
End If

'if xsector > xmax sector, set new max sector
If SectorX > MaxXSector Then
    MaxXSector = SectorX
End If

'if ysector > ymax sector set new max sector
If SectorY > MaxYSector Then
    MaxYSector = SectorY
End If

    'if TileLayer is 5,
    If TileLayer = 5 Then

    'loop to check every stored Vertex.
    'Checks to see if the co-ords are allready in use
    'TileCounter * 4 = Total Vertex Count.
    For i = 0 To (TileCounter5 * 4) Step 4

    'first matching X pixel in use
    If x1 = (TilesLayer5(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If x2 = (TilesLayer5(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If y1 = (TilesLayer5(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If y3 = (TilesLayer5(i + 2).Y) - YPixelDiff Then
                    'Exit. We dont want to place a tile
                    GoGood = False
                    Exit Sub
                End If
            End If
        End If
    End If

    Next i

    'if we want to place tile
    If GoGood = True Then
    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter5 = TextureCounter5 + 1
    ReDim Preserve TheTextures5(TextureCounter5)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures5(TextureCounter5) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures5(TextureCounter5) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures5(TextureCounter5)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures5(TextureCounter5) = TextureCounter
        End If
    End If
        
    'Increase Tile Counter & Tile Placer
    TileCounter5 = TileCounter5 + 1
    TilePlacer = TileCounter5 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer5(TilePlacer + 3)
    ReDim Preserve LayerTiles5(TilePlacer + 3)
    LayerTiles5(TilePlacer).color = color
    LayerTiles5(TilePlacer).trigger = trigger  'set trigger
    LayerTiles5(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles5(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles5(TilePlacer).X = x1
    LayerTiles5(TilePlacer).Y = y1
    TilesLayer5(TilePlacer) = CreateTLVertex(LayerTiles5(TilePlacer).X, LayerTiles5(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles5(TilePlacer + 1).X = x2
    LayerTiles5(TilePlacer + 1).Y = y2
    TilesLayer5(TilePlacer + 1) = CreateTLVertex(LayerTiles5(TilePlacer + 1).X, LayerTiles5(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles5(TilePlacer + 2).X = x3
    LayerTiles5(TilePlacer + 2).Y = y3
    TilesLayer5(TilePlacer + 2) = CreateTLVertex(LayerTiles5(TilePlacer + 2).X, LayerTiles5(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles5(TilePlacer + 3).X = x4
    LayerTiles5(TilePlacer + 3).Y = y4
    TilesLayer5(TilePlacer + 3) = CreateTLVertex(LayerTiles5(TilePlacer + 3).X, LayerTiles5(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)

    'end if for good=true
    End If

    'tileLayer is 4(Layer we are working with)
    ElseIf TileLayer = 4 Then

    
    'loop to check every stored Vertex.
    'Checks to see if the co-ords are allready in use
    'TileCounter * 4 = Total Vertex Count.
    For i = 0 To (TileCounter4 * 4) Step 4

    'first matching X pixel in use
    If x1 = (TilesLayer4(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If x2 = (TilesLayer4(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If y1 = (TilesLayer4(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If y3 = (TilesLayer4(i + 2).Y) - YPixelDiff Then
                    'Exit. We dont want to place a tile
                    GoGood = False
                    Exit Sub
                End If
            End If
        End If
    End If

    Next i

    'if we want to place tile
    If GoGood = True Then
    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter4 = TextureCounter4 + 1
    
    ReDim Preserve TheTextures4(TextureCounter4)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures4(TextureCounter4) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures4(TextureCounter4) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures4(TextureCounter4)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures4(TextureCounter4) = TextureCounter
        End If
    End If


    'Increase Tile Counter & Tile Placer
    TileCounter4 = TileCounter4 + 1
    TilePlacer = TileCounter4 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer4(TilePlacer + 3)
    ReDim Preserve LayerTiles4(TilePlacer + 3)
    LayerTiles4(TilePlacer).color = color
    LayerTiles4(TilePlacer).trigger = trigger  'set trigger
    LayerTiles4(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles4(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles4(TilePlacer).X = x1
    LayerTiles4(TilePlacer).Y = y1
    TilesLayer4(TilePlacer) = CreateTLVertex(LayerTiles4(TilePlacer).X, LayerTiles4(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles4(TilePlacer + 1).X = x2
    LayerTiles4(TilePlacer + 1).Y = y2
    TilesLayer4(TilePlacer + 1) = CreateTLVertex(LayerTiles4(TilePlacer + 1).X, LayerTiles4(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles4(TilePlacer + 2).X = x3
    LayerTiles4(TilePlacer + 2).Y = y3
    TilesLayer4(TilePlacer + 2) = CreateTLVertex(LayerTiles4(TilePlacer + 2).X, LayerTiles4(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles4(TilePlacer + 3).X = x4
    LayerTiles4(TilePlacer + 3).Y = y4
    TilesLayer4(TilePlacer + 3) = CreateTLVertex(LayerTiles4(TilePlacer + 3).X, LayerTiles4(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)

    'end if for good=true
    End If
   
    
    'tileLayer is 3(Layer we are working with)
    ElseIf TileLayer = 3 Then
    
    
    'loop to check every stored Vertex.
    'Checks to see if the co-ords are allready in use
    'TileCounter * 4 = Total Vertex Count.
    For i = 0 To (TileCounter3 * 4) Step 4

    'first matching X pixel in use
    If x1 = (TilesLayer3(i).X) - XPixelDiff Then
        'second matchin X pixel in use
        If x2 = (TilesLayer3(i + 1).X) - XPixelDiff Then
            'first matchin Y pixel in use
            If y1 = (TilesLayer3(i).Y) - YPixelDiff Then
                'second matching Y pixel in use
                If y3 = (TilesLayer3(i + 2).Y) - YPixelDiff Then
                    'Exit. We dont want to place a tile
                    GoGood = False
                    Exit Sub
                End If
            End If
        End If
    End If

    Next i
    
    'if we want to place tile
    If GoGood = True Then
    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter3 = TextureCounter3 + 1
    ReDim Preserve TheTextures3(TextureCounter3)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures3(TextureCounter3) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures3(TextureCounter3) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures3(TextureCounter3)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures3(TextureCounter3) = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter3 = TileCounter3 + 1
    TilePlacer = TileCounter3 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer3(TilePlacer + 3)
    ReDim Preserve LayerTiles3(TilePlacer + 3)
    LayerTiles3(TilePlacer).color = color
    LayerTiles3(TilePlacer).trigger = trigger  'set trigger
    LayerTiles3(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles3(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles3(TilePlacer).X = x1
    LayerTiles3(TilePlacer).Y = y1
    TilesLayer3(TilePlacer) = CreateTLVertex(LayerTiles3(TilePlacer).X, LayerTiles3(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles3(TilePlacer + 1).X = x2
    LayerTiles3(TilePlacer + 1).Y = y2
    TilesLayer3(TilePlacer + 1) = CreateTLVertex(LayerTiles3(TilePlacer + 1).X, LayerTiles3(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles3(TilePlacer + 2).X = x3
    LayerTiles3(TilePlacer + 2).Y = y3
    TilesLayer3(TilePlacer + 2) = CreateTLVertex(LayerTiles3(TilePlacer + 2).X, LayerTiles3(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles3(TilePlacer + 3).X = x4
    LayerTiles3(TilePlacer + 3).Y = y4
    TilesLayer3(TilePlacer + 3) = CreateTLVertex(LayerTiles3(TilePlacer + 3).X, LayerTiles3(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)

    'end if for good=true
    End If

    
    'tileLayer is 2(Layer we are working with)
    ElseIf TileLayer = 2 Then
    
    

    'Checks to see if tile is allready placed
    If Map(TheX, TheY).Used = True Then
        GoGood = False
    End If
 
    'if we want to place tile
    If GoGood = True Then
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        Map(TheX, TheY).TextureRefNum = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                Map(TheX, TheY).TextureRefNum = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            Map(TheX, TheY).TextureRefNum = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter2 = TileCounter2 + 1
    TilePlacer = TileCounter2 * 4


    ReDim Preserve TilesLayer2(TilePlacer + 3)
    
    Map(TheX, TheY).TileRefNum = TilePlacer
    Map(TheX, TheY).color = color
    Map(TheX, TheY).X = x1
    Map(TheX, TheY).Y = y1
    Map(TheX, TheY).XSector = SectorX
    Map(TheX, TheY).YSector = SectorY
    Map(TheX, TheY).trigger = trigger
    Map(TheX, TheY).Used = True
    Map(TheX, TheY).Walk = True

    'Set the 4 Corners of the square
    TilesLayer2(TilePlacer) = CreateTLVertex(x1, y1, 0, 1, color, 0, 0, 0)
    TilesLayer2(TilePlacer + 1) = CreateTLVertex(x2, y2, 0, 1, color, 0, 1, 0)
    TilesLayer2(TilePlacer + 2) = CreateTLVertex(x3, y3, 0, 1, color, 0, 0, 1)
    TilesLayer2(TilePlacer + 3) = CreateTLVertex(x4, y4, 0, 1, color, 0, 1, 1)

    'end if for good=true
    End If

    
    'tileLayer is 1(Layer we are working with)
    ElseIf TileLayer = 1 Then


    'Checks to see if tile is allready placed
    If Map(TheX, TheY).Used = True Then
        GoGood = False
    End If

    'if we want to place tile
    If GoGood = True Then
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        Map(TheX, TheY).TextureRefNum = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                Map(TheX, TheY).TextureRefNum = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            Map(TheX, TheY).TextureRefNum = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter1 = TileCounter1 + 1
    TilePlacer = TileCounter1 * 4


    ReDim Preserve TilesLayer1(TilePlacer + 3)

    Map(TheX, TheY).TileRefNum = TilePlacer
    Map(TheX, TheY).color = color
    Map(TheX, TheY).X = x1
    Map(TheX, TheY).Y = y1
    Map(TheX, TheY).XSector = SectorX
    Map(TheX, TheY).YSector = SectorY
    Map(TheX, TheY).trigger = trigger
    Map(TheX, TheY).Used = True
    Map(TheX, TheY).Walk = False

    'Set the 4 Corners of the square
    TilesLayer1(TilePlacer) = CreateTLVertex(x1, y1, 0, 1, color, 0, 0, 0)
    TilesLayer1(TilePlacer + 1) = CreateTLVertex(x2, y2, 0, 1, color, 0, 1, 0)
    TilesLayer1(TilePlacer + 2) = CreateTLVertex(x3, y3, 0, 1, color, 0, 0, 1)
    TilesLayer1(TilePlacer + 3) = CreateTLVertex(x4, y4, 0, 1, color, 0, 1, 1)

    'end if for good=true
    End If

    'end of tilelayers
    End If
'end if of texture check
End If
End Sub



Public Sub PlaceTileWithoutCheck(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture As String, TileLayer As Integer, trigger As String, color As Long)
Dim TilePlacer As Long

Dim GoGood As Boolean
Dim TextureGood As Boolean
Dim SectorX As Long
Dim SectorY As Long
Dim TheX As Long
Dim TheY As Long
Dim i As Long
Dim z As Long
Dim a As Integer
GoGood = True
TextureGood = False
If LCase(Right(Texture, 4)) = ".bmp" Then

'get the sector that the tile is placed in
SectorX = Int((x1 + XPixelDiff) / 480) + SectorXOffset
SectorY = Int((y1 + YPixelDiff) / 480) + SectorYOffset


TheX = (x1 + XPixelDiff) / 32
TheY = (y1 + YPixelDiff) / 32

'if sector = 0 force to be sector1
If SectorX = 0 Then
    SectorX = 1
End If

'if sector = 0 force to be sector1
If SectorY = 0 Then
    SectorY = 1
End If

'if xsector > xmax sector, set new max sector
If SectorX > MaxXSector Then
    MaxXSector = SectorX
End If

'if ysector > ymax sector set new max sector
If SectorY > MaxYSector Then
    MaxYSector = SectorY
End If

    'if TileLayer is 5,
    If TileLayer = 5 Then


    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter5 = TextureCounter5 + 1
    ReDim Preserve TheTextures5(TextureCounter5)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures5(TextureCounter5) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures5(TextureCounter5) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures5(TextureCounter5)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures5(TextureCounter5) = TextureCounter
        End If
    End If
        
    'Increase Tile Counter & Tile Placer
    TileCounter5 = TileCounter5 + 1
    TilePlacer = TileCounter5 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer5(TilePlacer + 3)
    ReDim Preserve LayerTiles5(TilePlacer + 3)
    LayerTiles5(TilePlacer).color = color
    LayerTiles5(TilePlacer).trigger = trigger  'set trigger
    LayerTiles5(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles5(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles5(TilePlacer).X = x1
    LayerTiles5(TilePlacer).Y = y1
    TilesLayer5(TilePlacer) = CreateTLVertex(LayerTiles5(TilePlacer).X, LayerTiles5(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles5(TilePlacer + 1).X = x2
    LayerTiles5(TilePlacer + 1).Y = y2
    TilesLayer5(TilePlacer + 1) = CreateTLVertex(LayerTiles5(TilePlacer + 1).X, LayerTiles5(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles5(TilePlacer + 2).X = x3
    LayerTiles5(TilePlacer + 2).Y = y3
    TilesLayer5(TilePlacer + 2) = CreateTLVertex(LayerTiles5(TilePlacer + 2).X, LayerTiles5(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles5(TilePlacer + 3).X = x4
    LayerTiles5(TilePlacer + 3).Y = y4
    TilesLayer5(TilePlacer + 3) = CreateTLVertex(LayerTiles5(TilePlacer + 3).X, LayerTiles5(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)


    'tileLayer is 4(Layer we are working with)
    ElseIf TileLayer = 4 Then


    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter4 = TextureCounter4 + 1
    
    ReDim Preserve TheTextures4(TextureCounter4)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures4(TextureCounter4) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures4(TextureCounter4) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures4(TextureCounter4)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures4(TextureCounter4) = TextureCounter
        End If
    End If


    'Increase Tile Counter & Tile Placer
    TileCounter4 = TileCounter4 + 1
    TilePlacer = TileCounter4 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer4(TilePlacer + 3)
    ReDim Preserve LayerTiles4(TilePlacer + 3)
    LayerTiles4(TilePlacer).color = color
    LayerTiles4(TilePlacer).trigger = trigger  'set trigger
    LayerTiles4(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles4(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles4(TilePlacer).X = x1
    LayerTiles4(TilePlacer).Y = y1
    TilesLayer4(TilePlacer) = CreateTLVertex(LayerTiles4(TilePlacer).X, LayerTiles4(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles4(TilePlacer + 1).X = x2
    LayerTiles4(TilePlacer + 1).Y = y2
    TilesLayer4(TilePlacer + 1) = CreateTLVertex(LayerTiles4(TilePlacer + 1).X, LayerTiles4(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles4(TilePlacer + 2).X = x3
    LayerTiles4(TilePlacer + 2).Y = y3
    TilesLayer4(TilePlacer + 2) = CreateTLVertex(LayerTiles4(TilePlacer + 2).X, LayerTiles4(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles4(TilePlacer + 3).X = x4
    LayerTiles4(TilePlacer + 3).Y = y4
    TilesLayer4(TilePlacer + 3) = CreateTLVertex(LayerTiles4(TilePlacer + 3).X, LayerTiles4(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)

   
    
    'tileLayer is 3(Layer we are working with)
    ElseIf TileLayer = 3 Then
    


    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter3 = TextureCounter3 + 1
    ReDim Preserve TheTextures3(TextureCounter3)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        TheTextures3(TextureCounter3) = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                TheTextures3(TextureCounter3) = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures3(TextureCounter3)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            TheTextures3(TextureCounter3) = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter3 = TileCounter3 + 1
    TilePlacer = TileCounter3 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer3(TilePlacer + 3)
    ReDim Preserve LayerTiles3(TilePlacer + 3)
    LayerTiles3(TilePlacer).color = color
    LayerTiles3(TilePlacer).trigger = trigger  'set trigger
    LayerTiles3(TilePlacer).XSector = SectorX 'set sectors
    LayerTiles3(TilePlacer).YSector = SectorY 'set sectors
    LayerTiles3(TilePlacer).X = x1
    LayerTiles3(TilePlacer).Y = y1
    TilesLayer3(TilePlacer) = CreateTLVertex(LayerTiles3(TilePlacer).X, LayerTiles3(TilePlacer).Y, 0, 1, color, 0, 0, 0)


    LayerTiles3(TilePlacer + 1).X = x2
    LayerTiles3(TilePlacer + 1).Y = y2
    TilesLayer3(TilePlacer + 1) = CreateTLVertex(LayerTiles3(TilePlacer + 1).X, LayerTiles3(TilePlacer + 1).Y, 0, 1, color, 0, 1, 0)


    LayerTiles3(TilePlacer + 2).X = x3
    LayerTiles3(TilePlacer + 2).Y = y3
    TilesLayer3(TilePlacer + 2) = CreateTLVertex(LayerTiles3(TilePlacer + 2).X, LayerTiles3(TilePlacer + 2).Y, 0, 1, color, 0, 0, 1)


    LayerTiles3(TilePlacer + 3).X = x4
    LayerTiles3(TilePlacer + 3).Y = y4
    TilesLayer3(TilePlacer + 3) = CreateTLVertex(LayerTiles3(TilePlacer + 3).X, LayerTiles3(TilePlacer + 3).Y, 0, 1, color, 0, 1, 1)

    
    'tileLayer is 2(Layer we are working with)
    ElseIf TileLayer = 2 Then
    
    
    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter2 = TextureCounter2 + 1
    ReDim Preserve TheTextures2(TextureCounter2)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        Map(TheX, TheY).TextureRefNum = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                Map(TheX, TheY).TextureRefNum = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            Map(TheX, TheY).TextureRefNum = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter2 = TileCounter2 + 1
    TilePlacer = TileCounter2 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer2(TilePlacer + 3)

    Map(TheX, TheY).TileRefNum = TilePlacer
    Map(TheX, TheY).color = color
    Map(TheX, TheY).X = x1
    Map(TheX, TheY).Y = y1
    Map(TheX, TheY).XSector = SectorX
    Map(TheX, TheY).YSector = SectorY
    Map(TheX, TheY).trigger = trigger
    Map(TheX, TheY).Used = True
    Map(TheX, TheY).Walk = True
    
    TilesLayer2(TilePlacer) = CreateTLVertex(x1, y1, 0, 1, color, 0, 0, 0)
    TilesLayer2(TilePlacer + 1) = CreateTLVertex(x2, y2, 0, 1, color, 0, 1, 0)
    TilesLayer2(TilePlacer + 2) = CreateTLVertex(x3, y3, 0, 1, color, 0, 0, 1)
    TilesLayer2(TilePlacer + 3) = CreateTLVertex(x4, y4, 0, 1, color, 0, 1, 1)


    
    'tileLayer is 1(Layer we are working with)
    ElseIf TileLayer = 1 Then


    'increase TextureCounter. one texture per Vertex.
    'increase Textures.
    TextureCounter1 = TextureCounter1 + 1
    ReDim Preserve TheTextures1(TextureCounter1)
    'if first texture used
    If TextureCounter = -1 Then
        'increase textureCounter
        'increase textures, thetextures
        TextureCounter = TextureCounter + 1
        ReDim Preserve Textures(TextureCounter)
        ReDim Preserve TheTextures(TextureCounter)
        
        'store texture, set texture name, and texture refrence number
        Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        TheTextures(TextureCounter) = Texture
        Map(TheX, TheY).TextureRefNum = TextureCounter
    'if not first texture used
    Else
        'loop threw all textures used
        For a = LBound(TheTextures()) To UBound(TheTextures())
            'check if texture is allready in memory
            If TheTextures(a) = Texture Then
                'store texture refrence number
                Map(TheX, TheY).TextureRefNum = a
                TextureGood = True
                Exit For
            End If
        Next a
                
        'if texture not in memory
        If Not TextureGood Then
            'increase texturecounter
            TextureCounter = TextureCounter + 1
            ReDim Preserve Textures(TextureCounter)
            ReDim Preserve TheTextures(TextureCounter)

            'store texture in memory
            Set Textures(TextureCounter) = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
            'store texture name, texture refrence number
            TheTextures(TextureCounter) = Texture
            Map(TheX, TheY).TextureRefNum = TextureCounter
        End If
    End If

    'Increase Tile Counter & Tile Placer
    TileCounter1 = TileCounter1 + 1
    TilePlacer = TileCounter1 * 4

    'Set the 4 Corners of the square
    ReDim Preserve TilesLayer1(TilePlacer + 3)

    Map(TheX, TheY).TileRefNum = TilePlacer
    Map(TheX, TheY).color = color
    Map(TheX, TheY).X = x1
    Map(TheX, TheY).Y = y1
    Map(TheX, TheY).XSector = SectorX
    Map(TheX, TheY).YSector = SectorY
    Map(TheX, TheY).trigger = trigger
    Map(TheX, TheY).Used = True
    Map(TheX, TheY).Walk = False

    
    TilesLayer1(TilePlacer) = CreateTLVertex(x1, y1, 0, 1, color, 0, 0, 0)
    TilesLayer1(TilePlacer + 1) = CreateTLVertex(x2, y2, 0, 1, color, 0, 1, 0)
    TilesLayer1(TilePlacer + 2) = CreateTLVertex(x3, y3, 0, 1, color, 0, 0, 1)
    TilesLayer1(TilePlacer + 3) = CreateTLVertex(x4, y4, 0, 1, color, 0, 1, 1)


    'end of tilelayers
    End If
'end if of texture check
End If
End Sub

Public Sub Render()
'//1. We need to clear the render device before we can draw anything
'       This must always happen before you start rendering stuff...
Dim TheCounter As String
Dim X As Integer
Dim Y As Integer
Dim i As Integer
'if window not minimized
If Me.WindowState <> 1 Then

On Error Resume Next
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0 '//Clear the screen black
Rendered = 0
'//2. Rendering the graphics...



D3DDevice.BeginScene
    'All rendering calls go between these two lines

'if running in game mode
If chkGame.Value = 1 Then


    If TileCounter1 >= 0 Then
    'set counter
    TheCounter = 0

    For X = XPixelDiff / 32 To (XPixelDiff / 32) + (XPixelDiff / 32) + Int(PicLevel.Width / 32) + 1
    For Y = YPixelDiff / 32 To (YPixelDiff / 32) + (YPixelDiff / 32) + Int(PicLevel.Height / 32) + 1
        
        If Map(X, Y).Used = True And Map(X, Y).Walk = False Then
        Rendered = Rendered + 1
        'If X = 11 And Y = 4 Then
        TilesLayer1((Map(X, Y).TileRefNum)).X = TilesLayer1((Map(X, Y).TileRefNum)).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum)).Y = TilesLayer1((Map(X, Y).TileRefNum)).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).X = TilesLayer1((Map(X, Y).TileRefNum) + 1).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).Y = TilesLayer1((Map(X, Y).TileRefNum) + 1).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).X = TilesLayer1((Map(X, Y).TileRefNum) + 2).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).Y = TilesLayer1((Map(X, Y).TileRefNum) + 2).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).X = TilesLayer1((Map(X, Y).TileRefNum) + 3).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).Y = TilesLayer1((Map(X, Y).TileRefNum) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(Map(X, Y).TextureRefNum)  '//Tell the device which texture we want to use...
        TheCounter = Map(X, Y).TileRefNum / 4
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer1(Map(X, Y).TileRefNum), Len(TilesLayer1(TheCounter))

        TilesLayer1((Map(X, Y).TileRefNum)).X = TilesLayer1((Map(X, Y).TileRefNum)).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum)).Y = TilesLayer1((Map(X, Y).TileRefNum)).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).X = TilesLayer1((Map(X, Y).TileRefNum) + 1).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).Y = TilesLayer1((Map(X, Y).TileRefNum) + 1).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).X = TilesLayer1((Map(X, Y).TileRefNum) + 2).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).Y = TilesLayer1((Map(X, Y).TileRefNum) + 2).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).X = TilesLayer1((Map(X, Y).TileRefNum) + 3).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).Y = TilesLayer1((Map(X, Y).TileRefNum) + 3).Y + YPixelDiff
        'End If
        End If
    Next Y
    Next X

    End If
    

    If TileCounter2 >= 0 Then
    'reset counter
    TheCounter = 0

    For X = XPixelDiff / 32 To (XPixelDiff / 32) + (XPixelDiff / 32) + Int(PicLevel.Width / 32) + 1
    For Y = YPixelDiff / 32 To (YPixelDiff / 32) + (YPixelDiff / 32) + Int(PicLevel.Height / 32) + 1
    
        If Map(X, Y).Used = True And Map(X, Y).Walk = True Then
        Rendered = Rendered + 1

        TilesLayer2((Map(X, Y).TileRefNum)).X = TilesLayer2((Map(X, Y).TileRefNum)).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum)).Y = TilesLayer2((Map(X, Y).TileRefNum)).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).X = TilesLayer2((Map(X, Y).TileRefNum) + 1).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).Y = TilesLayer2((Map(X, Y).TileRefNum) + 1).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).X = TilesLayer2((Map(X, Y).TileRefNum) + 2).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).Y = TilesLayer2((Map(X, Y).TileRefNum) + 2).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).X = TilesLayer2((Map(X, Y).TileRefNum) + 3).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).Y = TilesLayer2((Map(X, Y).TileRefNum) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(Map(X, Y).TextureRefNum)  '//Tell the device which texture we want to use...
        TheCounter = Map(X, Y).TileRefNum / 4
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer2(Map(X, Y).TileRefNum), Len(TilesLayer2(TheCounter))

        TilesLayer2((Map(X, Y).TileRefNum)).X = TilesLayer2((Map(X, Y).TileRefNum)).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum)).Y = TilesLayer2((Map(X, Y).TileRefNum)).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).X = TilesLayer2((Map(X, Y).TileRefNum) + 1).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).Y = TilesLayer2((Map(X, Y).TileRefNum) + 1).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).X = TilesLayer2((Map(X, Y).TileRefNum) + 2).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).Y = TilesLayer2((Map(X, Y).TileRefNum) + 2).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).X = TilesLayer2((Map(X, Y).TileRefNum) + 3).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).Y = TilesLayer2((Map(X, Y).TileRefNum) + 3).Y + YPixelDiff

        End If
    Next Y
    Next X
    
    End If
    

    'render layer3 objects that are above character befor character
    If TileCounter3 >= 0 Then
    'reset counter
    TheCounter = 0

    Do
    DoEvents
    If ((LayerTiles3((TheCounter * 4)).X <= XPixelDiff + 480 And LayerTiles3((TheCounter * 4)).X >= XPixelDiff) Or (LayerTiles3((TheCounter * 4) + 1).X <= XPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 1).X >= XPixelDiff)) Then
    If (LayerTiles3((TheCounter * 4)).Y <= YPixelDiff + 480 And LayerTiles3((TheCounter * 4)).Y >= YPixelDiff) Or (LayerTiles3((TheCounter * 4) + 2).Y <= YPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 2).Y >= YPixelDiff) Then
        
        'if objects are above character.
        If LayerTiles3((TheCounter * 4) + 3).Y < CharYPixelPos Then
        
        Rendered = Rendered + 1
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X - XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures3(TheCounter))  '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer3((TheCounter * 4)), Len(TilesLayer3(TheCounter))
        
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X + XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y + YPixelDiff
        
        End If
    End If
    End If
        TheCounter = TheCounter + 1
    Loop Until TheCounter > TileCounter3

    End If
    
    'render char
    D3DDevice.SetTexture 0, CharTexture(CharImage)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CharVertex(0), Len(CharVertex(0))
    

    'render layer3 objects that are below character after character
    If TileCounter3 >= 0 Then
    'reset counter
    TheCounter = 0

    Do
    DoEvents
    If ((LayerTiles3((TheCounter * 4)).X <= XPixelDiff + 480 And LayerTiles3((TheCounter * 4)).X >= XPixelDiff) Or (LayerTiles3((TheCounter * 4) + 1).X <= XPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 1).X >= XPixelDiff)) Then
    If (LayerTiles3((TheCounter * 4)).Y <= YPixelDiff + 480 And LayerTiles3((TheCounter * 4)).Y >= YPixelDiff) Or (LayerTiles3((TheCounter * 4) + 2).Y <= YPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 2).Y >= YPixelDiff) Then
        
        'if objects are below character.
        If LayerTiles3((TheCounter * 4) + 3).Y >= CharYPixelPos Then
        
        Rendered = Rendered + 1
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X - XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures3(TheCounter))  '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer3((TheCounter * 4)), Len(TilesLayer3(TheCounter))
        
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X + XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y + YPixelDiff
        
        End If
    End If
    End If
        TheCounter = TheCounter + 1
    Loop Until TheCounter > TileCounter3

    End If
    

    If TileCounter4 >= 0 Then
    'reset counter
    TheCounter = 0
        
    Do
    DoEvents
    If ((LayerTiles4((TheCounter * 4)).X <= XPixelDiff + 480 And LayerTiles4((TheCounter * 4)).X >= XPixelDiff) Or (LayerTiles4((TheCounter * 4) + 1).X <= XPixelDiff + 800 And LayerTiles4((TheCounter * 4) + 1).X >= XPixelDiff)) Then
    If ((LayerTiles4((TheCounter * 4)).Y <= YPixelDiff + 480 And LayerTiles4((TheCounter * 4)).Y >= YPixelDiff) Or (LayerTiles4((TheCounter * 4) + 2).Y <= YPixelDiff + 800 And LayerTiles4((TheCounter * 4) + 2).Y >= YPixelDiff)) Then
        
        Rendered = Rendered + 1
        TilesLayer4((TheCounter * 4)).X = LayerTiles4((TheCounter * 4)).X - XPixelDiff
        TilesLayer4((TheCounter * 4)).Y = LayerTiles4((TheCounter * 4)).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 1).X = LayerTiles4((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 1).Y = LayerTiles4((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 2).X = LayerTiles4((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 2).Y = LayerTiles4((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 3).X = LayerTiles4((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 3).Y = LayerTiles4((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures4(TheCounter)) '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer4((TheCounter * 4)), Len(TilesLayer4(TheCounter))
        
        TilesLayer4((TheCounter * 4)).X = LayerTiles4((TheCounter * 4)).X + XPixelDiff
        TilesLayer4((TheCounter * 4)).Y = LayerTiles4((TheCounter * 4)).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 1).X = LayerTiles4((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 1).Y = LayerTiles4((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 2).X = LayerTiles4((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 2).Y = LayerTiles4((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 3).X = LayerTiles4((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 3).Y = LayerTiles4((TheCounter * 4) + 3).Y + YPixelDiff
        
    End If
    End If
    TheCounter = TheCounter + 1
    Loop Until TheCounter > TileCounter4

    End If
    
    
    If TileCounter5 >= 0 Then
    'reset counter
    TheCounter = 0

    Do
    DoEvents

        TilesLayer5((TheCounter * 4)).X = LayerTiles5((TheCounter * 4)).X - XPixelDiff
        TilesLayer5((TheCounter * 4)).Y = LayerTiles5((TheCounter * 4)).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 1).X = LayerTiles5((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 1).Y = LayerTiles5((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 2).X = LayerTiles5((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 2).Y = LayerTiles5((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 3).X = LayerTiles5((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 3).Y = LayerTiles5((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures5(TheCounter))  '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer5((TheCounter * 4)), Len(TilesLayer5(TheCounter))
        
        TilesLayer5((TheCounter * 4)).X = LayerTiles5((TheCounter * 4)).X + XPixelDiff
        TilesLayer5((TheCounter * 4)).Y = LayerTiles5((TheCounter * 4)).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 1).X = LayerTiles5((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 1).Y = LayerTiles5((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 2).X = LayerTiles5((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 2).Y = LayerTiles5((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 3).X = LayerTiles5((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 3).Y = LayerTiles5((TheCounter * 4) + 3).Y + YPixelDiff
        TheCounter = TheCounter + 1
        
    
    Loop Until TheCounter > TileCounter5
    
    End If
    
    'if displaying text
    If GMode = 2 Then
        'render scroll
        D3DDevice.SetTexture 0, ScrollTexture
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ScrollVertex(0), Len(ScrollVertex(0))
    End If
    
    '//Draw the text
    TextRect.Top = 0
    TextRect.bottom = 20
    TextRect.Right = 75
    D3DX.DrawText MainFont, &HFFFFCC00, "test", TextRect, DT_TOP Or DT_LEFT
    

'if running in editor mode
Else
    'If Layer1 is to be rendered.
    If ChkLayer1.Value = 1 Then
        If TileCounter1 >= 0 Then
        
    For X = XPixelDiff / 32 To (XPixelDiff / 32) + (XPixelDiff / 32) + Int(PicLevel.Width / 32) + 1
    For Y = YPixelDiff / 32 To (YPixelDiff / 32) + (YPixelDiff / 32) + Int(PicLevel.Height / 32) + 1
        
        If Map(X, Y).Used = True And Map(X, Y).Walk = False Then
        Rendered = Rendered + 1
        'If X = 11 And Y = 4 Then
        TilesLayer1((Map(X, Y).TileRefNum)).X = TilesLayer1((Map(X, Y).TileRefNum)).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum)).Y = TilesLayer1((Map(X, Y).TileRefNum)).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).X = TilesLayer1((Map(X, Y).TileRefNum) + 1).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).Y = TilesLayer1((Map(X, Y).TileRefNum) + 1).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).X = TilesLayer1((Map(X, Y).TileRefNum) + 2).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).Y = TilesLayer1((Map(X, Y).TileRefNum) + 2).Y - YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).X = TilesLayer1((Map(X, Y).TileRefNum) + 3).X - XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).Y = TilesLayer1((Map(X, Y).TileRefNum) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(Map(X, Y).TextureRefNum)  '//Tell the device which texture we want to use...
        TheCounter = Map(X, Y).TileRefNum / 4
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer1(Map(X, Y).TileRefNum), Len(TilesLayer1(TheCounter))

        TilesLayer1((Map(X, Y).TileRefNum)).X = TilesLayer1((Map(X, Y).TileRefNum)).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum)).Y = TilesLayer1((Map(X, Y).TileRefNum)).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).X = TilesLayer1((Map(X, Y).TileRefNum) + 1).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 1).Y = TilesLayer1((Map(X, Y).TileRefNum) + 1).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).X = TilesLayer1((Map(X, Y).TileRefNum) + 2).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 2).Y = TilesLayer1((Map(X, Y).TileRefNum) + 2).Y + YPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).X = TilesLayer1((Map(X, Y).TileRefNum) + 3).X + XPixelDiff
        TilesLayer1((Map(X, Y).TileRefNum) + 3).Y = TilesLayer1((Map(X, Y).TileRefNum) + 3).Y + YPixelDiff
        'End If
        End If
    Next Y
    Next X

        End If
    End If
    
    'If Layer2 is to be rendered.
    If ChkLayer2.Value = 1 Then
        If TileCounter2 >= 0 Then


    For X = XPixelDiff / 32 To (XPixelDiff / 32) + (XPixelDiff / 32) + Int(PicLevel.Width / 32) + 1
    For Y = YPixelDiff / 32 To (YPixelDiff / 32) + (YPixelDiff / 32) + Int(PicLevel.Height / 32) + 1
    
        If Map(X, Y).Used = True And Map(X, Y).Walk = True Then
        Rendered = Rendered + 1

        TilesLayer2((Map(X, Y).TileRefNum)).X = TilesLayer2((Map(X, Y).TileRefNum)).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum)).Y = TilesLayer2((Map(X, Y).TileRefNum)).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).X = TilesLayer2((Map(X, Y).TileRefNum) + 1).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).Y = TilesLayer2((Map(X, Y).TileRefNum) + 1).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).X = TilesLayer2((Map(X, Y).TileRefNum) + 2).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).Y = TilesLayer2((Map(X, Y).TileRefNum) + 2).Y - YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).X = TilesLayer2((Map(X, Y).TileRefNum) + 3).X - XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).Y = TilesLayer2((Map(X, Y).TileRefNum) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(Map(X, Y).TextureRefNum)  '//Tell the device which texture we want to use...
        TheCounter = Map(X, Y).TileRefNum / 4
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer2(Map(X, Y).TileRefNum), Len(TilesLayer2(TheCounter))

        TilesLayer2((Map(X, Y).TileRefNum)).X = TilesLayer2((Map(X, Y).TileRefNum)).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum)).Y = TilesLayer2((Map(X, Y).TileRefNum)).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).X = TilesLayer2((Map(X, Y).TileRefNum) + 1).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 1).Y = TilesLayer2((Map(X, Y).TileRefNum) + 1).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).X = TilesLayer2((Map(X, Y).TileRefNum) + 2).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 2).Y = TilesLayer2((Map(X, Y).TileRefNum) + 2).Y + YPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).X = TilesLayer2((Map(X, Y).TileRefNum) + 3).X + XPixelDiff
        TilesLayer2((Map(X, Y).TileRefNum) + 3).Y = TilesLayer2((Map(X, Y).TileRefNum) + 3).Y + YPixelDiff

        End If
    Next Y
    Next X
    
        End If
    End If
    

    'If Layer3 is to be rendered.
    If ChkLayer3.Value = 1 Then
        If TileCounter3 >= 0 Then
        'reset counter
        TheCounter = 0

    Do
    DoEvents
    If ((LayerTiles3((TheCounter * 4)).X <= XPixelDiff + 480 And LayerTiles3((TheCounter * 4)).X >= XPixelDiff) Or (LayerTiles3((TheCounter * 4) + 1).X <= XPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 1).X >= XPixelDiff)) Then
    If (LayerTiles3((TheCounter * 4)).Y <= YPixelDiff + 480 And LayerTiles3((TheCounter * 4)).Y >= YPixelDiff) Or (LayerTiles3((TheCounter * 4) + 2).Y <= YPixelDiff + 800 And LayerTiles3((TheCounter * 4) + 2).Y >= YPixelDiff) Then

        Rendered = Rendered + 1
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X - XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures3(TheCounter))  '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer3((TheCounter * 4)), Len(TilesLayer3(TheCounter))
        
        TilesLayer3((TheCounter * 4)).X = LayerTiles3((TheCounter * 4)).X + XPixelDiff
        TilesLayer3((TheCounter * 4)).Y = LayerTiles3((TheCounter * 4)).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 1).X = LayerTiles3((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 1).Y = LayerTiles3((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 2).X = LayerTiles3((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 2).Y = LayerTiles3((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer3((TheCounter * 4) + 3).X = LayerTiles3((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer3((TheCounter * 4) + 3).Y = LayerTiles3((TheCounter * 4) + 3).Y + YPixelDiff
        
    End If
    End If
        TheCounter = TheCounter + 1
    Loop Until TheCounter > TileCounter3
    
        End If
    End If
    

    'If Layer4 is to be rendered.
    If ChkLayer4.Value = 1 Then
        If TileCounter4 >= 0 Then
        'reset counter
        TheCounter = 0
        
    Do
    DoEvents
    If ((LayerTiles4((TheCounter * 4)).X <= XPixelDiff + 480 And LayerTiles4((TheCounter * 4)).X >= XPixelDiff) Or (LayerTiles4((TheCounter * 4) + 1).X <= XPixelDiff + 800 And LayerTiles4((TheCounter * 4) + 1).X >= XPixelDiff)) Then
    If ((LayerTiles4((TheCounter * 4)).Y <= YPixelDiff + 480 And LayerTiles4((TheCounter * 4)).Y >= YPixelDiff) Or (LayerTiles4((TheCounter * 4) + 2).Y <= YPixelDiff + 800 And LayerTiles4((TheCounter * 4) + 2).Y >= YPixelDiff)) Then
        
        Rendered = Rendered + 1
        TilesLayer4((TheCounter * 4)).X = LayerTiles4((TheCounter * 4)).X - XPixelDiff
        TilesLayer4((TheCounter * 4)).Y = LayerTiles4((TheCounter * 4)).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 1).X = LayerTiles4((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 1).Y = LayerTiles4((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 2).X = LayerTiles4((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 2).Y = LayerTiles4((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer4((TheCounter * 4) + 3).X = LayerTiles4((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer4((TheCounter * 4) + 3).Y = LayerTiles4((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures4(TheCounter)) '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer4((TheCounter * 4)), Len(TilesLayer4(TheCounter))
        
        TilesLayer4((TheCounter * 4)).X = LayerTiles4((TheCounter * 4)).X + XPixelDiff
        TilesLayer4((TheCounter * 4)).Y = LayerTiles4((TheCounter * 4)).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 1).X = LayerTiles4((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 1).Y = LayerTiles4((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 2).X = LayerTiles4((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 2).Y = LayerTiles4((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer4((TheCounter * 4) + 3).X = LayerTiles4((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer4((TheCounter * 4) + 3).Y = LayerTiles4((TheCounter * 4) + 3).Y + YPixelDiff
        
    End If
    End If
    TheCounter = TheCounter + 1
    Loop Until TheCounter > TileCounter4
    
        End If
    End If
    
    

    'If Layer5 is to be rendered.
    If ChkLayer5.Value = 1 Then
        If TileCounter5 >= 0 Then
        'reset counter
        TheCounter = 0

    Do
    DoEvents

        TilesLayer5((TheCounter * 4)).X = LayerTiles5((TheCounter * 4)).X - XPixelDiff
        TilesLayer5((TheCounter * 4)).Y = LayerTiles5((TheCounter * 4)).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 1).X = LayerTiles5((TheCounter * 4) + 1).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 1).Y = LayerTiles5((TheCounter * 4) + 1).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 2).X = LayerTiles5((TheCounter * 4) + 2).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 2).Y = LayerTiles5((TheCounter * 4) + 2).Y - YPixelDiff
        TilesLayer5((TheCounter * 4) + 3).X = LayerTiles5((TheCounter * 4) + 3).X - XPixelDiff
        TilesLayer5((TheCounter * 4) + 3).Y = LayerTiles5((TheCounter * 4) + 3).Y - YPixelDiff
        
        D3DDevice.SetTexture 0, Textures(TheTextures5(TheCounter))  '//Tell the device which texture we want to use...
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TilesLayer5((TheCounter * 4)), Len(TilesLayer5(TheCounter))
        
        TilesLayer5((TheCounter * 4)).X = LayerTiles5((TheCounter * 4)).X + XPixelDiff
        TilesLayer5((TheCounter * 4)).Y = LayerTiles5((TheCounter * 4)).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 1).X = LayerTiles5((TheCounter * 4) + 1).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 1).Y = LayerTiles5((TheCounter * 4) + 1).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 2).X = LayerTiles5((TheCounter * 4) + 2).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 2).Y = LayerTiles5((TheCounter * 4) + 2).Y + YPixelDiff
        TilesLayer5((TheCounter * 4) + 3).X = LayerTiles5((TheCounter * 4) + 3).X + XPixelDiff
        TilesLayer5((TheCounter * 4) + 3).Y = LayerTiles5((TheCounter * 4) + 3).Y + YPixelDiff
        TheCounter = TheCounter + 1

    Loop Until TheCounter > TileCounter5
    
        End If
    End If
    
    'mouse cursor texture
    If Selected = True Then
        D3DDevice.SetTexture 0, CurrentTexture
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CurrentTile(0), Len(CurrentTile(0))
    End If
    

End If

D3DDevice.EndScene

'//3. Update the frame to the screen...
'       This is the same as the Primary.Flip method as used in DirectX 7
'       These values below should work for almost all cases...
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End If
End Sub

Public Sub SaveMap()
Dim Filename As String
Dim TheCounter As Long
Dim TheWorld As String
Dim TextureName As String
Dim i As Long
Dim X As Long
Dim x1 As Long
Dim Y As Long
Dim y1 As Long
Dim Temp As String
CommonDialog1.ShowSave
Filename = CommonDialog1.Filename
If Filename <> "" Then
Filename = Left(Filename, Len(Filename) - 4)
End If

'sort map from left to right
'Call Init2dTo1DSort
    


'if filename not empty
If Filename <> "" Then

'save all NPC's
For i = 0 To NPCCounter
    Open App.path & "/NPCS/" & NPC(i).Filename For Output As #1
        Print #1, NPC(i).Name
        Print #1, NPC(i).Dialogue1
        Print #1, NPC(i).Dialogue2
        Print #1, NPC(i).Dialogue3
        Print #1, NPC(i).Dialogue4
        Print #1, NPC(i).Dialogue5
        Print #1, NPC(i).Say1
        Print #1, NPC(i).Say2
        Print #1, NPC(i).Say3
        Print #1, NPC(i).Say4
        Print #1, NPC(i).Say5
    Close #1
Next i

'save all monsters
For i = 0 To MonsterCounter
    Open App.path & "/Monsters/" & Monster(i).Filename For Output As #1
        Print #1, Monster(i).Name
        Print #1, Monster(i).HP
        Print #1, Monster(i).MP
        Print #1, Monster(i).ATK
        Print #1, Monster(i).Def
        Print #1, Monster(i).Spd
        Print #1, Monster(i).Evade
        Print #1, Monster(i).Accuracy
        Print #1, Monster(i).DropChance
        Print #1, Monster(i).SpawnRate
        Print #1, Monster(i).Item
        Print #1, Monster(i).Value
        Print #1, Monster(i).Exp
    Close #1
Next i

'if map type = mmorpg
If TheMapType = 1 Then



    'save each sector
    For X = SectorXOffset To MaxXSector
        For Y = SectorYOffset To MaxYSector

    'save start of map file with settings
    Open X & "-" & Y & ".NWM" For Output As #1
    
    Print #1, TheMapType & vbNewLine & MapTyper.Name & vbNewLine & MapTyper.pass & vbNewLine & MapTyper.music & vbNewLine & MapTyper.description

    Close #1
        Next Y
    Next X



'Save Layer1 if Layer1 has more then 0 tiles
If TileCounter1 >= 0 Then
TheCounter = 0


For x1 = SectorXOffset To MaxXSector
    For y1 = SectorYOffset To MaxYSector
        'open file for output
        Open x1 & "-" & y1 & "a.NWM" For Output As #1
            
            For X = ((x1 - 1) * 14) + x1 - 1 To (x1 * 14) + x1 - 1
            For Y = ((y1 - 1) * 14) + y1 - 1 To (y1 * 14) + y1 - 1
            If Map(X, Y).Used = True And Map(X, Y).Walk = False Then

                TextureName = Replace(TheTextures(Map(X, Y).TextureRefNum), App.path & "\tiles\", "")
                
                Print #1, TextureName

                Print #1, Map(X, Y).X - ((x1 - SectorXOffset) * 480)
                Print #1, Map(X, Y).Y - ((y1 - SectorYOffset) * 480)
                Print #1, Map(X, Y).trigger
                Print #1, Map(X, Y).color

            End If
            Next Y
            Next X
            TheCounter = TheCounter + 1
        'close file
        Close #1
    Next y1
Next x1
    
End If

'Save Layer2 if Layer2 has more then 0 tiles
If TileCounter2 >= 0 Then


For x1 = SectorXOffset To MaxXSector
    For y1 = SectorYOffset To MaxYSector
        'open file for output
        Open x1 & "-" & y1 & "b.NWM" For Output As #1
            
            For X = ((x1 - 1) * 14) + x1 - 1 To (x1 * 14) + x1 - 1
            For Y = ((y1 - 1) * 14) + y1 - 1 To (y1 * 14) + y1 - 1
            If Map(X, Y).Used = True And Map(X, Y).Walk = True Then
            
                TextureName = Replace(TheTextures(Map(X, Y).TextureRefNum), App.path & "\tiles\", "")
        
                Print #1, TextureName

                Temp = Map(X, Y).X - ((x1 - SectorXOffset) * 480)
                If Temp < 0 Then
                    MsgBox 1
                End If
                
                Print #1, Map(X, Y).X - ((x1 - SectorXOffset) * 480)
                Print #1, Map(X, Y).Y - ((y1 - SectorYOffset) * 480)
                Print #1, Map(X, Y).trigger
                Print #1, Map(X, Y).color

            End If
            Next Y
            Next X
            TheCounter = TheCounter + 1
        'close file
        Close #1
    Next y1
Next x1
    
    
End If
'Save Layer3 if Layer3 has more then 0 tiles
If TileCounter3 >= 0 Then



For X = SectorXOffset To MaxXSector
    For Y = SectorYOffset To MaxYSector
        'open file for output
        Open X & "-" & Y & "c.NWM" For Output As #1
            
            For i = 0 To TileCounter3

    
            If (LayerTiles3(i * 4).XSector = X) And (LayerTiles3(i * 4).YSector = Y) Then
                
                TextureName = Replace(TheTextures(TheTextures3(i)), App.path & "\tiles\", "")
        
                Print #1, TextureName
                
                
                Print #1, (LayerTiles3(i * 4).X - ((X - SectorXOffset) * 480))
                Print #1, (LayerTiles3(i * 4).Y - ((Y - SectorYOffset) * 480))
                Print #1, LayerTiles3(i * 4).trigger
                Print #1, LayerTiles3(i * 4).color

            End If
            
            
            Next i
        'close file
        Close #1
    Next Y
Next X


End If
'Save Layer4 if layer4 has more then 0 tiles
If TileCounter4 >= 0 Then
TheWorld = ""


For X = SectorXOffset To MaxXSector
    For Y = SectorYOffset To MaxYSector
        'open file for output
        Open X & "-" & Y & "d.NWM" For Output As #1
            For i = 0 To TileCounter4

    
            If (LayerTiles4(i * 4).XSector = X) And (LayerTiles4(i * 4).YSector = Y) Then
                TextureName = Replace(TheTextures(TheTextures4(i)), App.path & "\tiles\", "")

                Print #1, TextureName

                Print #1, (LayerTiles4(i * 4).X - ((X - SectorXOffset) * 480))
                Print #1, (LayerTiles4(i * 4).Y - ((Y - SectorYOffset) * 480))
                Print #1, LayerTiles4(i * 4).trigger
                Print #1, LayerTiles4(i * 4).color

            End If
            
            
            Next i
        'close file
        Close #1
    Next Y
Next X

End If
'Save Layer5 if Layer5 has more then 0 tiles
If TileCounter5 >= 0 Then



For X = SectorXOffset To MaxXSector
    For Y = SectorYOffset To MaxYSector
        'open file for output
        Open X & "-" & Y & "e.NWM" For Output As #1
        
            For i = 0 To TileCounter5

    
            If (LayerTiles5(i * 4).XSector = X) And (LayerTiles5(i * 4).YSector = Y) Then
                
                TextureName = Replace(TheTextures(TheTextures5(i)), App.path & "\tiles\", "")
        
                Print #1, TextureName

                Print #1, (LayerTiles5(i * 4).X - ((X - SectorXOffset) * 480))
                Print #1, (LayerTiles5(i * 4).Y - ((Y - SectorYOffset) * 480))
                Print #1, LayerTiles5(i * 4).trigger
                Print #1, LayerTiles5(i * 4).color

            End If
            
            
            Next i
        'close file
        Close #1
    Next Y
Next X
    
    Save = True
End If










'if map type = battle scene
ElseIf TheMapType = 2 Then


    'save start of map file with settings
    Open Filename & ".NWM" For Output As #1
    
    Print #1, TheMapType & vbNewLine & MapTyper.Name & vbNewLine & MapTyper.pass & vbNewLine & MapTyper.music & vbNewLine & MapTyper.description

    Close #1
    
'Save Layer1 if Layer1 has more then 0 tiles
If TileCounter1 >= 0 Then



Open Filename & "a.NWM" For Output As #1
For X = SectorXOffset - 1 To (MaxXSector * 14) + MaxXSector
    For Y = SectorYOffset - 1 To (MaxYSector * 14) + MaxYSector
        'open file for output

            
            If Map(X, Y).Used = True And Map(X, Y).Walk = False Then
            
                TextureName = Replace(TheTextures(Map(X, Y).TextureRefNum), App.path & "\tiles\", "")
        
                Print #1, TextureName
                
                Print #1, Map(X, Y).X
                Print #1, Map(X, Y).Y
                Print #1, Map(X, Y).trigger
                Print #1, Map(X, Y).color

            End If
            TheCounter = TheCounter + 1

    Next Y
Next X
'close file
Close #1

    
End If

'Save Layer2 if Layer2 has more then 0 tiles
If TileCounter2 >= 0 Then


'open file for output
Open Filename & "b.NWM" For Output As #1
For X = SectorXOffset - 1 To (MaxXSector * 14) + MaxXSector
    For Y = SectorYOffset - 1 To MaxYSector * 14 + MaxYSector

            
            If Map(X, Y).Used = True And Map(X, Y).Walk = True Then
            
                TextureName = Replace(TheTextures(Map(X, Y).TextureRefNum), App.path & "\tiles\", "")
        
                Print #1, TextureName
                
                Print #1, Map(X, Y).X
                Print #1, Map(X, Y).Y
                Print #1, Map(X, Y).trigger
                Print #1, Map(X, Y).color

            End If
            TheCounter = TheCounter + 1

    Next Y
Next X
'close file
Close #1
    
    
End If
'Save Layer3 if Layer3 has more then 0 tiles
If TileCounter3 >= 0 Then




'open file for output
Open Filename & "c.NWM" For Output As #1
            
    For i = 0 To TileCounter3

        TextureName = Replace(TheTextures(TheTextures3(i)), App.path & "\tiles\", "")
        
        Print #1, TextureName
                
                
        Print #1, (LayerTiles3(i * 4).X)
        Print #1, (LayerTiles3(i * 4).Y)
        Print #1, LayerTiles3(i * 4).trigger
        Print #1, LayerTiles3(i * 4).color

            
    Next i
'close file
Close #1



End If
'Save Layer4 if layer4 has more then 0 tiles
If TileCounter4 >= 0 Then
TheWorld = ""



'open file for output
Open Filename & "d.NWM" For Output As #1
            
    For i = 0 To TileCounter4

        TextureName = Replace(TheTextures(TheTextures4(i)), App.path & "\tiles\", "")
        
        Print #1, TextureName

        Print #1, (LayerTiles4(i * 4).X)
        Print #1, (LayerTiles4(i * 4).Y)
        Print #1, LayerTiles4(i * 4).trigger
        Print #1, LayerTiles4(i * 4).color

    Next i
    
'close file
Close #1


End If

'Save Layer5 if Layer5 has more then 0 tiles
If TileCounter5 >= 0 Then




'open file for output
Open Filename & "e.NWM" For Output As #1
            
    For i = 0 To TileCounter5
    
                
        TextureName = Replace(TheTextures(TheTextures5(i)), App.path & "\tiles\", "")
        
        Print #1, TextureName

        Print #1, (LayerTiles5(i * 4).X)
        Print #1, (LayerTiles5(i * 4).Y)
        Print #1, LayerTiles5(i * 4).trigger
        Print #1, LayerTiles5(i * 4).color

             
    Next i
'close file
Close #1



End If

End If
End If
End Sub

Public Sub Set_Monster(Name As String, HP As Integer, MP As Integer, ATK As Integer, Def As Integer, Spd As Integer, Evade As Integer, Accuracy As Integer, DropChance As Integer, SpawnRate As Integer, Item As Integer, Value As Integer, Exp As Integer, Filename As String)
    MonsterCounter = MonsterCounter + 1
    ReDim Preserve Monster(MonsterCounter)
    
    Monster(MonsterCounter).Name = Name
    Monster(MonsterCounter).HP = HP
    Monster(MonsterCounter).MP = MP
    Monster(MonsterCounter).ATK = ATK
    Monster(MonsterCounter).Def = Def
    Monster(MonsterCounter).Spd = Spd
    Monster(MonsterCounter).Evade = Evade
    Monster(MonsterCounter).Accuracy = Accuracy
    Monster(MonsterCounter).DropChance = DropChance
    Monster(MonsterCounter).SpawnRate = SpawnRate
    Monster(MonsterCounter).Item = Item
    Monster(MonsterCounter).Value = Value
    Monster(MonsterCounter).Exp = Exp
    Monster(MonsterCounter).Filename = Filename

End Sub

Public Sub Set_NPC(Name As String, Dial1 As String, Dial2 As String, Dial3 As String, Dial4 As String, Dial5 As String, Say1 As String, Say2 As String, Say3 As String, Say4 As String, Say5 As String, Filename As String)
    NPCCounter = NPCCounter + 1
    ReDim Preserve NPC(NPCCounter)
    NPC(NPCCounter).Name = Name
    NPC(NPCCounter).Dialogue1 = Dial1
    NPC(NPCCounter).Dialogue2 = Dial2
    NPC(NPCCounter).Dialogue3 = Dial3
    NPC(NPCCounter).Dialogue4 = Dial4
    NPC(NPCCounter).Dialogue5 = Dial5
    NPC(NPCCounter).Say1 = Say1
    NPC(NPCCounter).Say2 = Say2
    NPC(NPCCounter).Say3 = Say3
    NPC(NPCCounter).Say4 = Say4
    NPC(NPCCounter).Say4 = Say5
    NPC(NPCCounter).Filename = Filename

End Sub


Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


Public Sub Init2dTo1DSort()
  Dim i As Long

  If TileCounter1 >= 0 Then
  'sort layer 1
  ReDim Preserve lArray(TileCounter1)
  For i = 0 To TileCounter1
    lArray(i) = (TilesLayer1(i * 4).X * 65536) Or TilesLayer1(i * 4).Y
  Next i

  QuickSortAscending lArray, 0, TileCounter1, 1
  
  End If
  
  
  If TileCounter2 >= 0 Then
  'sort layer 2
  ReDim Preserve lArray(TileCounter2)
  For i = 0 To TileCounter2
    lArray(i) = (TilesLayer2(i * 4).X * 65536) Or TilesLayer2(i * 4).Y
  Next i


  QuickSortAscending lArray, 0, TileCounter2, 2
  End If
  
  
  

End Sub

' Modified vbnet.mvps.org function
Public Sub QuickSortAscending(lArray() As Long, inLow As Long, inHi As Long, layer As Integer)
  Dim pivot   As Long
  Dim tmpSwap As Long
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  Dim tmpTLV As TLVERTEX
  Dim tmpTile As TileStorage
  Dim tmpNum As Long
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = lArray((inLow + inHi) \ 2)

  Do While (tmpLow <= tmpHi)
    Do While (lArray(tmpLow) < pivot And tmpLow < inHi)
      tmpLow = tmpLow + 1
    Loop
   
    Do While (pivot < lArray(tmpHi) And tmpHi > inLow)
      tmpHi = tmpHi - 1
    Loop

    If (tmpLow <= tmpHi) Then
      tmpSwap = lArray(tmpLow)
      lArray(tmpLow) = lArray(tmpHi)
      lArray(tmpHi) = tmpSwap
      
    'if layer = 1
    If layer = 1 Then
        tmpTile = LayerTiles1(tmpLow * 4)
        tmpNum = TheTextures1(tmpLow)
        TheTextures1(tmpLow) = TheTextures1(tmpHi)
        TheTextures1(tmpHi) = tmpNum


        LayerTiles1(tmpLow * 4) = LayerTiles1(tmpHi * 4)
        LayerTiles1(tmpHi * 4) = tmpTile
        
        tmpTile = LayerTiles1(tmpLow * 4 + 1)
        LayerTiles1(tmpLow * 4 + 1) = LayerTiles1(tmpHi * 4 + 1)
        LayerTiles1(tmpHi * 4 + 1) = tmpTile
        
        tmpTile = LayerTiles1(tmpLow * 4 + 2)
        LayerTiles1(tmpLow * 4 + 2) = LayerTiles1(tmpHi * 4 + 2)
        LayerTiles1(tmpHi * 4 + 2) = tmpTile
        
        tmpTile = LayerTiles1(tmpLow * 4 + 3)
        LayerTiles1(tmpLow * 4 + 3) = LayerTiles1(tmpHi * 4 + 3)
        LayerTiles1(tmpHi * 4 + 3) = tmpTile


        tmpTLV = TilesLayer1(tmpLow * 4)
        TilesLayer1(tmpLow * 4) = TilesLayer1(tmpHi * 4)
        TilesLayer1(tmpHi * 4) = tmpTLV
        
        tmpTLV = TilesLayer1(tmpLow * 4 + 1)
        TilesLayer1(tmpLow * 4 + 1) = TilesLayer1(tmpHi * 4 + 1)
        TilesLayer1(tmpHi * 4 + 1) = tmpTLV
        
        tmpTLV = TilesLayer1(tmpLow * 4 + 2)
        TilesLayer1(tmpLow * 4 + 2) = TilesLayer1(tmpHi * 4 + 2)
        TilesLayer1(tmpHi * 4 + 2) = tmpTLV
        
        tmpTLV = TilesLayer1(tmpLow * 4 + 3)
        TilesLayer1(tmpLow * 4 + 3) = TilesLayer1(tmpHi * 4 + 3)
        TilesLayer1(tmpHi * 4 + 3) = tmpTLV
        

    End If
      
      'if layer = 2
    If layer = 2 Then
        tmpTile = LayerTiles2(tmpLow * 4)
        tmpNum = TheTextures2(tmpLow)
        TheTextures2(tmpLow) = TheTextures2(tmpHi)
        TheTextures2(tmpHi) = tmpNum


        LayerTiles2(tmpLow * 4) = LayerTiles2(tmpHi * 4)
        LayerTiles2(tmpHi * 4) = tmpTile
        
        tmpTile = LayerTiles2(tmpLow * 4 + 1)
        LayerTiles2(tmpLow * 4 + 1) = LayerTiles2(tmpHi * 4 + 1)
        LayerTiles2(tmpHi * 4 + 1) = tmpTile
        
        tmpTile = LayerTiles2(tmpLow * 4 + 2)
        LayerTiles2(tmpLow * 4 + 2) = LayerTiles2(tmpHi * 4 + 2)
        LayerTiles2(tmpHi * 4 + 2) = tmpTile
        
        tmpTile = LayerTiles2(tmpLow * 4 + 3)
        LayerTiles2(tmpLow * 4 + 3) = LayerTiles2(tmpHi * 4 + 3)
        LayerTiles2(tmpHi * 4 + 3) = tmpTile


        tmpTLV = TilesLayer2(tmpLow * 4)
        TilesLayer2(tmpLow * 4) = TilesLayer2(tmpHi * 4)
        TilesLayer2(tmpHi * 4) = tmpTLV
        
        tmpTLV = TilesLayer2(tmpLow * 4 + 1)
        TilesLayer2(tmpLow * 4 + 1) = TilesLayer2(tmpHi * 4 + 1)
        TilesLayer2(tmpHi * 4 + 1) = tmpTLV
        
        tmpTLV = TilesLayer2(tmpLow * 4 + 2)
        TilesLayer2(tmpLow * 4 + 2) = TilesLayer2(tmpHi * 4 + 2)
        TilesLayer2(tmpHi * 4 + 2) = tmpTLV
        
        tmpTLV = TilesLayer2(tmpLow * 4 + 3)
        TilesLayer2(tmpLow * 4 + 3) = TilesLayer2(tmpHi * 4 + 3)
        TilesLayer2(tmpHi * 4 + 3) = tmpTLV
    End If
    
    
      'if layer = 3
    If layer = 3 Then
        tmpTile = LayerTiles3(tmpLow * 4)
        tmpNum = TheTextures3(tmpLow)
        TheTextures3(tmpLow) = TheTextures3(tmpHi)
        TheTextures3(tmpHi) = tmpNum


        LayerTiles3(tmpLow * 4) = LayerTiles3(tmpHi * 4)
        LayerTiles3(tmpHi * 4) = tmpTile
        
        tmpTile = LayerTiles3(tmpLow * 4 + 1)
        LayerTiles3(tmpLow * 4 + 1) = LayerTiles3(tmpHi * 4 + 1)
        LayerTiles3(tmpHi * 4 + 1) = tmpTile
        
        tmpTile = LayerTiles3(tmpLow * 4 + 2)
        LayerTiles3(tmpLow * 4 + 2) = LayerTiles3(tmpHi * 4 + 2)
        LayerTiles3(tmpHi * 4 + 2) = tmpTile
        
        tmpTile = LayerTiles3(tmpLow * 4 + 3)
        LayerTiles3(tmpLow * 4 + 3) = LayerTiles3(tmpHi * 4 + 3)
        LayerTiles3(tmpHi * 4 + 3) = tmpTile


        tmpTLV = TilesLayer3(tmpLow * 4)
        TilesLayer3(tmpLow * 4) = TilesLayer3(tmpHi * 4)
        TilesLayer3(tmpHi * 4) = tmpTLV
        
        tmpTLV = TilesLayer3(tmpLow * 4 + 1)
        TilesLayer3(tmpLow * 4 + 1) = TilesLayer3(tmpHi * 4 + 1)
        TilesLayer3(tmpHi * 4 + 1) = tmpTLV
        
        tmpTLV = TilesLayer3(tmpLow * 4 + 2)
        TilesLayer3(tmpLow * 4 + 2) = TilesLayer3(tmpHi * 4 + 2)
        TilesLayer3(tmpHi * 4 + 2) = tmpTLV
        
        tmpTLV = TilesLayer3(tmpLow * 4 + 3)
        TilesLayer3(tmpLow * 4 + 3) = TilesLayer3(tmpHi * 4 + 3)
        TilesLayer3(tmpHi * 4 + 3) = tmpTLV
    End If



      'if layer = 4
    If layer = 4 Then
        tmpTile = LayerTiles4(tmpLow * 4)
        tmpNum = TheTextures4(tmpLow)
        TheTextures4(tmpLow) = TheTextures4(tmpHi)
        TheTextures4(tmpHi) = tmpNum


        LayerTiles4(tmpLow * 4) = LayerTiles4(tmpHi * 4)
        LayerTiles4(tmpHi * 4) = tmpTile
        
        tmpTile = LayerTiles4(tmpLow * 4 + 1)
        LayerTiles4(tmpLow * 4 + 1) = LayerTiles4(tmpHi * 4 + 1)
        LayerTiles4(tmpHi * 4 + 1) = tmpTile
        
        tmpTile = LayerTiles4(tmpLow * 4 + 2)
        LayerTiles4(tmpLow * 4 + 2) = LayerTiles4(tmpHi * 4 + 2)
        LayerTiles4(tmpHi * 4 + 2) = tmpTile
        
        tmpTile = LayerTiles4(tmpLow * 4 + 3)
        LayerTiles4(tmpLow * 4 + 3) = LayerTiles4(tmpHi * 4 + 3)
        LayerTiles4(tmpHi * 4 + 3) = tmpTile


        tmpTLV = TilesLayer4(tmpLow * 4)
        TilesLayer4(tmpLow * 4) = TilesLayer4(tmpHi * 4)
        TilesLayer4(tmpHi * 4) = tmpTLV
        
        tmpTLV = TilesLayer4(tmpLow * 4 + 1)
        TilesLayer4(tmpLow * 4 + 1) = TilesLayer4(tmpHi * 4 + 1)
        TilesLayer4(tmpHi * 4 + 1) = tmpTLV
        
        tmpTLV = TilesLayer4(tmpLow * 4 + 2)
        TilesLayer4(tmpLow * 4 + 2) = TilesLayer4(tmpHi * 4 + 2)
        TilesLayer4(tmpHi * 4 + 2) = tmpTLV
        
        tmpTLV = TilesLayer4(tmpLow * 4 + 3)
        TilesLayer4(tmpLow * 4 + 3) = TilesLayer4(tmpHi * 4 + 3)
        TilesLayer4(tmpHi * 4 + 3) = tmpTLV
    End If


      'if layer = 5
    If layer = 5 Then
        tmpTile = LayerTiles5(tmpLow * 4)
        tmpNum = TheTextures5(tmpLow)
        TheTextures5(tmpLow) = TheTextures5(tmpHi)
        TheTextures5(tmpHi) = tmpNum


        LayerTiles5(tmpLow * 4) = LayerTiles5(tmpHi * 4)
        LayerTiles5(tmpHi * 4) = tmpTile
        
        tmpTile = LayerTiles5(tmpLow * 4 + 1)
        LayerTiles5(tmpLow * 4 + 1) = LayerTiles5(tmpHi * 4 + 1)
        LayerTiles5(tmpHi * 4 + 1) = tmpTile
        
        tmpTile = LayerTiles5(tmpLow * 4 + 2)
        LayerTiles5(tmpLow * 4 + 2) = LayerTiles5(tmpHi * 4 + 2)
        LayerTiles5(tmpHi * 4 + 2) = tmpTile
        
        tmpTile = LayerTiles5(tmpLow * 4 + 3)
        LayerTiles5(tmpLow * 4 + 3) = LayerTiles5(tmpHi * 4 + 3)
        LayerTiles5(tmpHi * 4 + 3) = tmpTile


        tmpTLV = TilesLayer5(tmpLow * 4)
        TilesLayer5(tmpLow * 4) = TilesLayer5(tmpHi * 4)
        TilesLayer5(tmpHi * 4) = tmpTLV
        
        tmpTLV = TilesLayer5(tmpLow * 4 + 1)
        TilesLayer5(tmpLow * 4 + 1) = TilesLayer5(tmpHi * 4 + 1)
        TilesLayer5(tmpHi * 4 + 1) = tmpTLV
        
        tmpTLV = TilesLayer5(tmpLow * 4 + 2)
        TilesLayer5(tmpLow * 4 + 2) = TilesLayer5(tmpHi * 4 + 2)
        TilesLayer5(tmpHi * 4 + 2) = tmpTLV
        
        tmpTLV = TilesLayer5(tmpLow * 4 + 3)
        TilesLayer5(tmpLow * 4 + 3) = TilesLayer5(tmpHi * 4 + 3)
        TilesLayer5(tmpHi * 4 + 3) = tmpTLV
    End If
    
    
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
      

    End If
  Loop
    
  If (inLow < tmpHi) Then QuickSortAscending lArray(), inLow, tmpHi, layer
  If (tmpLow < inHi) Then QuickSortAscending lArray(), tmpLow, inHi, layer
End Sub

Function LoadTiles(path As String, SearchStr As String, FileCount As Long, DirCount As Integer)
    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net
    'URL: http://www.allapi.net/

    Dim Filename As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    Dim nodX As Node
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    
    DoEvents
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                LoadTiles = LoadTiles + LoadTiles(path & dirNames(nDir - 1) & "\", SearchStr, FileCount, DirCount)
                ReDim Preserve dirNames(nDir)
                
                Set nodX = Tree.Nodes.Add(, , dirNames(nDir - 1), dirNames(nDir - 1))
                For i = 0 To Nqueues
                    
                    If Queues(i).word <> "" Then
                    If Right(Queues(i).word, 3) <> ".db" Then
 
                    Set nodX = Tree.Nodes.Add(dirNames(nDir - 1), tvwChild, Queues(i).word & FileCount, Queues(i).word)
                    End If
                    End If
                Next i
                Nqueues = 0
            End If

        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") Then
                LoadTiles = LoadTiles + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                Nqueues = Nqueues + 1
                ReDim Preserve Queues(Nqueues)
                Queues(Nqueues).word = Filename
                
                If InStr(path & Filename, ".") Then
                'List1.AddItem path & FileName
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If

End Function


Public Sub UnsetVars()
    SetTrigger = False
    SetNPC = False
    SetMonster = False
    SetColor = False
    DoDelete = False
    Locked = False
End Sub

Private Sub chkGame_Click()

    PicSizer.Picture = LoadPicture(App.path & "\VAMP.BMP")
    
    'reset char pixel positions
    CharXPixelPos = PicLevel.Width / 2 + (PicSizer.Width / 2)
    CharYPixelPos = PicLevel.Height / 2 + (PicSizer.Height / 2) + 15
    
    'reset char matrix co-ords
    CharXPos = 7
    CharYPos = 9
    
    'reset pixel diffs
    XPixelDiff = 0
    YPixelDiff = 0

    'reset scroll bar pos
    VScroll.Value = 0
    HScroll.Value = 0
    
    'set focus on level screen
    PicLevel.SetFocus
    
    'reset char image
    CharImage = 7
    CharStep = 1
    
    'if running in editor mode, turn to game mode
    If GMode = 0 Then
        GMode = 1
    'turn to editor mode
    Else
        GMode = 0
    End If
    
    If Timer.Enabled = True Then
        Timer.Enabled = False
    Else
        Timer.Enabled = True
    End If
End Sub

Private Sub Form_Load()
Dim LastTick As Long
Dim CurrentTick As Long
Dim Ticks As Long
Dim TempTileCount As Long
Dim TempSessionMinutes As String
Dim X As String
Dim i As Integer
Dim Time1 As String 'time user started editor
Dim Time2 As String 'keep track of current time
Selected = False
MouseIsDown = False
ColorKeyVal = &HFF00FF ' pink
XPixelDiff = 0
YPixelDiff = 0
ValRed = 255
ValGreen = 255
ValBlue = 255
GMode = 0

'load map editor settings
LoadXSectors = ReadIni("Options", "XSectors", "", App.path & "/NWEditor.ini")
LoadYSectors = ReadIni("Options", "YSectors", "", App.path & "/NWEditor.ini")

'Load Hours/Minutes user worked with editor.
OverallHours = ReadIni("Time", "Hours", "", App.path & "/NWEditor.ini")
OverallMinutes = ReadIni("Time", "Minutes", "", App.path & "/NWEditor.ini")

TileCounter1 = -1
TileCounter2 = -1
TileCounter3 = -1
TileCounter4 = -1
TileCounter5 = -1

TextureCounter = -1
TextureCounter1 = -1
TextureCounter2 = -1
TextureCounter3 = -1
TextureCounter4 = -1
TextureCounter5 = -1

NPCCounter = -1
MonsterCounter = -1
SetTrigger = False
Locked = False

SectorXOffset = 1
SectorYOffset = 1
MaxXSector = 1
MaxYSector = 1

CharXPos = 7
CharYPos = 9

SessionHours = 0
TempSessionMinutes = 0
ReDim Queues(0)

Me.Show '//Make sure our window is visible
Me.Caption = "Loading Tiles..."
Call LoadTiles(App.path & "/tiles", "*.*", 0, 0)



bRunning = Initialise()
Debug.Print "Device Creation Return Code : ", bRunning 'So you can see what happens...

'if command loading
If Len(Command) > 0 Then
    X = Left(Command, Len(Command) - 4)
    Active = True
    Save = True
    Call LoadMap(X)
End If


'Get Current Tick's (FPS)
LastTick = Int(GetTickCount() / 1000)
'get time user started editor
Time1 = Time
Do While bRunning
    'Get Current Tick
    Time2 = Time
    SessionMinutes = DateDiff("n", Time1, Time2)


    'update status panels with overall hrs and mins as well
    'as session hrs and mins
    StatusBar.Panels(4).Text = "Session: " & SessionHours & "hrs" & " " & SessionMinutes & "mins"
    StatusBar.Panels(5).Text = "Overall: " & OverallHours & "hrs" & " " & OverallMinutes & "mins"


    'if old session not new session (if a minute has passed)
    If TempSessionMinutes <> SessionMinutes Then
            TempSessionMinutes = SessionMinutes
            OverallMinutes = OverallMinutes + 1
    End If
    
    
    'if overall minutes = 60 add another hour to overall hour
    If OverallMinutes > 59 Then
        OverallHours = OverallHours + 1
        OverallMinutes = -1
    End If
    

    'if session minutes over 59 add hours. reset minute counter
    If SessionMinutes > 60 Then
        SessionHours = SessionHours + 1
        'reset time
        Time1 = Time
    End If
    

    
    CurrentTick = Int(GetTickCount() / 1000)
    If Int(CurrentTick) <= Int(LastTick) Then
        Ticks = Ticks + 1
    Else
        TempTileCount = TileCounter1 + TileCounter2 + TileCounter3 + TileCounter4 + TileCounter5 + 5
        frmMain.Caption = "NWE FPS = " & Ticks & "   Lay1 = " & TileCounter1 + 1 & "   Lay2 = " & TileCounter2 + 1 & "   Lay3 = " & TileCounter3 + 1 & "   Lay4 = " & TileCounter4 + 1 & "   Total Tiles = " & TempTileCount & " - Rendered = " & Rendered
        LastTick = CurrentTick
        Ticks = 0
    End If
    
    'if editor is active window
    If GetActiveWindow() = Me.hWnd Then
        Render '//Update the frame...
    End If
    
    DoEvents '//Allow windows time to think; otherwise you'll get into a really tight (and bad) loop...
Loop '//Begin the next frame...

'save overall minutes and hours used
Call WriteIni("Time", "Hours", Str(OverallHours), App.path & "/NWEditor.ini")
Call WriteIni("Time", "Minutes", Str(OverallMinutes), App.path & "/NWEditor.ini")

'//If we've gotten to this point the loop must have been terminated
'   So we need to clean up after ourselves. This isn't essential, but it'
'   good coding practise.

On Error Resume Next 'If the objects were never created;
'                               (the initialisation failed) we might get an
'                               error when freeing them... which we need to
'                               handle, but as we're closing anyway...
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing

'clean up textures
For i = 0 To TextureCounter
    Set Textures(i) = Nothing
Next i

'clean up char textures
For i = 0 To 11
    Set CharTexture(i) = Nothing
Next i

Set ScrollTexture = Nothing

Debug.Print "All Objects Destroyed"
'//Final termination:
Unload Me
End
End Sub



Private Function InitialiseGeometry() As Boolean
    
    
On Error GoTo BailOut: '//Setup our Error handler

        
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\1.BMP")
        Set CharTexture(0) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\1.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\2.BMP")
        Set CharTexture(1) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\2.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\3.BMP")
        Set CharTexture(2) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\3.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\4.BMP")
        Set CharTexture(3) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\4.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\5.BMP")
        Set CharTexture(4) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\5.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\6.BMP")
        Set CharTexture(5) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\6.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\7.BMP")
        Set CharTexture(6) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\7.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\8.BMP")
        Set CharTexture(7) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\8.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\9.BMP")
        Set CharTexture(8) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\9.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\10.BMP")
        Set CharTexture(9) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\10.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\11.BMP")
        Set CharTexture(10) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\11.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        PicSizer.Picture = LoadPicture(App.path & "\images\char\12.BMP")
        Set CharTexture(11) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\char\12.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        
        'create char's vertex
        CharVertex(0) = CreateTLVertex(PicLevel.Width / 2 - (PicSizer.Width / 2), PicLevel.Height / 2 - (PicSizer.Height / 2) + 15, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
        CharVertex(1) = CreateTLVertex(PicLevel.Width / 2 + (PicSizer.Width / 2), PicLevel.Height / 2 - (PicSizer.Height / 2) + 15, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
        CharVertex(2) = CreateTLVertex(PicLevel.Width / 2 - (PicSizer.Width / 2), PicLevel.Height / 2 + (PicSizer.Height / 2) + 15, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
        CharVertex(3) = CreateTLVertex(PicLevel.Width / 2 + (PicSizer.Width / 2), PicLevel.Height / 2 + (PicSizer.Height / 2) + 15, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)
        
        'set char pixel positions
        CharXPixelPos = PicLevel.Width / 2 + (PicSizer.Width / 2)
        CharYPixelPos = PicLevel.Height / 2 + (PicSizer.Height / 2) + 15
        
        
        PicSizer.Picture = LoadPicture(App.path & "\images\scroll3.bmp")
        Set ScrollTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\images\scroll3.bmp", PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        
        
        'create scrolls's vertex
        ScrollVertex(0) = CreateTLVertex(63.5, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
        ScrollVertex(1) = CreateTLVertex(PicLevel.Width - 63.5, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
        ScrollVertex(2) = CreateTLVertex(63.5, 0 + PicSizer.Height, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
        ScrollVertex(3) = CreateTLVertex(PicLevel.Width - 63.5, 0 + PicSizer.Height, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)
        
        
InitialiseGeometry = True
Exit Function
BailOut:
InitialiseGeometry = False
End Function



'//This is just a simple wrapper function that makes filling the structures much much easier...
Private Function CreateTLVertex(X As Single, Y As Single, z As Single, rhw As Single, color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
CreateTLVertex.X = X
CreateTLVertex.Y = Y
CreateTLVertex.z = z
CreateTLVertex.rhw = rhw
CreateTLVertex.color = color
CreateTLVertex.specular = specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function

Private Sub Form_Resize()

'reset control placement properties

'The Controls are arranged on the form on how much room would be
'needed for the treeview and preview pic while keeping it with
'a reasonable dimensioned form.
'the rest of the controls are then based upon that.
'when moving to fullscreen mode. The width of the treeview and tilepreview
'stay the same while the rest of the controls adjuct accordinally.


'Treeview = FormHeight - PreviewPicHeight
Tree.Height = frmMain.ScaleHeight - PicTilePreview.Height - StatusBar.Height - chkGame.Height

'top of check game mode check box
chkGame.Top = Tree.Height

'Top Of PreviewPic = FormHeight - PreviewPicHeight
PicTilePreview.Top = frmMain.ScaleHeight - PicTilePreview.Height - StatusBar.Height


'VerticleScroll is height of form
VScroll.Height = frmMain.ScaleHeight - StatusBar.Height


'Top Horizontal Scroll = height of form - Horizontal Scroll Height
HScroll.Top = frmMain.ScaleHeight - HScroll.Height - StatusBar.Height
'Height Of Horitontal Scroll = FormWidth - TreeView Width - Vertical Scroll Width
HScroll.Width = frmMain.ScaleWidth - Tree.Width - VScroll.Width


'Height of Level Picture = FormHeight - Horizontal Scroll Height
PicLevel.Height = frmMain.ScaleHeight - HScroll.Height - ChkLayer1.Height - Option1.Height - StatusBar.Height - 10
'Width Of Level Picture is the same length of the Horizontal Scroll Bar
PicLevel.Width = HScroll.Width

PicSizer.Top = frmMain.ScaleHeight + 10
'end of setting control properties

End Sub

Private Sub Form_Unload(Cancel As Integer)
    bRunning = False
End Sub

Private Sub HScroll_Change()
XPixelDiff = (Round(HScroll.Value / 32, 0)) * 32
PicLevel.SetFocus
End Sub





















Private Sub mnuChangeColors_Click()

    Load frmColor
    'if no object/tile selected
    If LCase(Right(App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4)) <> ".bmp" Then
            Call frmColor.SetVars(0, 0, 0, 0, 0, 0, 0, 0, App.path & "\tiles\tree\dwarftree01.bmp", 0, 185, 218)
    Else
        Call frmColor.SetVars(0, 0, 0, 0, 0, 0, 0, 0, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 0, PicSizer.Width, PicSizer.Height)
    End If
    
    frmColor.Show , frmMain
    Locked = True

End Sub

Private Sub mnuColorsResetColors_Click()
    ValRed = 255
    ValGreen = 255
    ValBlue = 255
End Sub

Private Sub mnuFileExit_Click()
Dim X As Long

'if user is active and didnt save
If Active = True And Save = False Then
X = MsgBox("Would you like to save first?", vbYesNo, "Save first?")

    'if user wants to save.
    If X = 8 Then
    Call SaveMap
    End If
End If

bRunning = False
    
End Sub

Private Sub mnuFileLoad_Click()

Dim X As Long
Dim Name As String
Dim ThePath As String
'if user is active and didnt save

If Active = True And Save = False Then

X = MsgBox("Are you sure you want load a map without saving first?", vbYesNoCancel, "Continue without saving?")

    Select Case X
        'if user cancles
        Case 2:
            Exit Sub
        'if user says no
        Case 6:
        'get filename to open
        
        CommonDialog1.ShowOpen

        'load map
        Call LoadMap(CommonDialog1.Filename)
        'if user says yes
        Case 7:
            Call SaveMap
        'get filename to open
        CommonDialog1.ShowOpen
        'load map
        Call LoadMap(CommonDialog1.Filename)
 
        
    End Select



Else
    CommonDialog1.ShowOpen
    
    Call LoadMap(CommonDialog1.Filename)
End If


End Sub







Private Sub mnuFileNew_Click()
Dim X As Long
'if user is active and havnt saved map
If Active = True And Save = False Then

    'ask user if he wants to start a new map first without saving
    X = MsgBox("Are you sure you want to start a new map without saving?", vbYesNoCancel, "Start a new map")
    Select Case X
        'if user cancles
        Case 2:
            Exit Sub
        'if user says no
        Case 6:
            Call NewMap
        'if user says yes
        Case 7:
            Call SaveMap
            Call NewMap
            
    End Select

'else
Else
    Call NewMap
End If

End Sub

Private Sub mnuFileSave_Click()
    Call SaveMap
End Sub







Private Sub mnuMapAutoFillCustom_Click()
Load frmAutoFill
frmAutoFill.Show , frmMain
End Sub

Private Sub mnuMapAutoFillScreen_Click()
    Call AutoFill(PicLevel.Height - 32, PicLevel.Width - 32, 0)
End Sub

Private Sub mnuMapAutoGenCustom_Click()
Load frmAutoGen
frmAutoGen.Show , frmMain
End Sub

Private Sub mnuMapAutoGenScreen_Click()
Call AutoGenerate("Mixgrass1viengrass", PicLevel.Height - 32, PicLevel.Width - 32, 0)
End Sub



Private Sub mnuMapClearLayerLayer_Click(Index As Integer)
TileCounter1 = -1
TileCounter2 = -1
TileCounter3 = -1
TileCounter4 = -1
TileCounter5 = -1

TextureCounter = -1
TextureCounter1 = -1
TextureCounter2 = -1
TextureCounter3 = -1
TextureCounter4 = -1
TextureCounter5 = -1


End Sub

Private Sub mnuMapProperties_Click()
frmMapProperties.Show , Me
End Sub

Private Sub mnuNPCSInsertNPC_Click()
frmNPCEditor.Show , frmMain
End Sub

Private Sub mnuScriptingArena_Click()
frmArena.Show , Me
End Sub

Private Sub mnuScriptingDetails_Click()
frmScript.Show , Me
End Sub

Private Sub mnuViewOptions_Click()
frmOptions.Show , Me
End Sub

Private Sub PicLevel_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyR Then
        SetTrigger = True
    End If
    If KeyCode = vbKeyN Then
        SetNPC = True
    End If
    If KeyCode = vbKeyM Then
        SetMonster = True
    End If
    If KeyCode = vbKeyC Then
        SetColor = True
    End If
    
    If KeyCode = vbKey1 Then
        Layer1Move = True
    End If
    
    If KeyCode = vbKey2 Then
        Layer2Move = True
    End If
    
    'if running in editor mode
    If GMode = 0 Then
    
        If KeyCode = vbKeyLeft Then
            'if scroll bar great or equal to 32. dont want to go lower then 0
            If HScroll.Value >= 32 Then
                'decrase scroll bar value
                HScroll.Value = HScroll.Value - 32
            End If
        End If
        
        If KeyCode = vbKeyRight Then
            'if scroll bar great or equal to 32736. dont want to go above 32736
            If HScroll.Value <= 32736 Then
                'increase scroll bar value
                HScroll.Value = HScroll.Value + 32
            End If
        End If
        
        
        If KeyCode = vbKeyDown Then
            'if scroll bar great or equal to 32736. dont want to go above 32736
            If VScroll.Value <= 32736 Then
                'increase scroll bar value
                VScroll.Value = VScroll.Value + 32
            End If
        End If
        
        If KeyCode = vbKeyUp Then
            'if scroll bar great or equal to 32. dont want to go lower then 0
            If VScroll.Value >= 32 Then
                'decrease scroll bar value
                VScroll.Value = VScroll.Value - 32
            End If
        End If
        
        
    End If
    
End Sub

Private Sub PicLevel_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyR Then
        SetTrigger = False
    End If
    If KeyCode = vbKeyN Then
        SetNPC = False
    End If
    If KeyCode = vbKeyM Then
        SetMonster = False
    End If
    If KeyCode = vbKeyC Then
        SetColor = False
    End If
    If KeyCode = vbKeyE Then
        DoEdit = False
    End If
End Sub


Private Sub PicLevel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim GoGood As Boolean
Dim TilePlacer As Long
Dim i As Long
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single
GoGood = True

'if running in editor mode
If chkGame.Value = 0 Then


If Selected = True Then
Save = False
Active = True


    If Locked = False Then
    'Mouse Is Down. Used for MouseMove
    MouseIsDown = True
    
    'avoid placing tiles if off to left of level area
    If X < 0 Then
        Exit Sub
    End If

    'avoid placing tiles if off to right of level area
    If X >= PicLevel.Width Then
        Exit Sub
    End If

    'avoid placing tiles if above level area
    If Y < 0 Then
        Exit Sub
    End If

    'avoid placing tiles if below level area
    If Y >= PicLevel.Height Then
        Exit Sub
    End If

    'adjust pixels by difference of original area
    X = X + XPixelDiff
    Y = Y + YPixelDiff


    'place tile (left click button)
    If Button = 1 Then

    'because I am rounding the pixels to the nearest 32. there
    'are some problems with certain pixel co-ords rounding

    'if the distance between x's are less then 32. increase X
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32))) < 32 Then
        X = X + 1
    End If

    'if the distance between y's are less then 32. increase Y
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32))) < 32 Then
        Y = Y + 1
    End If


    'if the distance between x's are > 64. decrease x
    '1 pixel reduces the rounding to 32
    If (((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32)) >= 64 Then
        X = X - 1
    End If

    'if the distance between y's are > 64. decrease y
    '1 pixel reduces the rounding to 32
    If (((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32)) >= 64 Then
        Y = Y - 1
    End If


    'if moving tile to layer1
    If Layer1Move = True Then
        Call MoveLayer(X, Y, 1)
        Layer1Move = False
        Exit Sub
    End If
    
    'if moving tile to layer2
    If Layer2Move = True Then
        Call MoveLayer(X, Y, 2)
        Layer2Move = False
        Exit Sub
    End If
    
    'if TileLayer is 5,
    If Option5.Value = True Then



    'set the 4 corners of the square
    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)

    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)
    
    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)

    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)
    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, "", RGB(ValRed, ValGreen, ValBlue))



    'tileLayer is 4(Layer we are working with)
    ElseIf Option4.Value = True Then



    'Set the 4 Corners of the square

    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)


    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)


    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)


    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)

    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, "", RGB(ValRed, ValGreen, ValBlue))

    
    'tileLayer is 3(Layer we are working with)
    ElseIf Option3.Value = True Then
    

    'Set the 4 Corners of the square

    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)

    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)

    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)

    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)
    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Exit Sub
        Locked = True
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, "", RGB(ValRed, ValGreen, ValBlue))

    
    'tileLayer is 2(Layer we are working with)
    ElseIf Option2.Value = True Then
    

    'set the 4 corners of the square
    x1 = (Round((X - 16) / 32, 0)) * 32
    y1 = (Round((Y - 16) / 32, 0)) * 32

    x2 = (Round((X + 16) / 32, 0)) * 32
    y2 = (Round((Y - 16) / 32, 0)) * 32


    x3 = (Round((X - 16) / 32, 0)) * 32
    y3 = (Round((Y + 16) / 32, 0)) * 32

    x4 = (Round((X + 16) / 32, 0)) * 32
    y4 = (Round((Y + 16) / 32, 0)) * 32

    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Exit Sub
        Locked = True
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, "", RGB(ValRed, ValGreen, ValBlue))



    'tileLayer is 1(Layer we are working with)
    ElseIf Option1.Value = True Then

    

    'set the 4 corners of the square
    x1 = (Round((X - 16) / 32, 0)) * 32
    y1 = (Round((Y - 16) / 32, 0)) * 32

    x2 = (Round((X + 16) / 32, 0)) * 32
    y2 = (Round((Y - 16) / 32, 0)) * 32

    x3 = (Round((X - 16) / 32, 0)) * 32
    y3 = (Round((Y + 16) / 32, 0)) * 32

    x4 = (Round((X + 16) / 32, 0)) * 32
    y4 = (Round((Y + 16) / 32, 0)) * 32

    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, "", RGB(ValRed, ValGreen, ValBlue))


    'end of tilelayers
    End If


    'right click (tile delete)
    Else
    
        If Option5.Value = True Then
            Call DeleteTile(X, Y, 5)
        ElseIf Option4.Value = True Then
            Call DeleteTile(X, Y, 4)
        ElseIf Option3.Value = True Then
            Call DeleteTile(X, Y, 3)
        ElseIf Option2.Value = True Then
            Call DeleteTile(X, Y, 2)
        ElseIf Option1.Value = True Then
            Call DeleteTile(X, Y, 1)
        End If
        
    'end of tile delete
    End If
    
    End If
    
'end selected
End If

'end of running editor mode
End If

End Sub

Private Sub PicLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'//We're going to move the transparent sprite around using the mouse
'   doing this can demonstrate how transparencies work....
Dim TilePlacer As Long
Dim TileLayer As Integer
Dim i As Long
Dim z As Long
Dim GoGood As Boolean
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single


StatusBar.Panels(1).Text = "Sector:: " & Int((X + XPixelDiff) / 480) + SectorXOffset & "," & Int((Y + YPixelDiff) / 480) + SectorYOffset
StatusBar.Panels(2).Text = "Screen: " & (Round((X) / 32, 0) + 1) & "," & (Round((Y) / 32, 0) + 1)
StatusBar.Panels(3).Text = "World: " & (Round(((X) + XPixelDiff) / 32, 0) + 1) & "," & (Round(((Y) + YPixelDiff) / 32, 0) + 1)



'if running in editor mode
If chkGame.Value = 0 Then

'if a tile is selected
If Selected = True Then
    'store our cursor tile to mouse co-ords
    'shows where next tile will be placed
    CurrentTile(0) = CreateTLVertex(X - (PicSizer.Width / 2), Y - (PicSizer.Height / 2), 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(X + (PicSizer.Width / 2), Y - (PicSizer.Height / 2), 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(X - (PicSizer.Width / 2), Y + (PicSizer.Height / 2), 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(X + (PicSizer.Width / 2), Y + (PicSizer.Height / 2), 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)
End If

'The Mouse Is Down
If MouseIsDown = True Then

    If Locked = False Then
    'Mouse Is Down. Used for MouseMove
    MouseIsDown = True
    
    'avoid placing tiles if off to left of level area
    If X < 0 Then
        Exit Sub
    End If

    'avoid placing tiles if off to right of level area
    If X >= PicLevel.Width Then
        Exit Sub
    End If

    'avoid placing tiles if above level area
    If Y < 0 Then
        Exit Sub
    End If

    'avoid placing tiles if below level area
    If Y >= PicLevel.Height Then
        Exit Sub
    End If


    'place tile (left click button)
    If Button = 1 Then
    'good to set a tile
    GoGood = True


    'adjust pixels by difference of original area
    X = X + XPixelDiff
    Y = Y + YPixelDiff
    

    'because I am rounding the pixels to the nearest 32. there
    'are some problems with certain pixel co-ords rounding

    'if the distance between x's are less then 32. increase X
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32))) < 32 Then
        X = X + 1
    End If

    'if the distance between y's are less then 32. increase Y
    'distance will never be lower then 31. so adding 1 pixel is fine
    If ((((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32))) < 32 Then
        Y = Y + 1
    End If


    'if the distance between x's are > 64. decrease x
    '1 pixel reduces the rounding to 32
    If (((Round((X + 16) / 32, 0)) * 32) - ((Round((X - 16) / 32, 0)) * 32)) >= 64 Then
        X = X - 1
    End If

    'if the distance between y's are > 64. decrease y
    '1 pixel reduces the rounding to 32
    If (((Round((Y + 16) / 32, 0)) * 32) - ((Round((Y - 16) / 32, 0)) * 32)) >= 64 Then
        Y = Y - 1
    End If



    

    'if TileLayer is 5,
    If Option5.Value = True Then




    'set the 4 corners of the square
    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)

    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)
    
    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)

    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)
    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 5, "", RGB(ValRed, ValGreen, ValBlue))



    'tileLayer is 4(Layer we are working with)
    ElseIf Option4.Value = True Then



    'Set the 4 Corners of the square

    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)


    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)


    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)


    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)

    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 4, "", RGB(ValRed, ValGreen, ValBlue))


    
    'tileLayer is 3(Layer we are working with)
    ElseIf Option3.Value = True Then
    


    'Set the 4 Corners of the square

    x1 = X - (PicSizer.Width / 2)
    y1 = Y - (PicSizer.Height / 2)

    x2 = X + (PicSizer.Width / 2)
    y2 = Y - (PicSizer.Height / 2)

    x3 = X - (PicSizer.Width / 2)
    y3 = Y + (PicSizer.Height / 2)

    x4 = X + (PicSizer.Width / 2)
    y4 = Y + (PicSizer.Height / 2)
    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 3, "", RGB(ValRed, ValGreen, ValBlue))


    
    'tileLayer is 2(Layer we are working with)
    ElseIf Option2.Value = True Then
    
    

    'set the 4 corners of the square
    x1 = (Round((X - 16) / 32, 0)) * 32
    y1 = (Round((Y - 16) / 32, 0)) * 32

    x2 = (Round((X + 16) / 32, 0)) * 32
    y2 = (Round((Y - 16) / 32, 0)) * 32


    x3 = (Round((X - 16) / 32, 0)) * 32
    y3 = (Round((Y + 16) / 32, 0)) * 32

    x4 = (Round((X + 16) / 32, 0)) * 32
    y4 = (Round((Y + 16) / 32, 0)) * 32

    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 2, "", RGB(ValRed, ValGreen, ValBlue))


    
    'tileLayer is 1(Layer we are working with)
    ElseIf Option1.Value = True Then

    

    'set the 4 corners of the square
    x1 = (Round((X - 16) / 32, 0)) * 32
    y1 = (Round((Y - 16) / 32, 0)) * 32

    x2 = (Round((X + 16) / 32, 0)) * 32
    y2 = (Round((Y - 16) / 32, 0)) * 32

    x3 = (Round((X - 16) / 32, 0)) * 32
    y3 = (Round((Y + 16) / 32, 0)) * 32

    x4 = (Round((X + 16) / 32, 0)) * 32
    y4 = (Round((Y + 16) / 32, 0)) * 32
    
    'call trigger form, transfer co-ords, layer, texture to new form
    If SetTrigger = True Then
        frmTriggers.Show , frmMain
        Locked = True
        Call frmTriggers.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1)
        Exit Sub
    End If
    'call NPC form, transfer co-ords, layer, texture to new form
    If SetNPC = True Then
        frmNPCEditor.Show , frmMain
        Locked = True
        Call frmNPCEditor.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        Exit Sub
    End If
    'call Monster form, transfer co-ords, layer, texture to new form
    If SetMonster = True Then
        Load frmMonster
        Call frmMonster.StoreTrigger(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        frmMonster.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'if change colors
    If SetColor = True Then
        Load frmColor
        Call frmColor.SetVars(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, PicSizer.Width, PicSizer.Height)
        frmColor.Show , frmMain
        Locked = True
        Exit Sub
    End If
    'place tile
    Call PlaceTile(x1, x2, x3, x4, y1, y2, y3, y4, App.path & "\tiles\" & Tree.SelectedItem.FullPath, 1, "", RGB(ValRed, ValGreen, ValBlue))

    'end of tilelayers
    End If
    
    'right click (tile delete)
    Else
    
    
    'end of tile delete
    End If
    
    End If
End If

'end of running in editor mode
End If
End Sub


Private Sub PicLevel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Mouse Is Up. No longer place tiles while MouseMove
MouseIsDown = False
End Sub





Private Sub PicTilePreview_Click()
Dim stuff As TileStorage
ReDim Preserve stuff.Objects(0)
stuff.Objects(0).X = 2
MsgBox stuff.Objects(0).X
GMode = 2
End Sub


Private Sub Timer_Timer()

    Call GetKeys

End Sub

Private Sub Tree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo eih
'if map editing is active
If Active = True Then

    'if left button
    If Button = 1 Then
    
    If Right(Tree.SelectedItem.Text, 4) = ".bmp" Or Right(Tree.SelectedItem.Text, 4) = ".BMP" Then
        PicTilePreview.Picture = LoadPicture(App.path & "\tiles\" & Tree.SelectedItem.FullPath)
        PicSizer.Picture = LoadPicture(App.path & "\tiles\" & Tree.SelectedItem.FullPath)
        Set CurrentTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\tiles\" & Tree.SelectedItem.FullPath, PicSizer.Width, PicSizer.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
        Selected = True
        PicLevel.SetFocus
    End If

    'if right button
    ElseIf Button = 2 Then
        PopupMenu mnuMapAutoFill
    End If
    
End If
eih:
End Sub


Private Sub VScroll_Change()
YPixelDiff = (Round(VScroll.Value / 32, 0)) * 32
PicLevel.SetFocus
End Sub


