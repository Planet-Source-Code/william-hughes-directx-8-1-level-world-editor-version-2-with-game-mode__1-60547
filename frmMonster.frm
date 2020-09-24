VERSION 5.00
Begin VB.Form frmMonster 
   Caption         =   "Monster Editor"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmMonster.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbLoadMonster 
      Caption         =   "Load Monster"
      Height          =   375
      Left            =   4800
      TabIndex        =   52
      Top             =   6960
      Width           =   3015
   End
   Begin VB.HScrollBar HScrollEvade 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   50
      Top             =   2400
      Width           =   3015
   End
   Begin VB.HScrollBar HScrollMP 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   47
      Top             =   960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monster Attacks"
      Height          =   1335
      Left            =   120
      TabIndex        =   41
      Top             =   5520
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox lstAttack 
         Height          =   840
         Left            =   2640
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox lstAttacks 
         Height          =   840
         ItemData        =   "frmMonster.frx":08CA
         Left            =   120
         List            =   "frmMonster.frx":092E
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.HScrollBar HScrollAccuracy 
      Height          =   255
      Left            =   1200
      Max             =   10000
      TabIndex        =   39
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1920
      TabIndex        =   37
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   35
      Top             =   240
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change Monster Color"
      Height          =   2295
      Left            =   4800
      TabIndex        =   24
      Top             =   4560
      Width           =   3015
      Begin VB.HScrollBar HScrollBlue 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   27
         Top             =   1800
         Value           =   255
         Width           =   2175
      End
      Begin VB.HScrollBar HScrollGreen 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   26
         Top             =   1200
         Value           =   255
         Width           =   2175
      End
      Begin VB.HScrollBar HScrollRed 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   25
         Top             =   600
         Value           =   255
         Width           =   2175
      End
      Begin VB.Label lblBlue 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblGreen 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblRed 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Blue"
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
         TabIndex        =   30
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Green"
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
         TabIndex        =   29
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Red"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.HScrollBar HScrollHP 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   11
      Top             =   600
      Width           =   3015
   End
   Begin VB.HScrollBar HScrollATK 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   10
      Top             =   1320
      Width           =   3015
   End
   Begin VB.HScrollBar HScrollDEF 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   9
      Top             =   1680
      Width           =   3015
   End
   Begin VB.HScrollBar HScrollSPD 
      Height          =   255
      Left            =   960
      Max             =   1000
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtDropItemChance 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtSpawnRate 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton cmbSaveMonster 
      Caption         =   "Save Monster"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   4455
   End
   Begin VB.PictureBox PicMonster 
      Height          =   4215
      Left            =   4800
      ScaleHeight     =   4155
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblEvade 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   51
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Evade:"
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
      Left            =   240
      TabIndex        =   49
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblMP 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   48
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "MP:"
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
      Left            =   240
      TabIndex        =   46
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblAccuracy 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   40
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accuracy:"
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
      Left            =   240
      TabIndex        =   38
      Top             =   2760
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Monster Filename:"
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
      Left            =   240
      TabIndex        =   36
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
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
      Left            =   240
      TabIndex        =   34
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATK:"
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
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEF:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPD:"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Item Chance 1 Out Of:"
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
      Left            =   240
      TabIndex        =   20
      Top             =   3360
      Width           =   2385
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spawn Rate In Seconds:"
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
      Left            =   240
      TabIndex        =   19
      Top             =   3720
      Width           =   2130
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
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
      Left            =   240
      TabIndex        =   18
      Top             =   4080
      Width           =   435
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp To Give:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label lblHP 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblATK 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblDEF 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSPD 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   330
   End
End
Attribute VB_Name = "frmMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Dx As DirectX8 'The master Object, everything comes from here
Dim D3D As Direct3D8 'This controls all things 3D
Dim D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Dim bRunning As Boolean 'Controls whether the program is running or not...
Dim D3DX As D3DX8 '//A helper library
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

Dim CurrentTile(0 To 3) As TLVERTEX '//This is going to square that follows the mouse...

Dim CurrentTexture As Direct3DTexture8

Dim MonsterMode As Integer 'if new monster or editing previous monster. 1 = new, 0 = edit
Dim Width1 As Single
Dim Height1 As Single
Dim Texture As String


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

D3DWindow.hDeviceWindow = PicMonster.hWnd

'//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
'//See the lesson text for more information on this line...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, PicMonster.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                        D3DWindow)

'//Set the vertex shader to use our vertex format
D3DDevice.SetVertexShader FVF

'//Transformed and lit vertices dont need lighting
'   so we disable it...
D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True


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



Private Function InitialiseGeometry() As Boolean
    
    
On Error GoTo BailOut: '//Setup our Error handler

'//NOTE THAT WE ARE PASSING VALUES FOR THE tu AND tv ARGUMENTS



'## SECOND SQUARE ##


InitialiseGeometry = True
Exit Function
BailOut:
InitialiseGeometry = False
End Function


Public Sub Render()
'//1. We need to clear the render device before we can draw anything
'       This must always happen before you start rendering stuff...
Dim TheCounter As String
On Error Resume Next
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0 '//Clear the screen black

'//2. Rendering the graphics...

D3DDevice.BeginScene
    'All rendering calls go between these two lines
    

        D3DDevice.SetTexture 0, CurrentTexture
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CurrentTile(0), Len(CurrentTile(0))
    
D3DDevice.EndScene

'//3. Update the frame to the screen...
'       This is the same as the Primary.Flip method as used in DirectX 7
'       These values below should work for almost all cases...
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub


Public Sub StoreTrigger(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture1 As String, layer As Integer, Width2 As Single, Height2 As Single)
    Width1 = Width2
    Height1 = Height2
    Texture = Texture1
    Triggerer.x1 = x1
    Triggerer.x2 = x2
    Triggerer.x3 = x3
    Triggerer.x4 = x4
    Triggerer.y1 = y1
    Triggerer.y2 = y2
    Triggerer.y3 = y3
    Triggerer.y4 = y4
    Triggerer.layer = layer
    Triggerer.Texture = Texture1
End Sub

Public Sub EditMonster(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture1 As String, layer As Integer, Width2 As Single, Height2 As Single, Name As String, HP As Integer, Str As Integer, Def As Integer, Spd As Integer, Magi As Integer, Agress As Integer, Behav As Integer, DropChance As Integer, SpawnRate As Integer, Item As Integer, Value As Integer, Exp As Integer, Filename As Integer)
    Width1 = Width2
    Height1 = Height2
    Texture = Texture1
    Triggerer.x1 = x1
    Triggerer.x2 = x2
    Triggerer.x3 = x3
    Triggerer.x4 = x4
    Triggerer.y1 = y1
    Triggerer.y2 = y2
    Triggerer.y3 = y3
    Triggerer.y4 = y4
    Triggerer.layer = layer
    Triggerer.Texture = Texture1
    
    txtName = Name
    HScrollHP.Value = HP
    HScrollDEF.Value = Def
    HScrollSPD.Value = Spd
    'cmbBehavior.Text = Behav
    txtDropItemChance = DropChance
    txtSpawnRate = SpawnRate
    cmbItem.Text = Item
    txtValue = Value
    txtExp = Exp
    txtFilename = Filename
    
End Sub
Private Sub Command1_Click()

End Sub

Private Sub cmbSaveMonster_Click()
'if monster name and filename blank
If txtName = "" Or txtFilename = "" Then
    MsgBox "NPC must have a name and filename", vbCritical, "Monster Error"
'if monster name filename not blank
Else
    'if file doesnt exit
    If Dir(App.path & "/Monsters/" & txtFilename & ".MON") = "" Then
        'check if filename stored in memory
        If frmMain.MonCheck(txtFilename) = False Then
            Call frmMain.Set_Monster(txtName, HScrollHP.Value, HScrollMP.Value, HScrollATK.Value, HScrollDEF.Value, HScrollSPD.Value, HScrollEvade.Value, HScrollAccuracy.Value, txtDropItemChance.Text, txtSpawnRate.Text, cmbItem.Text, txtValue.Text, txtExp.Text, txtFilename & ".MON")
            Call frmMain.PlaceTile(Triggerer.x1, Triggerer.x2, Triggerer.x3, Triggerer.x4, Triggerer.y1, Triggerer.y2, Triggerer.y3, Triggerer.y4, Triggerer.Texture, Triggerer.layer, txtFilename & ".Mon", RGB(ValRed, ValGreen, ValBlue))
            Unload Me
        Else
            MsgBox "Monster filename allready taken. Choose another Monster filename", vbCritical, " Monster Error"
        End If
    
    'if file exist
    Else
        MsgBox "Monster filename allready taken. Choose another Monster filename", vbCritical, "Monster Error"
    End If
End If
End Sub


Private Sub Form_Activate()
bRunning = Initialise()
HScrollRed.Value = ValRed
HScrollBlue.Value = ValBlue
HScrollGreen.Value = ValGreen

Set CurrentTexture = D3DX.CreateTextureFromFileEx(D3DDevice, Texture, Width1, Height1, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF00FF, ByVal 0, ByVal 0)
CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

Do While bRunning
    Render '//Update the frame...
    DoEvents '//Allow windows time to think; otherwise you'll get into a really tight (and bad) loop...
Loop '//Begin the next frame...
Me.Caption = bRunning
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
Set CurrentTexture = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call frmMain.UnsetVars
    bRunning = False
End Sub



Private Sub HScrollAccuracy_Change()
lblAccuracy.Caption = HScrollAccuracy.Value
End Sub

Private Sub HScrollATK_Change()
lblATK.Caption = HScrollATK.Value
End Sub

Private Sub HScrollBlue_Change()
lblBlue.Caption = HScrollBlue.Value
ValBlue = HScrollBlue.Value

    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollDEF_Change()
lblDEF.Caption = HScrollDEF.Value
End Sub

Private Sub HScrollEvade_Change()
lblEvade.Caption = HScrollEvade.Value
End Sub

Private Sub HScrollGreen_Change()
lblGreen.Caption = HScrollGreen.Value
ValGreen = HScrollGreen.Value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollHP_Change()
lblHP.Caption = HScrollHP.Value
End Sub



Private Sub HScrollMP_Change()
lblMP.Caption = HScrollMP.Value
End Sub

Private Sub HScrollRed_Change()
lblRed.Caption = HScrollRed.Value
ValRed = HScrollRed.Value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollSPD_Change()
lblSPD.Caption = HScrollSPD.Value
End Sub











