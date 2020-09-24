VERSION 5.00
Begin VB.Form frmNPCEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NPC Editor"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   Icon            =   "frmNPCEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbLoadNPC 
      Caption         =   "Load NPC"
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   6960
      Width           =   4815
   End
   Begin VB.ListBox lstVillage 
      Height          =   1230
      ItemData        =   "frmNPCEditor.frx":08CA
      Left            =   8160
      List            =   "frmNPCEditor.frx":08EF
      TabIndex        =   40
      Top             =   6120
      Width           =   2535
   End
   Begin VB.ListBox lstMissions 
      Height          =   5130
      Left            =   8160
      TabIndex        =   38
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change NPC Color"
      Height          =   2295
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   3015
      Begin VB.HScrollBar HScrollRed 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   31
         Top             =   600
         Value           =   255
         Width           =   2175
      End
      Begin VB.HScrollBar HScrollGreen 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   30
         Top             =   1200
         Value           =   255
         Width           =   2175
      End
      Begin VB.HScrollBar HScrollBlue 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   29
         Top             =   1800
         Value           =   255
         Width           =   2175
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
         TabIndex        =   37
         Top             =   240
         Width           =   1935
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
         TabIndex        =   36
         Top             =   960
         Width           =   1935
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
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblRed 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblGreen 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblBlue 
         Caption         =   "255"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.PictureBox PicNPC 
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   2955
      TabIndex        =   27
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmbSetNPC 
      Caption         =   "Set NPC"
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   6480
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtNPCFilename 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtSay5 
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtDialogue5 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txtSay4 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtDialogue4 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtSay3 
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtDialogue3 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtSay2 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtDialogue2 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtSay1 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtDialogue1 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "(ex: George234)"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "NPC Filename:"
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
         TabIndex        =   23
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Say:"
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
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Dialogue 5:"
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
         TabIndex        =   10
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Say:"
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
         TabIndex        =   9
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Dialogue 4:"
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
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Say:"
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
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Dialogue 3:"
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
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Say:"
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
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Dialogue 2:"
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Say:"
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
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Dialogue 1:"
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
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Village:"
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
      Left            =   8160
      TabIndex        =   41
      Top             =   5880
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Missions:"
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
      Left            =   8160
      TabIndex        =   39
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmNPCEditor"
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

Dim NPCMode As Integer 'if new npc or editing previous npc. 1 = new, 0 = edit
Dim Width1 As Single
Dim Height1 As Single
Dim Texture As String
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

Public Sub EditNPC(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture1 As String, layer As Integer, Width2 As Single, Height2 As Single, Name As String, Dial1 As String, Dial2 As String, Dial3 As String, Dial4 As String, Dial5 As String, Say1 As String, Say2 As String, Say3 As String, Say4 As String, Say5 As String, Filename As String)
    
    Texture = Texture1
    Width1 = Width2
    Height1 = Height2
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
    
    txtName = Name
    txtDialogue1 = Dial1
    txtDialogue2 = Dial2
    txtDialogue3 = Dial3
    txtDialogue4 = Dial4
    txtDialogue5 = Dial5
    txtSay1 = Say1
    txtSay2 = Say2
    txtSay3 = Say3
    txtSay4 = Say4
    txtSay5 = Say5
    txtNPCFilename = Filename
End Sub

Public Sub StoreTrigger(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture1 As String, layer As Integer, Width2 As Single, Height2 As Single)
    NPCMode = 1
    
    Texture = Texture1
    Width1 = Width2
    Height1 = Height2
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

Private Sub Command1_Click()

End Sub


Private Sub cmbSetNPC_Click()

'if new NPC
If NPCMode = 1 Then
'if npc name and filename blank
If txtName = "" Or txtNPCFilename = "" Then
    MsgBox "NPC must have a name and filename", vbCritical, "NPC Error"
'if npc name filename not blank
Else
    'if file doesnt exit
    If Dir(App.path & "/NPCS/" & txtNPCFilename & ".NPC") = "" Then
        'check if filename stored in memory
        If frmMain.NPCCheck(txtNPCFilename) = False Then
            Call frmMain.Set_NPC(txtName, txtDialogue1, txtDialogue2, txtDialogue3, txtDialogue4, txtDialogue5, txtSay1, txtSay2, txtSay3, txtSay4, txtSay5, txtNPCFilename & ".NPC")
            Call frmMain.PlaceTile(Triggerer.x1, Triggerer.x2, Triggerer.x3, Triggerer.x4, Triggerer.y1, Triggerer.y2, Triggerer.y3, Triggerer.y4, Triggerer.Texture, Triggerer.layer, txtNPCFilename & ".NPC", RGB(ValRed, ValGreen, ValBlue))
            Unload Me
        Else
            MsgBox "NPC filename allready taken. Choose another NPC filename", vbCritical, "NPC Error"
        End If
    
    'if file exist
    Else
        MsgBox "NPC filename allready taken. Choose another NPC filename", vbCritical, "NPC Error"
    End If
End If

'if editing previous NPC
Else


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
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
Set CurrentTexture = Nothing
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

D3DWindow.hDeviceWindow = PicNPC.hWnd

'//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
'//See the lesson text for more information on this line...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, PicNPC.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
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

Private Sub Form_Unload(Cancel As Integer)
   Call frmMain.UnsetVars
   bRunning = False
End Sub


Private Sub HScrollBlue_Change()
lblBlue.Caption = HScrollBlue.Value
ValBlue = HScrollBlue.Value

    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollGreen_Change()
lblGreen.Caption = HScrollGreen.Value
ValGreen = HScrollGreen.Value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollRed_Change()
lblRed.Caption = HScrollRed.Value
ValRed = HScrollRed.Value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

