VERSION 5.00
Begin VB.Form frmColor 
   Caption         =   "Color Changer"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbSetImage 
      Caption         =   "Set Image"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Color"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CheckBox ChkSave 
      Caption         =   "Save color values"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   2535
   End
   Begin VB.PictureBox PicShow 
      Height          =   4815
      Left            =   3120
      ScaleHeight     =   4755
      ScaleWidth      =   4635
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.HScrollBar HScrollBlue 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   1680
      Value           =   255
      Width           =   2175
   End
   Begin VB.HScrollBar HScrollGreen 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   1080
      Value           =   255
      Width           =   2175
   End
   Begin VB.HScrollBar HScrollRed 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   480
      Value           =   255
      Width           =   2175
   End
   Begin VB.Label lblBlue 
      Caption         =   "255"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblGreen 
      Caption         =   "255"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblRed 
      Caption         =   "255"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmColor"
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


Private Function InitialiseGeometry() As Boolean
    
    
On Error GoTo BailOut: '//Setup our Error handler

'//NOTE THAT WE ARE PASSING VALUES FOR THE tu AND tv ARGUMENTS



'## SECOND SQUARE ##


InitialiseGeometry = True
Exit Function
BailOut:
InitialiseGeometry = False
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

D3DWindow.hDeviceWindow = PicShow.hWnd

'//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
'//See the lesson text for more information on this line...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, PicShow.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
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

Public Sub SetVars(x1 As Single, x2 As Single, x3 As Single, x4 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, Texture1 As String, layer As Integer, Width2 As Single, Height2 As Single)

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
    If layer = 0 Then
        cmbSetImage.Visible = False
    End If
End Sub



Private Sub cmbSetImage_Click()
Call frmMain.PlaceTile(Triggerer.x1, Triggerer.x2, Triggerer.x3, Triggerer.x4, Triggerer.y1, Triggerer.y2, Triggerer.y3, Triggerer.y4, Triggerer.Texture, Triggerer.layer, "", RGB(ValRed, ValGreen, ValBlue))
bRunning = False
Unload Me
End Sub

Private Sub Command1_Click()
bRunning = False
Unload Me
End Sub

Private Sub Command2_Click()

End Sub


Private Sub Form_Activate()
bRunning = Initialise()
HScrollRed.value = ValRed
HScrollBlue.value = ValBlue
HScrollGreen.value = ValGreen

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

Private Sub Form_Unload(Cancel As Integer)
    Call frmMain.UnsetVars
    bRunning = False
End Sub


Private Sub HScrollBlue_Change()
lblBlue.Caption = HScrollBlue.value
ValBlue = HScrollBlue.value

    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollGreen_Change()
lblGreen.Caption = HScrollGreen.value
ValGreen = HScrollGreen.value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

Private Sub HScrollRed_Change()
lblRed.Caption = HScrollRed.value
ValRed = HScrollRed.value
    CurrentTile(0) = CreateTLVertex(0, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 0)
    CurrentTile(1) = CreateTLVertex(Width1, 0, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 0)
    CurrentTile(2) = CreateTLVertex(0, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 0, 1)
    CurrentTile(3) = CreateTLVertex(Width1, Height1, 0, 1, RGB(ValRed, ValGreen, ValBlue), 0, 1, 1)

End Sub

