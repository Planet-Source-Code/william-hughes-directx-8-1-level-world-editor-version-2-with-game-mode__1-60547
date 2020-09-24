Attribute VB_Name = "NarutoWebLevelEditorBas"
'###############################
'
'           Title: NarutoWeb Level Editor Module
'           Desc: Level Editor For NarutoWeb (Online Naruto MMORPG)
'           Written by: William Hughes
'           Started: March 24th 2004
'           Contact: Sim@po2.net
'           Website: www.narutoweb.net
'
'###############################
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Global Active As Boolean
Global Save As Boolean
Global ValRed As Long
Global ValGreen As Long
Global ValBlue As Long
Global TheMapType As Integer
Declare Function GetTickCount Lib "kernel32" () As Long


Type MapType
    name As String
    pass As String
    music As String
    description As String
End Type


Global MapTyper As MapType
Global LeftClick As Integer '0 = no Overwrite tiles 1 = Overwrite tiles
Global LoadXSectors As Integer '# of X sectors Wide to load on loading
Global LoadYSectors As Integer '# of Y Sectors Heigh to load on loading


Public Sub ListLoad(list As ListBox, name As String)
Dim Temp As String
    Open name For Input As #1
    While Not (EOF(1))
        Input #1, Temp
        list.AddItem Temp
    Wend
    Close #1
End Sub


Public Function ReadIni(name As String, key As String, lp As String, filename As String) As String
    Dim Count As Integer
    Dim ret As String * 255
    Count = GetPrivateProfileString(name, key, "", ret, 255, filename)
    ReadIni = Left(ret, Count)
End Function

Public Sub WriteIni(name As String, key As String, lp As String, filename As String)

WritePrivateProfileString name, key, lp, filename

End Sub


Public Sub LoadTilez()
'OLD TILE LOAD SUB
'NO LONGER USED
Dim nodX As Node
Dim Owner As String
Dim ShortName As String
Dim TheKey As String
Dim TheText As String
Dim Count As Integer
Dim ln As String

Count = 0

Open App.path & "/tiles/list.txt" For Input As #1
While Not (EOF(1))
    Input #1, ln$

    If InStr(ln$, ".") Then
        TheKey = Left(ln$, Len(ln$) - 1)
        If InStr(ln$, "\") Then
            Do
            DoEvents
                ln$ = Right(ln$, Len(ln$) - InStr(ln$, "\"))
            Loop Until InStr(ln$, "\") < 1
        End If
        Set nodX = frmMain.Tree.Nodes.Add(Owner, tvwChild, TheKey, ln$)
    Else
        ln$ = Left(ln$, Len(ln$) - 1)
        If InStr(ln$, "\") Then
            TheKey = ln$
            Owner = Left(ln$, InStr(ln$, "\") - 1)
            TheText$ = Right(ln$, Len(ln$) - InStr(ln$, "\"))

            Do
            DoEvents
                TheText$ = Right(TheText$, Len(TheText$) - InStr(TheText$, "\"))
            Loop Until InStr(TheText$, "\") < 1
            Owner = Left(ln$, Len(ln$) - Len(TheText$) - 1)
            Set nodX = frmMain.Tree.Nodes.Add(Owner, tvwChild, TheKey$, TheText$)
            Owner = ln$
        Else
        
            Set nodX = frmMain.Tree.Nodes.Add(, , ln$, ln$)
            Owner = ln$
            
        End If
    End If
Wend

Close #1



End Sub


