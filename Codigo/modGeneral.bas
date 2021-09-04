Attribute VB_Name = "modGeneral"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit

Public Type typDevMODE

    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long

End Type

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo CheckKeys_Err
    

    If HotKeysAllow = False Then Exit Sub

    If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.y < 1 Then Exit Sub ' 15
            If LegalPos(UserPos.X, UserPos.y - 1) And WalkMode = True Then
                If dLastWalk + 100 > GetTickCount Then Exit Sub
                UserPos.y = UserPos.y - 1
                MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                dLastWalk = GetTickCount
            ElseIf WalkMode = False Then
                UserPos.y = UserPos.y - 1
            End If
        bRefreshRadar = True ' Radar
        FrmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > XMaxMapSize Then Exit Sub ' 82
            If LegalPos(UserPos.X + 1, UserPos.y) And WalkMode = True Then
                If dLastWalk + 100 > GetTickCount Then Exit Sub
                UserPos.X = UserPos.X + 1
                MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                dLastWalk = GetTickCount
            ElseIf WalkMode = False Then
                UserPos.X = UserPos.X + 1
            End If
        bRefreshRadar = True ' Radar
        FrmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.y > XMaxMapSize Then Exit Sub ' 86
            If LegalPos(UserPos.X, UserPos.y + 1) And WalkMode = True Then
                If dLastWalk + 100 > GetTickCount Then Exit Sub
                UserPos.y = UserPos.y + 1
                MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                dLastWalk = GetTickCount
            ElseIf WalkMode = False Then
                UserPos.y = UserPos.y + 1
            End If
        
        bRefreshRadar = True ' Radar
        FrmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < 1 Then Exit Sub ' 20
            If LegalPos(UserPos.X - 1, UserPos.y) And WalkMode = True Then
                If dLastWalk + 100 > GetTickCount Then Exit Sub
                UserPos.X = UserPos.X - 1
                MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                dLastWalk = GetTickCount
            ElseIf WalkMode = False Then
                UserPos.X = UserPos.X - 1
            End If

        bRefreshRadar = True ' Radar
        FrmMain.SetFocus
        Exit Sub

    End If
    
    
    Exit Sub

CheckKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.CheckKeys", Erl)
    Resume Next
    
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ReadField_Err
    
    Dim i         As Integer
    Dim LastPos   As Integer
    Dim CurChar   As String * 1
    Dim FieldNum  As Integer
    Dim Seperator As String

    Seperator = Chr(SepASCII)
    LastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function

            End If

            LastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = Pos Then
        ReadField = mid(Text, LastPos + 1)

    End If

    
    Exit Function

ReadField_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.ReadField", Erl)
    Resume Next
    
End Function

''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal Path As String) As String
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo autoCompletaPath_Err
    
    Path = Replace(Path, "/", "\")

    If Left(Path, 1) = "\" Then
        ' agrego app.path & path
        Path = App.Path & Path

    End If

    If Right(Path, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        Path = Path & "\"

    End If

    autoCompletaPath = Path

    
    Exit Function

autoCompletaPath_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.autoCompletaPath", Erl)
    Resume Next
    
End Function

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    On Error GoTo Fallo

    Dim tStr As String
    Dim Leer As New clsIniReader

    IniPath = App.Path & "\"

    If FileExist(IniPath & "WorldEditor.ini", vbArchive) = False Then
        Rem  FrmMain.mnuGuardarUltimaConfig.Checked = True
        DirGraficos = IniPath & "Graficos\"
        DirIndex = IniPath & "INIT\"
        DirMidi = IniPath & "MIDI\"
        frmMusica.fleMusicas.Path = DirMidi
        DirDats = IniPath & "..\Recursos\DAT\"
        UserPos.X = 50
        UserPos.y = 50
        PantallaX = 19
        PantallaY = 22
        MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
        Exit Sub

    End If

    Call Leer.Initialize(IniPath & "WorldEditor.ini")

    ' Obj de Translado
    Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
    FrmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
    FrmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
    FrmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))

    ' Guardar Ultima Configuracion
    Rem FrmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))

    'Reciente
    FrmMain.Dialog.InitDir = Leer.GetValue("PATH", "UltimoMapa")

    tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
    UserPos.X = Val(ReadField(1, tStr, Asc("-")))
    UserPos.y = Val(ReadField(2, tStr, Asc("-")))

    If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
        UserPos.X = 50

    End If

    If UserPos.y < YMinMapSize Or UserPos.y > YMaxMapSize Then
        UserPos.y = 50

    End If

    ' Menu Mostrar
    FrmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
    FrmMain.mnuVerCapa2.Checked = Val(Leer.GetValue("MOSTRAR", "Capa2"))
    FrmMain.mnuVerCapa3.Checked = Val(Leer.GetValue("MOSTRAR", "Capa3"))
    FrmMain.mnuVerCapa4.Checked = Val(Leer.GetValue("MOSTRAR", "Capa4"))
    FrmMain.mnuVerTranslados.Checked = Val(Leer.GetValue("MOSTRAR", "Translados"))
    FrmMain.mnuVerObjetos.Checked = Val(Leer.GetValue("MOSTRAR", "Objetos"))
    FrmMain.mnuVerNPCs.Checked = Val(Leer.GetValue("MOSTRAR", "NPCs"))
    FrmMain.mnuVerTriggers.Checked = Val(Leer.GetValue("MOSTRAR", "Triggers"))
    FrmMain.mnuVerMarco.Checked = Val(Leer.GetValue("MOSTRAR", "Marco")) ' Marco
    FrmMain.mnuVerGrilla.Checked = Val(Leer.GetValue("MOSTRAR", "Grilla")) ' Grilla
    VerMarco = FrmMain.mnuVerMarco.Checked
    VerGrilla = FrmMain.mnuVerGrilla.Checked
    FrmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
    FrmMain.cVerTriggers.Value = FrmMain.mnuVerTriggers.Checked
    FrmMain.cVerBloqueos.Value = FrmMain.mnuVerBloqueos.Checked



    ' Tamaño de visualizacion
    PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
    PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))

    If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
    If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32

    ' [GS] 02/10/06
    ' Tamaño de visualizacion en el cliente
    ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
    ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))

    If ClienteHeight <= 0 Then ClienteHeight = 13
    If ClienteWidth <= 0 Then ClienteWidth = 17

        If FrmMain.cVerBloqueos.Value = False Then
            FrmMain.LvBOpcion(0).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(0).BackColor = &H80FF80
        End If
    
        If FrmMain.cVerTriggers.Value = False Then
            FrmMain.LvBOpcion(1).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(1).BackColor = &H80FF80
        End If
        
        If FrmMain.mnuVerCapa1.Checked = False Then
            FrmMain.LvBOpcion(4).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(4).BackColor = &H80FF80
        End If
        
        If FrmMain.mnuVerCapa2.Checked = False Then
            FrmMain.LvBOpcion(5).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(5).BackColor = &H80FF80
        End If
    
        If FrmMain.mnuVerCapa3.Checked = False Then
            FrmMain.LvBOpcion(6).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(6).BackColor = &H80FF80
        End If
        
        If FrmMain.mnuVerCapa4.Checked = False Then
            FrmMain.LvBOpcion(7).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(7).BackColor = &H80FF80
        End If
    
        If FrmMain.mnuVerObjetos.Checked = False Then
            FrmMain.LvBOpcion(2).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(2).BackColor = &H80FF80
        End If
        
        If FrmMain.mnuVerTriggers.Checked = False Then
            FrmMain.LvBOpcion(3).BackColor = &H80000000
        Else
            FrmMain.LvBOpcion(3).BackColor = &H80FF80
        End If
    
    Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical

    Resume Next

End Sub

Public Function TomarBPP() As Integer
    
    On Error GoTo TomarBPP_Err
    
    Dim ModoDeVideo As typDevMODE
    Call EnumDisplaySettings(0, -1, ModoDeVideo)
    TomarBPP = CInt(ModoDeVideo.dmBitsPerPel)

    
    Exit Function

TomarBPP_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.TomarBPP", Erl)
    Resume Next
    
End Function

Public Sub CambioDeVideo()
    '*************************************************
    'Author: Loopzer
    '*************************************************
    
    On Error GoTo CambioDeVideo_Err
    
    Exit Sub
    Dim ModoDeVideo As typDevMODE
    Dim r           As Long
    Call EnumDisplaySettings(0, -1, ModoDeVideo)

    If ModoDeVideo.dmPelsWidth < 1024 Or ModoDeVideo.dmPelsHeight < 768 Then

        Select Case MsgBox("La aplicacion necesita una resolucion minima de 1024 X 768 ,¿Acepta el Cambio de resolucion?", vbInformation + vbOKCancel, "World Editor")

            Case vbOK
                ModoDeVideo.dmPelsWidth = 1024
                ModoDeVideo.dmPelsHeight = 768
                ModoDeVideo.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                r = ChangeDisplaySettings(ModoDeVideo, CDS_TEST)

                If r <> 0 Then
                    MsgBox "Error al cambiar la resolucion, La aplicacion se cerrara."
                    End

                End If

            Case vbCancel
                End

        End Select

    End If

    
    Exit Sub

CambioDeVideo_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.CambioDeVideo", Erl)
    Resume Next
    
End Sub

Public Sub Main()

    '*************************************************
    'Author: Unkwown
    'Last modified: 25/11/08 - GS
    '*************************************************
    On Error Resume Next

    CambioDeVideo
    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
    Dim Chkflag        As Integer
    Windows_Temp_Dir = General_Get_Temp_Dir
    Call CargarMapIni
    Call IniciarCabecera(MiCabecera)
    ColorAmb = 0 'Luz Base por defecto
    FormatoIAO = True

    If FileExist(IniPath & "WorldEditor.jpg", vbArchive) Then frmCargando.picture1.Picture = LoadPicture(IniPath & "WorldEditor.jpg")
    frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
    frmCargando.Show
    frmCargando.SetFocus
    DoEvents
    
    Call InitTileEngine(FrmMain.hWnd, FrmMain.MainViewShp.Top + 47, FrmMain.MainViewShp.Left + 4, 32, 32, PantallaX, PantallaY, 9)
    'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
    
    Call modIndices.CargarMoldes
    
    Call modIndices.CargarIndicesDeGraficos
    frmCargando.X.Caption = "Cargando Indice de Superficies..."
    modIndices.CargarIndicesSuperficie
    DoEvents
    frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
    DoEvents
    Call CargarParticulasBinary
    
    Set LightA = New clsLight
    Call engine.Engine_Init
    Call engine.setup_ambient
    
    frmCargando.P1.Visible = True
    frmCargando.L(0).Visible = True
    frmCargando.X.Caption = "Cargando Cuerpos..."
    modIndices.CargarIndicesDeCuerpos
    DoEvents
    frmCargando.P2.Visible = True
    frmCargando.L(1).Visible = True
    frmCargando.X.Caption = "Cargando Cabezas..."
    modIndices.CargarIndicesDeCabezas
    DoEvents
    frmCargando.P3.Visible = True
    frmCargando.L(2).Visible = True
    frmCargando.X.Caption = "Cargando NPC's..."
    modIndices.CargarIndicesNPC
    DoEvents
    frmCargando.P4.Visible = True
    frmCargando.L(3).Visible = True
    frmCargando.X.Caption = "Cargando Objetos..."
    modIndices.CargarIndicesOBJ
    DoEvents
    frmCargando.P5.Visible = True
    frmCargando.L(4).Visible = True
    frmCargando.X.Caption = "Cargando Triggers..."
    modIndices.CargarIndicesTriggers
    DoEvents
    frmCargando.P6.Visible = True
    frmCargando.L(5).Visible = True
    DoEvents

    'frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando Ventana de Edición..."
    DoEvents

    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) = 0 Then
   
    End If

    MMiniMap_capa2 = True
    MMiniMap_capa1 = True
    engine.Map_Base_Light_Set RGB(255, 255, 255)
    frmCargando.Hide
    FrmMain.Show
    modMapIO.NuevoMapa
    DoEvents

    With MainDestRect
        .Left = (TilePixelWidth * TileBufferSize) - TilePixelWidth
        .Top = (TilePixelHeight * TileBufferSize) - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight

    End With

    With MainViewRect
        .Left = (FrmMain.Left / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (FrmMain.Top / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight

    End With

    prgRun = True
    cFPS = 0
    Chkflag = 0
    dTiempoGT = GetTickCount

    maskBloqueo = &HF

    engine.Start
    
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa FrmMain.Dialog.FileName

        End If

    End If

    LiberarDirectSound
    Dim F

    For Each F In Forms

        Unload F
    Next
    End

End Sub

Public Function GetVar(File As String, Main As String, var As String) As String
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo GetVar_Err
    
    Dim L        As Integer
    Dim char     As String
    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
    szReturn = vbNullString
    sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), File
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)

    
    Exit Function

GetVar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.GetVar", Erl)
    Resume Next
    
End Function

Public Sub WriteVar(File As String, Main As String, var As String, Value As String)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo WriteVar_Err
    
    writeprivateprofilestring Main, var, Value, File

    
    Exit Sub

WriteVar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.WriteVar", Erl)
    Resume Next
    
End Sub

Public Sub ToggleWalkMode()

    '*************************************************
    'Author: Unkwown
    'Last modified: 28/05/06 - GS
    '*************************************************
    On Error GoTo fin:

    If WalkMode = False Then
        WalkMode = True
    Else
        FrmMain.mnuModoCaminata.Checked = False
        WalkMode = False

    End If

    If WalkMode = False Then
        'Erase character
        Call EraseChar(UserCharIndex)
        MapData(UserPos.X, UserPos.y).CharIndex = 0
    Else

        'MakeCharacter
        If LegalPos(UserPos.X, UserPos.y) Then
            Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.y)
            UserCharIndex = MapData(UserPos.X, UserPos.y).CharIndex
            FrmMain.mnuModoCaminata.Checked = True
        Else
            MsgBox "ERROR: Ubicacion ilegal."
            WalkMode = False

        End If

    End If

fin:

End Sub

Public Sub FixCoasts(ByVal grhindex As Long, ByVal X As Integer, ByVal y As Integer)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo FixCoasts_Err
    

    If grhindex = 7284 Or grhindex = 7290 Or grhindex = 7291 Or grhindex = 7297 Or grhindex = 7300 Or grhindex = 7301 Or grhindex = 7302 Or grhindex = 7303 Or grhindex = 7304 Or grhindex = 7306 Or grhindex = 7308 Or grhindex = 7310 Or grhindex = 7311 Or grhindex = 7313 Or grhindex = 7314 Or grhindex = 7315 Or grhindex = 7316 Or grhindex = 7317 Or grhindex = 7319 Or grhindex = 7321 Or grhindex = 7325 Or grhindex = 7326 Or grhindex = 7327 Or grhindex = 7328 Or grhindex = 7332 Or grhindex = 7338 Or grhindex = 7339 Or grhindex = 7345 Or grhindex = 7348 Or grhindex = 7349 Or grhindex = 7350 Or grhindex = 7351 Or grhindex = 7352 Or grhindex = 7349 Or grhindex = 7350 Or grhindex = 7351 Or grhindex = 7354 Or grhindex = 7357 Or grhindex = 7358 Or grhindex = 7360 Or grhindex = 7362 Or grhindex = 7363 Or grhindex = 7365 Or grhindex = 7366 Or grhindex = 7367 Or grhindex = 7368 Or grhindex = 7369 Or grhindex = 7371 Or grhindex = 7373 Or grhindex = 7375 Or grhindex = 7376 Then MapData(X, y).Graphic(2).grhindex = 0

    
    Exit Sub

FixCoasts_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.FixCoasts", Erl)
    Resume Next
    
End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo RandomNumber_Err
    
    Randomize Timer
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    
    Exit Function

RandomNumber_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.RandomNumber", Erl)
    Resume Next
    
End Function

Public Function RandomNumber2(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo RandomNumber2_Err
    
    Randomize Timer
    RandomNumber2 = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    
    Exit Function

RandomNumber2_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.RandomNumber2", Erl)
    Resume Next
    
End Function

''
' Actualiza todos los Chars en el mapa
'

Public Sub RefreshAllChars()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    On Error Resume Next

    Dim loopc As Integer
    FrmMain.ApuntadorRadar.Move UserPos.X - 12, UserPos.y - 10

    bRefreshRadar = False

End Sub

''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
    
    On Error GoTo CaptionWorldEditor_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If Trabajando = vbNullString Then
        Trabajando = "Nuevo Mapa"

    End If

    FrmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"

    If Editado = True Then
        FrmMain.Caption = FrmMain.Caption & " (modificado)"

    End If

    
    Exit Sub

CaptionWorldEditor_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.CaptionWorldEditor", Erl)
    Resume Next
    
End Sub

Private Sub LoadClientSetup()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 26/05/2006
    '26/05/2005 - GS . DirIndex
    '**************************************************************
    
    On Error GoTo LoadClientSetup_Err
    
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open DirIndex & "ao.dat" For Binary Access Read Lock Write As fHandle
    Get fHandle, , ClientSetup
    Close fHandle

    
    Exit Sub

LoadClientSetup_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGeneral.LoadClientSetup", Erl)
    Resume Next
    
End Sub
