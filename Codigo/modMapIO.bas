Attribute VB_Name = "modMapIO"
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
'MOTOR DX8 POR LADDER
''
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit
Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal FileName As String) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo FalloFile

    Dim nFileNum  As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1

End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    
    On Error GoTo FileExist_Err
    

    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True

    End If

    
    Exit Function

FileExist_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.FileExist", Erl)
    Resume Next
    
End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String)
    
    On Error GoTo AbrirMapa_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    If FormatoIAO = True Then
        Call MapaV2_Cargar(Path)
    Else
        Call MapaV3_Cargar(Path)

    End If
    
    
    Exit Sub

AbrirMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.AbrirMapa", Erl)
    Resume Next
    
End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    FrmMain.Dialog.CancelError = True

    On Error GoTo ErrHandler

    If LenB(Path) = 0 Then
        FrmMain.ObtenerNombreArchivo True
        Path = FrmMain.Dialog.FileName

        If LenB(Path) = 0 Then Exit Sub

    End If

    Call MapaV2_Guardar(Path)

ErrHandler:

End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo DeseaGuardarMapa_Err
    

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa Path

        End If

    End If

    
    Exit Sub

DeseaGuardarMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.DeseaGuardarMapa", Erl)
    Resume Next
    
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc As Integer
    Dim y     As Integer
    Dim X     As Integer

    bAutoGuardarMapaCount = 0

    'frmMain.mnuUtirialNuevoFormato.Checked = True
    FrmMain.mnuReAbrirMapa.Enabled = False
    FrmMain.TimAutoGuardarMapa.Enabled = False

    MapaCargado = False

    For loopc = 0 To FrmMain.MapPest.Count - 1
        FrmMain.MapPest(loopc).Enabled = False
    Next

    FrmMain.MousePointer = 11

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            ' Capa 1
            MapData(X, y).Graphic(1).grhindex = 1
        
            ' Bloqueos
            MapData(X, y).Blocked = 0

            ' Capas 2, 3 y 4
            MapData(X, y).Graphic(2).grhindex = 0
            MapData(X, y).Graphic(3).grhindex = 0
            MapData(X, y).Graphic(4).grhindex = 0

            ' NPCs
            If MapData(X, y).CharIndex > 0 Then
                EraseChar MapData(X, y).CharIndex
                MapData(X, y).NPCIndex = 0

            End If

            ' OBJs
            MapData(X, y).OBJInfo.objindex = 0
            MapData(X, y).OBJInfo.Amount = 0
            MapData(X, y).ObjGrh.grhindex = 0

            ' Translados
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
        
            ' Triggers
            MapData(X, y).Trigger = 0
        
            InitGrh MapData(X, y).Graphic(1), 1
        Next X
    Next y

    MapInfo.MapVersion = 0
    MapInfo.Name = "Nuevo Mapa"
    MapInfo.Music = 0
    MapInfo.PK = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.InviSinEfecto = 0
    MapInfo.ResuSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = "No"
    MapInfo.NoEncriptarMP = 0

    Call MapInfo_Actualizar
    Call DibujarMiniMapa

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    FrmMain.MousePointer = 0

    ' Vacio deshacer
    modEdicion.Deshacer_Clear

    MapaCargado = True
    EngineRun = True

    'FrmMain.SetFocus

End Sub

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte

    If FormatoIAO Then
        Save_Map_Data (SaveAs)
        MapInfo.Changed = 0
        Exit Sub

    End If

    Kill SaveAs

    FrmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"

    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    'If FrmMain.lblMapVersion.Caption < 32767 Then
    '    FrmMain.lblMapVersion.Caption = FrmMain.lblMapVersion + 1
    '    frmMapInfo.txtMapVersion = FrmMain.lblMapVersion.Caption
    'End If
    Put FreeFileMap, , CInt(1)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , CByte(Llueve)
    Put FreeFileMap, , CByte(Nieba)
    Put FreeFileMap, , CByte(nieblaV)
    Put FreeFileMap, , CLng(ColorAmb)
    Put FreeFileMap, , CInt(Ambiente)
    Put FreeFileMap, , CInt(AmbienteNoche)
    Put FreeFileMap, , CInt(MidiMusic)
    Put FreeFileMap, , CInt(Mp3Music)
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt 'Cabezera Particulas
    Put FreeFileMap, , TempInt 'Cabezera Luces
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            ByFlags = 0
                
            If MapData(X, y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, y).Graphic(2).grhindex Then ByFlags = ByFlags Or 2
            If MapData(X, y).Graphic(3).grhindex Then ByFlags = ByFlags Or 4
            If MapData(X, y).Graphic(4).grhindex Then ByFlags = ByFlags Or 8
            If MapData(X, y).Trigger Then ByFlags = ByFlags Or 16
            If MapData(X, y).particle_group Then ByFlags = ByFlags Or 32
            If MapData(X, y).luz.Rango Then ByFlags = ByFlags Or 64
                
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, y).Graphic(1).grhindex
                
            For loopc = 2 To 4

                If MapData(X, y).Graphic(loopc).grhindex Then Put FreeFileMap, , MapData(X, y).Graphic(loopc).grhindex
            Next loopc
                
            If MapData(X, y).Trigger Then Put FreeFileMap, , MapData(X, y).Trigger
                    
            If MapData(X, y).particle_group > 0 Then Put FreeFileMap, , MapData(X, y).particle_Index
            ' MsgBox ("Particula: " & MapData(X, y).particle_group & " en: " & X & "-" & y)
                
            If MapData(X, y).luz.Rango > 0 Then
                'MsgBox ("Luz: " & MapData(X, y).Luz.Rango & " en: " & X & "-" & y & "Color: " & MapData(X, y).Luz.Color)
                Put FreeFileMap, , MapData(X, y).luz.Rango
                Put FreeFileMap, , MapData(X, y).luz.color

            End If

            '.inf file
                
            ByFlags = 0
                
            If MapData(X, y).TileExit.Map Then ByFlags = ByFlags Or 1
            If MapData(X, y).NPCIndex Then ByFlags = ByFlags Or 2
            If MapData(X, y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
            Put FreeFileInf, , ByFlags
                
            If MapData(X, y).TileExit.Map Then
                Put FreeFileInf, , MapData(X, y).TileExit.Map
                Put FreeFileInf, , MapData(X, y).TileExit.X
                Put FreeFileInf, , MapData(X, y).TileExit.y

            End If
                
            If MapData(X, y).NPCIndex Then
                
                Put FreeFileInf, , CInt(MapData(X, y).NPCIndex)

            End If
                
            If MapData(X, y).OBJInfo.objindex Then
                Put FreeFileInf, , MapData(X, y).OBJInfo.objindex
                Put FreeFileInf, , MapData(X, y).OBJInfo.Amount

            End If
            
        Next X
    Next y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestañas(SaveAs)

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs

    'Change mouse icon
    FrmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub

''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa
''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV3_Cargar(ByVal Map As String)
    
    On Error GoTo MapaV3_Cargar_Err
    

    Dim loopc       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim y           As Integer
    Dim X           As Integer
    Dim i           As Byte
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    FrmMain.MousePointer = 11

    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left$(Map, Len(Map) - 4)
    Map = Map & ".inf"
    
    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt

    'Vamos a limpiar las luces y particulas del mapa anterior
    'Engine.Particle_Group_Remove_All
    
    'Load arrays Ver ReyarB
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
                For i = 0 To 3

                    If .light_value(i) = True Then
                        .light_value(i) = False

                    End If

                Next i
            
                Get FreeFileMap, , ByFlags
                .Blocked = (ByFlags And 1)
            
                'Layer 1

                Get FreeFileMap, , .Graphic(1).grhindex
                Call InitGrh(.Graphic(1), .Graphic(1).grhindex)
            
                'Layer 2 used?
                If ByFlags And 2 Then

                    Get FreeFileMap, , .Graphic(2).grhindex
                    Call InitGrh(.Graphic(2), .Graphic(2).grhindex)
 
                Else
                
                    .Graphic(2).grhindex = 0
                    
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then

                    Get FreeFileMap, , .Graphic(3).grhindex
                    Call InitGrh(.Graphic(3), .Graphic(3).grhindex)

                Else
                
                    .Graphic(3).grhindex = 0
                    
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    Get FreeFileMap, , .Graphic(4).grhindex
                    Call InitGrh(.Graphic(4), .Graphic(4).grhindex)

                Else
                    
                    .Graphic(4).grhindex = 0

                End If
             
                'Trigger used?
                If ByFlags And 16 Then
                    Get FreeFileMap, , .Trigger
                Else
                    .Trigger = 0

                End If

                'Cargamos el archivo ".INF"
                Get FreeFileInf, , ByFlags
            
                If ByFlags And 1 Then
                    
                    With .TileExit
                    
                        Get FreeFileInf, , .Map
                        Get FreeFileInf, , .X
                        Get FreeFileInf, , .y
                    
                    End With

                End If
    
                If ByFlags And 2 Then
                
                    'Get and make NPC
                    Get FreeFileInf, , .NPCIndex
    
                    If .NPCIndex < 0 Or .NPCIndex > 592 Then
                        .NPCIndex = 0
                    Else
                        Body = NpcData(.NPCIndex).Body
                        Head = NpcData(.NPCIndex).Head
                        Heading = NpcData(.NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)

                    End If

                End If

                If ByFlags And 4 Then
                    
                    'Get and make Object
                    Get FreeFileInf, , .OBJInfo.objindex
                    Get FreeFileInf, , .OBJInfo.Amount

                    If .OBJInfo.objindex > 0 And ObjData(.OBJInfo.objindex).grhindex > 0 Then
                        Call InitGrh(.ObjGrh, ObjData(.OBJInfo.objindex).grhindex)
                    
                    End If

                End If
            
            End With
    
        Next X
    Next y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    Call MapInfo_Cargar(Map)
    
    With FrmMain
    
        '        FrmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        Call modEdicion.Deshacer_Clear
        
        'Change mouse icon
        .MousePointer = 0
                
        Call DibujarMiniMapa

    End With
    
    MapaCargado = True

    
    Exit Sub

MapaV3_Cargar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.MapaV3_Cargar", Erl)
    Resume Next
    
End Sub

Public Sub MapaV2_Cargar(ByVal Map As String)
    '****************************************************************************
    'Author ^[GS]^
    'Last modified 200506
    '
    
    On Error GoTo MapaV2_Cargar_Err
    

    If FormatoIAO Then
        Load_Map_Data_CSM (Map)

        Dim ObtenerMapa As String
    
        ObtenerMapa = FrmMain.MapPest(4).Caption
    
        FrmMain.Label16.Caption = ReadField(3, ObtenerMapa, Asc("a"))
        Exit Sub

    End If

    Dim loopc       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim y           As Integer
    Dim X           As Integer
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    DoEvents

    'Change mouse icon
    FrmMain.MousePointer = 11

    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1

    Map = Left(Map, Len(Map) - 4)
    Map = Map & ".inf"

    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera

    Get FreeFileMap, , Llueve
    Get FreeFileMap, , Nieba
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt

    If Not formatoAo Then
        Get FreeFileMap, , TempInt 'Cabezara Particula
        Get FreeFileMap, , TempInt 'Cabezara Luces

    End If

    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt

    engine.Light_Remove_All
    engine.Particle_Group_Remove_All

    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            Get FreeFileMap, , ByFlags

            MapData(X, y).Blocked = (ByFlags And 1)

            Get FreeFileMap, , MapData(X, y).Graphic(1).grhindex
            InitGrh MapData(X, y).Graphic(1), MapData(X, y).Graphic(1).grhindex

            'Layer 2 used
            If ByFlags And 2 Then
                Get FreeFileMap, , MapData(X, y).Graphic(2).grhindex
                InitGrh MapData(X, y).Graphic(2), MapData(X, y).Graphic(2).grhindex
            Else
                MapData(X, y).Graphic(2).grhindex = 0

            End If

            'Layer 3 used
            If ByFlags And 4 Then
                Get FreeFileMap, , MapData(X, y).Graphic(3).grhindex
                InitGrh MapData(X, y).Graphic(3), MapData(X, y).Graphic(3).grhindex
            Else
                MapData(X, y).Graphic(3).grhindex = 0

            End If

            'Layer 4 used
            If ByFlags And 8 Then
                Get FreeFileMap, , MapData(X, y).Graphic(4).grhindex
                InitGrh MapData(X, y).Graphic(4), MapData(X, y).Graphic(4).grhindex
            Else
                MapData(X, y).Graphic(4).grhindex = 0

            End If

            'Trigger used
            If ByFlags And 16 Then
                Get FreeFileMap, , MapData(X, y).Trigger
            Else
                MapData(X, y).Trigger = 0

            End If

            If Not formatoAo Then
                If ByFlags And 32 Then
                    Get FreeFileMap, , MapData(X, y).particle_Index

                    If MapData(X, y).particle_Index > 0 Then
                        General_Particle_Create MapData(X, y).particle_Index, X, y

                    End If

                Else
                    MapData(X, y).particle_group = 0

                End If

                If ByFlags And 64 Then
                    Get FreeFileMap, , MapData(X, y).luz.Rango
                    Get FreeFileMap, , MapData(X, y).luz.color

                    If MapData(X, y).luz.Rango > 0 Then
                        engine.Light_Create X, y, MapData(X, y).luz.color, MapData(X, y).luz.Rango, X & y

                    End If

                Else
                    MapData(X, y).luz.Rango = 0
                    MapData(X, y).luz.color = 0

                End If

            End If

            '.inf file
            Get FreeFileInf, , ByFlags

            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(X, y).TileExit.Map
                Get FreeFileInf, , MapData(X, y).TileExit.X
                Get FreeFileInf, , MapData(X, y).TileExit.y

            End If

            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(X, y).NPCIndex

                If MapData(X, y).NPCIndex > 0 Then
                    MapData(X, y).NPCIndex = 0
                Else
                    Body = NpcData(MapData(X, y).NPCIndex).Body
                    Head = NpcData(MapData(X, y).NPCIndex).Head
                    Heading = NpcData(MapData(X, y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)

                End If

            End If

            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(X, y).OBJInfo.objindex
                Get FreeFileInf, , MapData(X, y).OBJInfo.Amount

                If MapData(X, y).OBJInfo.objindex > 0 Then
                    InitGrh MapData(X, y).ObjGrh, ObjData(MapData(X, y).OBJInfo.objindex).grhindex

                End If

            End If

        Next X
    Next y

    'Close files
    Close FreeFileMap
    Close FreeFileInf
    Call Pestañas(Map)
    Call DibujarMiniMapa
    engine.Light_Render_All

    bRefreshRadar = True ' Radar

    Map = Left$(Map, Len(Map) - 4) & ".dat"

    MapInfo_Cargar Map
    'FrmMain.lblMapVersion.Caption = MapInfo.MapVersion

    If Nieba Then
        FrmMain.Check2.Value = 1
    Else
        FrmMain.Check2.Value = 0

    End If

    If nieblaV Then
        FrmMain.niebla.Value = 1
    Else
        FrmMain.niebla.Value = 0

    End If

    ' If ColorAmb = -1 Then
    '     FrmMain.Check3.value = 1
    '     FrmMain.Picture3.BackColor = CLng(&HFFFFFF)
    '     engine.Map_Base_Light_Set &HFFFFFF
    ' Else
    '     engine.Map_Base_Light_Set ColorAmb
    '     FrmMain.Picture3.BackColor = CLng(ColorAmb)
    '
    ' FrmMain.Check3.value = 0
    ' End If

    FrmMain.TxtWav.Text = Ambiente
    FrmMain.TxtMidi.Text = MidiMusic
    FrmMain.TxtMp3.Text = Mp3Music

    If Llueve Then
        FrmMain.check1.Value = 1
    Else
        FrmMain.check1.Value = 0

    End If

    MapDat.music_numberHi = Mp3Music
    MapDat.ambient = Ambiente
    MapDat.lluvia = Llueve
    MapDat.nieve = Nieba
    MapDat.niebla = nieblaV
    'Set changed flag
    MapInfo.Changed = 0

    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear

    'Change mouse icon
    FrmMain.MousePointer = 0
    MapaCargado = True
    
    
    Exit Sub

MapaV2_Cargar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.MapaV2_Cargar", Erl)
    Resume Next
    
End Sub

' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo MapInfo_Guardar_Err
    

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save

    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.Name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))
    
    Call WriteVar(Archivo, MapTitulo, "Light", MapInfo.Light)
    
    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")

    End If

    
    Exit Sub

MapInfo_Guardar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.MapInfo_Guardar", Erl)
    Resume Next
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    Dim Leer  As New clsIniReader
    Dim loopc As Integer
    Dim Path  As String
    MapTitulo = Empty
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1

        If mid(Archivo, loopc, 1) = "\" Then
            Path = Left(Archivo, loopc)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.Name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    MapInfo.Light = Leer.GetValue(MapTitulo, "Light")

    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False

    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    'FORMATO IAO
    MapDat.map_name = MapInfo.Name
    MapDat.backup_mode = MapInfo.BackUp
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.music_numberLow = MapInfo.Music
    MapDat.zone = MapInfo.Zona
    MapDat.terrain = MapInfo.Terreno

    MidiMusic = MapDat.music_numberLow
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.Name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
 
    frmMapInfo.chkMapBackup.Value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.Value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapInviSinEfecto.Value = MapInfo.InviSinEfecto
    frmMapInfo.chkMapResuSinEfecto.Value = MapInfo.ResuSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.Value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.Value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion

    If MapInfo.Light <> vbNullString Then
        frmMapInfo.r1.Enabled = True
        frmMapInfo.G1.Enabled = True
        frmMapInfo.b1.Enabled = True
        frmMapInfo.Text1.Enabled = True
        frmMapInfo.lvButtons_H1.Enabled = True
        frmMapInfo.picture1.Enabled = True
        frmMapInfo.check1.Value = 0
        frmMapInfo.Text1 = MapInfo.Light
        frmMapInfo.picture1.BackColor = frmMapInfo.Text1
    Else
        frmMapInfo.picture1.BackColor = &HFFFFFF
        MapInfo.Light = 0
        frmMapInfo.r1.Enabled = False
        frmMapInfo.G1.Enabled = False
        frmMapInfo.b1.Enabled = False
        frmMapInfo.Text1.Enabled = False
        frmMapInfo.Text1 = &HFFFFFF
        frmMapInfo.lvButtons_H1.Enabled = False
        frmMapInfo.picture1.Enabled = False
        frmMapInfo.check1.Value = 1

    End If

    engine.Map_Base_Light_Set frmMapInfo.Text1

End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo Pestañas_Err
    

    Dim loopc As Integer

    If FormatoIAO Then

        For loopc = Len(Map) To 1 Step -1

            If mid(Map, loopc, 1) = "\" Then
                PATH_Save = Left(Map, loopc)
                Exit For

            End If

        Next
        Map = Right(Map, Len(Map) - (Len(PATH_Save)))

        For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

            If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
                NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
                NameMap_Save = Left(Map, loopc)
                Exit For

            End If

        Next

        For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)

            If FileExist(PATH_Save & NameMap_Save & loopc & ".csm", vbArchive) = True Then
                FrmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
                FrmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
                FrmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
            Else
                FrmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

            End If

        Next

    Else

        For loopc = Len(Map) To 1 Step -1

            If mid(Map, loopc, 1) = "\" Then
                PATH_Save = Left(Map, loopc)
                Exit For

            End If

        Next
        Map = Right(Map, Len(Map) - (Len(PATH_Save)))

        For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

            If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
                NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
                NameMap_Save = Left(Map, loopc)
                Exit For

            End If

        Next

        For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)

            If FileExist(PATH_Save & NameMap_Save & loopc & ".map", vbArchive) = True Then
                FrmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
                FrmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
                FrmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
            Else
                FrmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

            End If

        Next

    End If

    
    Exit Sub

Pestañas_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.Pestañas", Erl)
    Resume Next
    
End Sub
