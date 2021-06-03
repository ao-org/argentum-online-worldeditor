Attribute VB_Name = "modEdicion"
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
' modEdicion
'
' @remarks Funciones de Edicion
' @author gshaxor@gmail.com
' @version 0.1.38
' @date 20061016

Option Explicit

Public maskBloqueo As Byte

Public Sub General_Var_Write(ByVal File As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Writes a var to a text file
    '*****************************************************************
    
    On Error GoTo General_Var_Write_Err
    
    writeprivateprofilestring Main, var, Value, File

    
    Exit Sub

General_Var_Write_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.General_Var_Write", Erl)
    Resume Next
    
End Sub

Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal var As String) As String
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Get a var to from a text file
    '*****************************************************************
    
    On Error GoTo General_Var_Get_Err
    
    Dim L        As Long
    Dim char     As String
    Dim sSpaces  As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), File
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)

    
    Exit Function

General_Var_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.General_Var_Get", Erl)
    Resume Next
    
End Function

''
' Vacia el Deshacer
'
Public Sub Deshacer_Clear()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Clear_Err
    
    Dim i As Integer

    ' Vacio todos los campos afectados
    For i = 1 To maxDeshacer
        MapData_Deshacer_Info(i).Libre = True
    Next
    ' no ahi que deshacer
    FrmMain.mnuDeshacer.Enabled = False

    
    Exit Sub

Deshacer_Clear_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Clear", Erl)
    Resume Next
    
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByVal Desc As String)
    
    On Error GoTo Deshacer_Add_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    If FrmMain.mnuUtilizarDeshacer.Checked = False Then Exit Sub

    Dim i As Integer
    Dim F As Integer
    Dim j As Integer

    ' Desplazo todos los deshacer uno hacia atras
    For i = maxDeshacer To 2 Step -1
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize
                MapData_Deshacer(i, F, j) = MapData_Deshacer(i - 1, F, j)
            Next
        Next
        MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i - 1)
    Next

    ' Guardo los valores
    For F = XMinMapSize To XMaxMapSize
        For j = YMinMapSize To YMaxMapSize
            MapData_Deshacer(1, F, j) = MapData(F, j)
        Next
    Next
    MapData_Deshacer_Info(1).Desc = Desc
    MapData_Deshacer_Info(1).Libre = False
    FrmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
    FrmMain.mnuDeshacer.Enabled = True

    
    Exit Sub

Deshacer_Add_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Add", Erl)
    Resume Next
    
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Recover_Err
    
    Dim i       As Integer
    Dim F       As Integer
    Dim j       As Integer
    Dim Body    As Integer
    Dim Head    As Integer
    Dim Heading As Byte

    If MapData_Deshacer_Info(1).Libre = False Then

        ' Aplico deshacer
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize

                If (MapData(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex <> MapData_Deshacer(1, F, j).NPCIndex) Or (MapData(F, j).NPCIndex <> 0 And MapData_Deshacer(1, F, j).NPCIndex = 0) Then
                    ' Si ahi un NPC, y en el deshacer es otro lo borramos
                    ' (o) Si aun no NPC y en el deshacer no esta
                    MapData(F, j).NPCIndex = 0
                    Call EraseChar(MapData(F, j).CharIndex)

                End If

                If MapData_Deshacer(1, F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex = 0 Then
                    ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                    Body = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Body
                    Head = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Head
                    Heading = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, F, j)
                Else
                    MapData(F, j) = MapData_Deshacer(1, F, j)

                End If

            Next
        Next
        MapData_Deshacer_Info(1).Libre = True

        ' Desplazo todos los deshacer uno hacia adelante
        For i = 1 To maxDeshacer - 1
            For F = XMinMapSize To XMaxMapSize
                For j = YMinMapSize To YMaxMapSize
                    MapData_Deshacer(i, F, j) = MapData_Deshacer(i + 1, F, j)
                Next
            Next
            MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
        Next
        ' borro el ultimo
        MapData_Deshacer_Info(maxDeshacer).Libre = True

        ' ahi para deshacer?
        If MapData_Deshacer_Info(1).Libre = True Then
            FrmMain.mnuDeshacer.Caption = "&Deshacer (no ahi nada que deshacer)"
            FrmMain.mnuDeshacer.Enabled = False
        Else
            FrmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
            FrmMain.mnuDeshacer.Enabled = True

        End If

    Else
        MsgBox "No ahi acciones para deshacer", vbInformation

    End If

    Call DibujarMiniMapa

    
    Exit Sub

Deshacer_Recover_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Recover", Erl)
    Resume Next
    
End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Private Function EditWarning() As Boolean
    
    On Error GoTo EditWarning_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
        EditWarning = True
    Else
        EditWarning = False

    End If

    
    Exit Function

EditWarning_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.EditWarning", Erl)
    Resume Next
    
End Function

''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Bloquear_Bordes_Err
    
    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Bloquear los bordes" ' Hago deshacer

    For y = 1 To 100
        For X = 1 To 100

            If X < 13 Or X > 88 Or y < 10 Or y > 91 Then
                MapData(X, y).Blocked = &HF
            MapInfo.Changed = 1
            End If

        Next X
    Next y

    ' Bloqueo las 4 esquinitas ReyarB
    If MapData(13, 10).TileExit.Map = 0 Then
       MapData(13, 10).Blocked = &HF
       MapInfo.Changed = 1
    End If
    
    If MapData(13, 91).TileExit.Map = 0 Then
        MapData(13, 91).Blocked = &HF
        MapInfo.Changed = 1
    End If
    
    If MapData(88, 10).TileExit.Map = 0 Then
        MapData(88, 10).Blocked = &HF
        MapInfo.Changed = 1
    End If
    
    If MapData(88, 91).TileExit.Map = 0 Then
        MapData(88, 91).Blocked = &HF
        MapInfo.Changed = 1
    End If
    
    
'    MapData(13, 10).TileExit.Map = 0
'    MapData(13, 10).TileExit.X = 0
'    MapData(13, 10).TileExit.y = 0
'
'    MapData(13, 91).TileExit.Map = 0
'    MapData(13, 91).TileExit.X = 0
'    MapData(13, 91).TileExit.y = 0
'
'    MapData(88, 10).TileExit.Map = 0
'    MapData(88, 10).TileExit.X = 0
'    MapData(8, 10).TileExit.y = 0
'
'    MapData(88, 91).TileExit.Map = 0
'    MapData(88, 91).TileExit.X = 0
'    MapData(88, 91).TileExit.y = 0

    Call DibujarMiniMapa

    'Set changed flag
    

    
    Exit Sub

Bloquear_Bordes_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Bloquear_Bordes", Erl)
    Resume Next
    
End Sub

''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim y       As Integer
    Dim X       As Integer
    Dim Cuantos As Integer
    Dim k       As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)

    If Cuantos > 0 Then
        modEdicion.Deshacer_Add "Insertar Superficie al Azar" ' Hago deshacer

        For k = 1 To Cuantos
            X = RandomNumber(10, 90)
            y = RandomNumber(10, 90)

            If frmConfigSup.MOSAICO.Value = vbChecked Then
                Dim aux As Long
                Dim dy  As Integer
                Dim dX  As Integer

                If frmConfigSup.DespMosaic.Value = vbChecked Then
                    dy = Val(frmConfigSup.DMLargo)
                    dX = Val(frmConfigSup.DMAncho.Text)
                Else
                    dy = 0
                    dX = 0

                End If
                
                If FrmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    aux = Val(FrmMain.cGrh.Text) + (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)

                    If FrmMain.cInsertarBloqueo.Value = True Then
                        MapData(X, y).Blocked = &HF
                    Else
                        MapData(X, y).Blocked = 0

                    End If

                    MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                    InitGrh MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)), aux
                Else
                    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                    tXX = X
                    tYY = y
                    desptile = 0

                    For i = 1 To frmConfigSup.mLargo.Text
                        For j = 1 To frmConfigSup.mAncho.Text
                            aux = Val(FrmMain.cGrh.Text) + desptile
                         
                            If FrmMain.cInsertarBloqueo.Value = True Then
                                MapData(tXX, tYY).Blocked = &HF
                            Else
                                MapData(tXX, tYY).Blocked = 0

                            End If

                            MapData(tXX, tYY).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                         
                            InitGrh MapData(tXX, tYY).Graphic(Val(FrmMain.cCapas.Text)), aux
                            tXX = tXX + 1
                            desptile = desptile + 1
                        Next
                        tXX = X
                        tYY = tYY + 1
                    Next
                    tYY = y

                End If

            End If

        Next

    End If

    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Superficie_Bordes_Err
    

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then

                If frmConfigSup.MOSAICO.Value = vbChecked Then
                    Dim aux As Long
                    aux = Val(FrmMain.cGrh.Text) + ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)

                    If FrmMain.cInsertarBloqueo.Value = True Then
                        MapData(X, y).Blocked = &HF
                    Else
                        MapData(X, y).Blocked = 0

                    End If

                    MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                    'Setup GRH
                    InitGrh MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)), aux
                Else

                    'Else Place graphic
                    If FrmMain.cInsertarBloqueo.Value = True Then
                        MapData(X, y).Blocked = &HF
                    Else
                        MapData(X, y).Blocked = 0

                    End If
            
                    MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)).grhindex = Val(FrmMain.cGrh.Text)
            
                    'Setup GRH
    
                    InitGrh MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)), Val(FrmMain.cGrh.Text)

                End If

                'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    EraseChar MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0

                End If

                'Erase Objs
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.grhindex = 0

                'Clear exits
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Superficie_Bordes_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Superficie_Bordes", Erl)
    Resume Next
    
End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Superficie_Todo_Err
    

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If frmConfigSup.MOSAICO.Value = vbChecked Then
                Dim aux As Long
                aux = Val(FrmMain.cGrh.Text) + ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)), aux
            Else
                'Else Place graphic
                MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)).grhindex = Val(FrmMain.cGrh.Text)
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(FrmMain.cCapas.Text)), Val(FrmMain.cGrh.Text)

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Superficie_Todo_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Superficie_Todo", Erl)
    Resume Next
    
End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara

Public Sub Bloqueo_Todo(ByVal Valor As Byte)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Bloqueo_Todo_Err
    

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Bloquear todo el mapa" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, y).Blocked = Valor
        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Bloqueo_Todo_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Bloqueo_Todo", Erl)
    Resume Next
    
End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Borrar_Mapa_Err
    

    If EditWarning Then Exit Sub

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Borrar todo el mapa" ' Hago deshacer
    
    Call engine.Light_Remove_All
    LightA.Delete_All_LigthRound
    
    engine.Particle_Group_Remove_All

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, y).Graphic(1).grhindex = 1
            'Change blockes status
            MapData(X, y).Blocked = 0

            'Erase layer 2 and 3
            MapData(X, y).Graphic(2).grhindex = 0
            MapData(X, y).Graphic(3).grhindex = 0
            MapData(X, y).Graphic(4).grhindex = 0

            'Erase NPCs
            If MapData(X, y).NPCIndex > 0 Then
                EraseChar MapData(X, y).CharIndex
                MapData(X, y).NPCIndex = 0

            End If

            'Erase Objs
            MapData(X, y).OBJInfo.objindex = 0
            MapData(X, y).OBJInfo.Amount = 0
            MapData(X, y).ObjGrh.grhindex = 0

            'Clear exits
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
        
            InitGrh MapData(X, y).Graphic(1), 1
            
            MapData(X, y).luz.Rango = 0
            MapData(X, y).particle_Index = 0
            MapData(X, y).particle_group = 0
            
            MapData(X, y).Trigger = 0

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Borrar_Mapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Borrar_Mapa", Erl)
    Resume Next
    
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_NPCs_Err
    

    modEdicion.Deshacer_Add "Quitar todos los NPCs" & IIf(Hostiles = True, " Hostiles", "") ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).NPCIndex > 0 Then
                Call EraseChar(MapData(X, y).CharIndex)
                MapData(X, y).NPCIndex = 0

            End If
        
        Next X
    Next y

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_NPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_NPCs", Erl)
    Resume Next
    
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Objetos_Err
    

    If EditWarning Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Objetos" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).OBJInfo.objindex > 0 Then
                If MapData(X, y).Graphic(3).grhindex = MapData(X, y).ObjGrh.grhindex Then MapData(X, y).Graphic(3).grhindex = 0
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Objetos_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_Objetos", Erl)
    Resume Next
    
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Triggers_Err
    

    If EditWarning Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Triggers" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).Trigger > 0 Then
                MapData(X, y).Trigger = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Triggers_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_Triggers", Erl)
    Resume Next
    
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo Quitar_Translados_Err
    

    If EditWarning Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Translados" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Translados_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_Translados", Erl)
    Resume Next
    
End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Bordes_Err
    

    If EditWarning Then Exit Sub

    '*****************************************************************
    'Clears a border in a room with current GRH
    '*****************************************************************

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Quitar todos los Bordes" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        
                MapData(X, y).Graphic(1).grhindex = 1
                InitGrh MapData(X, y).Graphic(1), 1
                MapData(X, y).Blocked = 0
            
                'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    EraseChar MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0

                End If

                'Erase Objs
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.grhindex = 0

                'Clear exits
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
            
                ' Triggers
                MapData(X, y).Trigger = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Bordes_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_Bordes", Erl)
    Resume Next
    
End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa

Public Sub Quitar_Capa(ByVal Capa As Byte)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Capa_Err
    

    If EditWarning Then Exit Sub

    '*****************************************************************
    'Clears one layer
    '*****************************************************************

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Quitar Capa " & Capa ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If Capa = 1 Then
                MapData(X, y).Graphic(Capa).grhindex = 1
            Else
                MapData(X, y).Graphic(Capa).grhindex = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Capa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Quitar_Capa", Erl)
    Resume Next
    
End Sub

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tX As Integer, tY As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    ' Selecciones
    
    On Error GoTo DobleClick_Err
    
    Seleccionando = False ' GS
    SeleccionIX = 0
    SeleccionIY = 0
    SeleccionFX = 0
    SeleccionFY = 0
    ' Translados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit

    If tTrans.Map > 0 Then

        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                modMapIO.GuardarMapa FrmMain.Dialog.FileName

            End If

        End If
    
        If LenB(FrmMain.Dialog.FileName) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & ".csm", vbArchive) = True Then
                Call modMapIO.NuevoMapa
                FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.Map & ".csm"
                modMapIO.AbrirMapa FrmMain.Dialog.FileName
                UserPos.X = tTrans.X
                UserPos.y = tTrans.y

                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y - 7
                    CharList(UserCharIndex).Heading = SOUTH

                End If

                FrmMain.mnuReAbrirMapa.Enabled = True

            End If

        End If

    End If

    
    Exit Sub

DobleClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.DobleClick", Erl)
    Resume Next
    
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(Button As Integer, tX As Byte, tY As Byte, Optional ByVal Deshacer As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ClickEdit_Err
    

    Dim loopc    As Integer
    Dim NPCIndex As Integer
    Dim objindex As Integer
    Dim Head     As Integer
    Dim Body     As Integer
    Dim Heading  As Byte
    
    If tY < 1 Or tY > 100 Then Exit Sub
    If tX < 1 Or tX > 100 Then Exit Sub
    
    If Button = 0 Then
        ' Pasando sobre :P
        SobreY = tY
        SobreX = tX
        
        Exit Sub

    End If
    
    'Right
    
    If Button = vbRightButton Then
        ' Posicion
        FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & ENDL & ENDL & "Posición " & tX & "," & tY
        
        ' Bloqueos
        If MapData(tX, tY).Blocked > 0 Then FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & " (BLOQ)"
        
        ' Translados
        If MapData(tX, tY).TileExit.Map <> 0 Then
            If FrmMain.mnuAutoCapturarTranslados.Checked = True Then
                FrmMain.tTMapa.Text = MapData(tX, tY).TileExit.Map
                FrmMain.tTX.Text = MapData(tX, tY).TileExit.X
                FrmMain.tTY = MapData(tX, tY).TileExit.y

            End If

            FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & " (Trans.: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.X & "," & MapData(tX, tY).TileExit.y & ")"

        End If
        
        ' NPCs
        If MapData(tX, tY).NPCIndex > 0 Then
            If MapData(tX, tY).NPCIndex > 499 Then
                FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & " (NPC-Hostil: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).Name & ")"
            Else
                FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & " (NPC: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).Name & ")"

            End If

        End If
        
        ' OBJs
        If MapData(tX, tY).OBJInfo.objindex > 0 Then
            FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & " (Obj: " & MapData(tX, tY).OBJInfo.objindex & " - " & ObjData(MapData(tX, tY).OBJInfo.objindex).Name & " - Cant.:" & MapData(tX, tY).OBJInfo.Amount & ")"

        End If
        
        ' Capas
        FrmMain.StatTxt.Text = FrmMain.StatTxt.Text & ENDL & "Capa1: " & MapData(tX, tY).Graphic(1).grhindex & " - Capa2: " & MapData(tX, tY).Graphic(2).grhindex & " - Capa3: " & MapData(tX, tY).Graphic(3).grhindex & " - Capa4: " & MapData(tX, tY).Graphic(4).grhindex

        If FrmMain.mnuAutoCapturarSuperficie.Checked = True And FrmMain.cSeleccionarSuperficie.Value = False Then
            If MapData(tX, tY).Graphic(4).grhindex <> 0 Then
                FrmMain.cCapas.Text = 4
                FrmMain.cGrh.Text = MapData(tX, tY).Graphic(4).grhindex
            ElseIf MapData(tX, tY).Graphic(3).grhindex <> 0 Then
                FrmMain.cCapas.Text = 3
                FrmMain.cGrh.Text = MapData(tX, tY).Graphic(3).grhindex
            ElseIf MapData(tX, tY).Graphic(2).grhindex <> 0 Then
                FrmMain.cCapas.Text = 2
                FrmMain.cGrh.Text = MapData(tX, tY).Graphic(2).grhindex
            ElseIf MapData(tX, tY).Graphic(1).grhindex <> 0 Then
                FrmMain.cCapas.Text = 1
                FrmMain.cGrh.Text = MapData(tX, tY).Graphic(1).grhindex

            End If
            frmRemplazo.GrhReplaceFrom.Text = FrmMain.cGrh.Text
        End If
        
        ' Limpieza
        If Len(FrmMain.StatTxt.Text) > 4000 Then
            FrmMain.StatTxt.Text = Right(FrmMain.StatTxt.Text, 3000)

        End If

        FrmMain.StatTxt.SelStart = Len(FrmMain.StatTxt.Text)
        
        Exit Sub

    End If
    
    'Left click
    If Button = vbLeftButton Then
            
        'Erase 2-3
        If FrmMain.cQuitarEnTodasLasCapas.Value = True Then
            If Deshacer Then modEdicion.Deshacer_Add "Quitar Todas las Capas (2/3)" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag

            For loopc = 2 To 3
                MapData(tX, tY).Graphic(loopc).grhindex = 0
            Next loopc

            Call DibujarMiniMapa
            Exit Sub

        End If
    
        'Borrar "esta" Capa
        If FrmMain.cQuitarEnEstaCapa.Value = True Then
            If Val(FrmMain.cCapas.Text) = 1 Then
                If MapData(tX, tY).Graphic(1).grhindex <> 1 Then
                    If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa 1" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(1).grhindex = 1
                    Exit Sub

                End If

            ElseIf MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa " & FrmMain.cCapas.Text  ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex = 0
                Call DibujarMiniMapa
                Exit Sub

            End If

        End If
    
        '************** Place grh
        If FrmMain.cSeleccionarSuperficie.Value = True Then
            
            If frmConfigSup.MOSAICO.Value = vbChecked Then
                Dim aux As Long
                Dim dy  As Integer
                Dim dX  As Integer

                If frmConfigSup.DespMosaic.Value = vbChecked Then
                    dy = Val(frmConfigSup.DMLargo)
                    dX = Val(frmConfigSup.DMAncho.Text)
                Else
                    dy = 0
                    dX = 0

                End If
                    
                If FrmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    aux = Val(FrmMain.cGrh.Text) + (((tY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((tX + dX) Mod frmConfigSup.mAncho.Text)

                    If MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex <> aux Or MapData(tX, tY).Blocked <> FrmMain.SelectPanel(2).Value Then
                        MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                        InitGrh MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)), aux

                    End If

                Else
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                    tXX = tX
                    tYY = tY
                    desptile = 0

                    For i = 1 To frmConfigSup.mLargo.Text
                        For j = 1 To frmConfigSup.mAncho.Text
                            aux = Val(FrmMain.cGrh.Text) + desptile

                            If tYY > 100 Then Exit Sub
                            If tXX > 100 Then Exit Sub
                            MapData(tXX, tYY).Graphic(Val(FrmMain.cCapas.Text)).grhindex = aux
                            InitGrh MapData(tXX, tYY).Graphic(Val(FrmMain.cCapas.Text)), aux
                            tXX = tXX + 1
                            desptile = desptile + 1
                        Next
                        tXX = tX
                        tYY = tYY + 1
                    Next
                    tYY = tY
                    
                End If
              
            Else

                'Else Place graphic
                If MapData(tX, tY).Blocked <> FrmMain.SelectPanel(2).Value Or MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex <> Val(FrmMain.cGrh.Text) Then
                    If Deshacer Then modEdicion.Deshacer_Add "Quitar Superficie en Capa " & FrmMain.cCapas.Text ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)).grhindex = Val(FrmMain.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(tX, tY).Graphic(Val(FrmMain.cCapas.Text)), Val(FrmMain.cGrh.Text)
                    
                End If

            End If

            Call DibujarMiniMapa
            
        End If

        '************** Place blocked tile
        If FrmMain.cInsertarBloqueo.Value = True Then
            If MapData(tX, tY).Blocked <> maskBloqueo Then
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = maskBloqueo
                
            End If

        ElseIf FrmMain.cQuitarBloqueo.Value = True Then

            If MapData(tX, tY).Blocked <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = 0

            End If

        End If
    
        '************** Place exit
        If FrmMain.cInsertarTrans.Value = True Then
            If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And FrmMain.cInsertarTransOBJ.Value = True Then
                If ObjData(Cfg_TrOBJ).ObjType = 19 Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto de Translado" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tX, tY).ObjGrh, ObjData(Cfg_TrOBJ).grhindex
                    MapData(tX, tY).OBJInfo.objindex = Cfg_TrOBJ
                    MapData(tX, tY).OBJInfo.Amount = 1

                End If

            End If

            If Val(FrmMain.tTMapa.Text) < -1 Or Val(FrmMain.tTMapa.Text) > 9000 Then
                MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(FrmMain.tTX.Text) < 0 Or Val(FrmMain.tTX.Text) > 100 Then
                MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(FrmMain.tTY.Text) < 0 Or Val(FrmMain.tTY.Text) > 100 Then
                MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                Exit Sub

            End If

            If FrmMain.cUnionManual.Value = True Then
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado de Union Manual' Hago deshacer"
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = Val(FrmMain.tTMapa.Text)

                If tX >= 90 Then ' 21 ' derecha
                    MapData(tX, tY).TileExit.X = 12
                    MapData(tX, tY).TileExit.y = tY
                ElseIf tX <= 11 Then ' 9 ' izquierda
                    MapData(tX, tY).TileExit.X = 91
                    MapData(tX, tY).TileExit.y = tY

                End If

                If tY >= 91 Then ' 94 '''' hacia abajo
                    MapData(tX, tY).TileExit.y = 11
                    MapData(tX, tY).TileExit.X = tX
                ElseIf tY <= 10 Then ''' hacia arriba
                    MapData(tX, tY).TileExit.y = 90
                    MapData(tX, tY).TileExit.X = tX

                End If

            Else
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = Val(FrmMain.tTMapa.Text)
                MapData(tX, tY).TileExit.X = Val(FrmMain.tTX.Text)
                MapData(tX, tY).TileExit.y = Val(FrmMain.tTY.Text)

            End If

        ElseIf FrmMain.cQuitarTrans.Value = True Then
            If Deshacer Then modEdicion.Deshacer_Add "Quitar Translado" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag
            MapData(tX, tY).TileExit.Map = 0
            MapData(tX, tY).TileExit.X = 0
            MapData(tX, tY).TileExit.y = 0

        End If
    
        '************** Place NPC
        If FrmMain.cInsertarFunc(0).Value = True Then
            If FrmMain.cNumFunc(0).Text > 0 Then
                NPCIndex = FrmMain.cNumFunc(0).Text

                If NPCIndex <> MapData(tX, tY).NPCIndex Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                    MapData(tX, tY).NPCIndex = NPCIndex

                End If

            End If

        ElseIf FrmMain.cInsertarFunc(1).Value = True Then

            If FrmMain.cNumFunc(1).Text > 0 Then
                NPCIndex = FrmMain.cNumFunc(1).Text

                If NPCIndex <> (MapData(tX, tY).NPCIndex) Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC Hostil' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                    MapData(tX, tY).NPCIndex = NPCIndex

                End If

            End If

        ElseIf FrmMain.cQuitarFunc(0).Value = True Or FrmMain.cQuitarFunc(1).Value = True Then

            If MapData(tX, tY).NPCIndex > 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar NPC" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).NPCIndex = 0
                Call EraseChar(MapData(tX, tY).CharIndex)
                
                Debug.Print "QUITAR NPC"
                ' Call EraseChar(MapData(X, Y).CharIndex)
                '  MapData(X, Y).NPCIndex = 0
                
            End If

        End If
    
        ' ***************** Control de Funcion de Objetos *****************
        If FrmMain.cInsertarFunc(2).Value = True Then ' Insertar Objeto
            If FrmMain.cNumFunc(2).Text > 0 Then
                objindex = FrmMain.cNumFunc(2).Text

                If MapData(tX, tY).OBJInfo.objindex <> objindex Or MapData(tX, tY).OBJInfo.Amount <> Val(FrmMain.cCantFunc(2).Text) Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tX, tY).ObjGrh, ObjData(objindex).grhindex
                    MapData(tX, tY).OBJInfo.objindex = objindex
                    MapData(tX, tY).OBJInfo.Amount = Val(FrmMain.cCantFunc(2).Text)

                    Select Case ObjData(objindex).ObjType

                        Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                            MapData(tX, tY).Graphic(3) = MapData(tX, tY).ObjGrh
                            MapData(tX, tY).Blocked = &HF

                    End Select

                End If

            End If

        ElseIf FrmMain.cQuitarFunc(2).Value = True Then ' Quitar Objeto

            If MapData(tX, tY).OBJInfo.objindex <> 0 Or MapData(tX, tY).OBJInfo.Amount <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Objeto" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag

                If MapData(tX, tY).Graphic(3).grhindex = MapData(tX, tY).ObjGrh.grhindex Then MapData(tX, tY).Graphic(3).grhindex = 0
                MapData(tX, tY).ObjGrh.grhindex = 0
                MapData(tX, tY).OBJInfo.objindex = 0
                MapData(tX, tY).OBJInfo.Amount = 0
                MapData(tX, tY).Blocked = 0

            End If

        End If
        
        ' ***************** Control de Funcion de Triggers *****************
        If FrmMain.cInsertarTrigger.Value = True Then ' Insertar Trigger
            If TriggerBox < 10 Then
                TriggerBox = FrmMain.lListado(4).ListIndex

            End If

            If MapData(tX, tY).Trigger <> TriggerBox Then
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = TriggerBox

            End If

        ElseIf FrmMain.cQuitarTrigger.Value = True Then ' Quitar Trigger

            If MapData(tX, tY).Trigger <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = 0

            End If
            
        End If
        
        'Ladder
        If FrmMain.insertarParticula.Value = True Then
        Dim particulaindex As Integer
        
        particulaindex = ReadField(2, FrmMain.ListaParticulas.List(FrmMain.ListaParticulas.ListIndex), Asc("#"))
        
            General_Particle_Create CLng(particulaindex), tX, tY
            
            MapData(tX, tY).particle_Index = CLng(particulaindex)

        End If
        
        If FrmMain.quitarparticula.Value = True Then
            MapData(tX, tY).particle_group = 0
            MapData(tX, tY).particle_Index = 0

        End If
        
        If FrmMain.insertarLuz.Value = True Then
            MapData(tX, tY).luz.Rango = FrmMain.RangoLuz
            MapData(tX, tY).luz.color = CLng(FrmMain.ColorLuz)
            
            If MapData(tX, tY).luz.Rango < 100 Then
                engine.Light_Create tX, tY, CLng(FrmMain.ColorLuz), FrmMain.RangoLuz, tX & tY
                engine.Light_Render_All
            Else
                Dim r, g, b As Byte
                b = (CLng(FrmMain.ColorLuz) And 16711680) / 65536
                g = (CLng(FrmMain.ColorLuz) And 65280) / 256
                r = CLng(FrmMain.ColorLuz) And 255
                LightA.Create_Light_To_Map tX, tY, FrmMain.RangoLuz - 99, b, g, r
                LightA.LightRenderAll

            End If
    
        End If
        
        If FrmMain.QuitarLuz.Value = True Then
        
            Dim rangoS As Byte
        
            rangoS = MapData(tX, tY).luz.Rango = 0
        
            MapData(tX, tY).luz.Rango = 0
            MapData(tX, tY).luz.color = 0
             
            Dim id As Integer

            If rangoS < 100 Then
                id = engine.Light_Find(tX & tY)
                engine.Light_Remove id
                engine.Light_Render_All
            Else
                LightA.Delete_Light_To_Map tX, tY

            End If
             
        End If

    End If

    Rem Call DibujarMiniMapa
    
    Exit Sub

ClickEdit_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.ClickEdit", Erl)
    Resume Next
    
End Sub

Public Function Selected_Color()
    
    On Error GoTo Selected_Color_Err
    

    Dim c   As Long
  
    Dim r   As Integer ' Red component value   (0 to 255)
    Dim g   As Integer ' Green component value (0 to 255)
    Dim b   As Integer ' Blue component value  (0 to 255)
  
    Dim Out As String  ' Function output string
    
    ' Setup the color selection palette dialog.
    With FrmMain.CommonDialog1
  
        ' Set initial flags to open the full palette and allow an
        ' initial default color selection.
        .FLAGS = cdlCCFullOpen + cdlCCRGBInit
      
        .color = RGB(255, 255, 255)
      
        ' Display the full color palette
        .ShowColor
        c = .color
                      
    End With

    r = c And 255              ' Get lowest 8 bits  - Red
    g = Int(c / 256) And 255   ' Get middle 8 bits  - Green
    b = Int(c / 65536) And 255 ' Get highest 8 bits - Blue
  
    ' If H mode is selected, replace default with hex RGB values.
    Out = "&H" & Format(Hex(r), "0#") & Format(Hex(g), "0#") & Format(Hex(b), "0#")
    FrmMain.LuzColor.BackColor = RGB(r, g, b)

    Selected_Color = Out

    
    Exit Function

Selected_Color_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Selected_Color", Erl)
    Resume Next
    
End Function

