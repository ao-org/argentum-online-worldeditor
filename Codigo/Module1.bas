Attribute VB_Name = "Module1"
Option Explicit

'MOTOR DX8 POR LADDER
'MOTOR DX8 POR LADDER
'MOTOR DX8 POR LADDER
'MOTOR DX8 POR LADDER
Public AlphaTecho As Boolean
Public TriggerBox As Byte
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************
    
    On Error GoTo General_Field_Read_Err
    
    Dim i          As Long
    Dim LastPos    As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        General_Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

    
    Exit Function

General_Field_Read_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module1.General_Field_Read", Erl)
    Resume Next
    
End Function

Public Sub DibujarMiniMapa()
   
    If Working Then Exit Sub
    Dim map_x   As Long, map_y As Long
    Dim termine As Boolean

    On Error Resume Next

    FrmMain.MiniMap.BackColor = vbBlack
 
    For map_y = 1 To 100
        For map_x = 1 To 100
        
            If MMiniMap_capa1 Then
                If MapData(map_x, map_y).Graphic(1).grhindex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(1).grhindex).MiniMap_color

                End If

            End If
            
            If MMiniMap_capa2 Then
                If MapData(map_x, map_y).Graphic(2).grhindex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(2).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_capa3 Then
                If MapData(map_x, map_y).Graphic(3).grhindex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(3).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_capa4 Then
                If MapData(map_x, map_y).Graphic(4).grhindex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_Npcs Then
                If MapData(map_x, map_y).NPCIndex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, vbYellow

                End If

            End If
        
            If MMiniMap_objetos Then
                If MapData(map_x, map_y).OBJInfo.objindex > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).ObjGrh.grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_Bloqueos Then
                If MapData(map_x, map_y).Blocked > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, vbRed

                End If

            End If
        
            If MMiniMap_particulas Then
                If MapData(map_x, map_y).particle_Index > 0 Then
                    SetPixel FrmMain.MiniMap.hdc, map_x - 1, map_y - 1, vbWhite

                End If

            End If

            If MMiniMap_Nombre Then
                FrmMain.MiniMap.CurrentX = 30
                FrmMain.MiniMap.CurrentY = 26
                FrmMain.MiniMap.Print FrmMain.MapPest(4).Caption

            End If
   
        Next map_x
    Next map_y
     
    FrmMain.MiniMap.Refresh
    DibujarMiniMapaParaMAPA

End Sub

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long
    
    On Error GoTo General_Field_Count_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Count the number of fields in a delimited string
    '*****************************************************************
    'If string is empty there aren't any fields
    If Len(Text) = 0 Then
        Exit Function

    End If

    Dim i        As Long
    Dim FieldNum As Long
    FieldNum = 0

    For i = 1 To Len(Text)

        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1

        End If

    Next i

    General_Field_Count = FieldNum + 1

    
    Exit Function

General_Field_Count_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module1.General_Field_Count", Erl)
    Resume Next
    
End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long
    
    On Error GoTo General_Particle_Create_Err
    

    If ParticulaInd <= 0 Then Exit Function
    Dim rgb_list(0 To 3) As Long
    rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
    rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
    rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
    rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

    General_Particle_Create = engine.Particle_Group_Create(X, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
       StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
       StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
       StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
       StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
       StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
       StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

    
    Exit Function

General_Particle_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module1.General_Particle_Create", Erl)
    Resume Next
    
End Function

Public Sub CargarParticulasBinary()
    '*************************************
    'Coded by OneZero (onezero_ss@hotmail.com)
    'Last Modified: 6/4/03
    'Loads the Particles.ini file to the ComboBox
    'Edited by Juan Martín Sotuyo Dodero to add speed and life
    '*************************************
    
    On Error GoTo CargarParticulasBinary_Err
    
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim temp       As Integer
    
    Dim handle     As Integer

    'Open files
    handle = FreeFile()

    Dim StreamFile As String

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "particles.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ind!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ind"
    #Else
        StreamFile = App.Path & "\..\Recursos\init\particles.ind"
    #End If

    Dim n As Integer
    
    n = FreeFile()

    Open StreamFile For Binary Access Read As #n
    'num de cabezas
    Get #n, , ParticulasTotales

    ReDim StreamData(1 To ParticulasTotales) As Stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        Get #n, , StreamData(loopc)
        FrmMain.ListaParticulas.AddItem StreamData(loopc).Name & " - #" & loopc
    Next loopc
    
    Close #n

    Exit Sub
    ParticulasTotales = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales

        temp = General_Var_Get(StreamFile, Val(loopc), "resize")
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = General_Var_Get(StreamFile, Val(loopc), "rx")
        StreamData(loopc).grh_resizey = General_Var_Get(StreamFile, Val(loopc), "ry")
        
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, ",")
        Next i

        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
        FrmMain.ListaParticulas.AddItem StreamData(loopc).Name & " - #" & loopc
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

    
    Exit Sub

CargarParticulasBinary_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module1.CargarParticulasBinary", Erl)
    Resume Next
    
End Sub

Public Sub CargarParticulas()
    '*************************************
    'Coded by OneZero (onezero_ss@hotmail.com)
    'Last Modified: 6/4/03
    'Loads the Particles.ini file to the ComboBox
    'Edited by Juan Martín Sotuyo Dodero to add speed and life
    '*************************************
    
    On Error GoTo CargarParticulas_Err
    
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim temp       As Integer
    
    Dim StreamFile As String

    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "particles.ini", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de particles.ini!"
            MsgBox Err.Description

        End If

        StreamFile = Windows_Temp_Dir & "particles.ini"
    #Else
        StreamFile = App.Path & "\..\Recursos\init\particles.ini"
    #End If

    ParticulasTotales = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To ParticulasTotales) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To ParticulasTotales
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).Y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).X2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).Y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        temp = General_Var_Get(StreamFile, Val(loopc), "resize")
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = General_Var_Get(StreamFile, Val(loopc), "rx")
        StreamData(loopc).grh_resizey = General_Var_Get(StreamFile, Val(loopc), "ry")
        
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, ",")
        Next i

        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
        FrmMain.ListaParticulas.AddItem StreamData(loopc).Name & " - #" & loopc
    Next loopc
        
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "particles.ini"
    #End If

    
    Exit Sub

CargarParticulas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module1.CargarParticulas", Erl)
    Resume Next
    
End Sub

Public Sub DibujarMiniMapaParaMAPA()
   
    Dim map_x   As Long, map_y As Long
    Dim termine As Boolean

    On Error Resume Next

    Dim offx As Byte
    Dim offy As Byte

    offx = 13
    FrmMain.MiniMapas2.AutoSize = True

    FrmMain.MiniMapas2.BackColor = vbBlack
 
    For map_y = 7 To 94
        For map_x = 13 To 89
        
            If MMiniMap_capa1 Then
                If MapData(map_x, map_y).Graphic(1).grhindex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, GrhData(MapData(map_x, map_y).Graphic(1).grhindex).MiniMap_color

                End If

            End If
            
            If MMiniMap_capa2 Then
                If MapData(map_x, map_y).Graphic(2).grhindex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, GrhData(MapData(map_x, map_y).Graphic(2).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_capa3 Then
                If MapData(map_x, map_y).Graphic(3).grhindex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, GrhData(MapData(map_x, map_y).Graphic(3).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_capa4 Then
                If MapData(map_x, map_y).Graphic(4).grhindex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, GrhData(MapData(map_x, map_y).Graphic(4).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_Npcs Then
                If MapData(map_x, map_y).NPCIndex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, vbYellow

                End If

            End If
        
            If MMiniMap_objetos Then
                If MapData(map_x, map_y).OBJInfo.objindex > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, GrhData(ObjData(MapData(map_x, map_y).OBJInfo.objindex).grhindex).MiniMap_color

                End If

            End If
        
            If MMiniMap_Bloqueos Then
                If MapData(map_x, map_y).Blocked > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, vbRed

                End If

            End If
        
            If MMiniMap_particulas Then
                If MapData(map_x, map_y).particle_Index > 0 Then
                    SetPixel FrmMain.MiniMapas2.hdc, map_x - offx, map_y - 8, vbWhite

                End If

            End If
    
            If MMiniMap_Nombre Then
                FrmMain.MiniMapas2.CurrentX = 20
                FrmMain.MiniMapas2.CurrentY = 36
                FrmMain.MiniMapas2.Print FrmMain.MapPest(4).Caption

            End If
   
        Next map_x
    Next map_y

    FrmMain.MiniMapas2.AutoSize = True

    FrmMain.MiniMapas2.Refresh
 
End Sub

