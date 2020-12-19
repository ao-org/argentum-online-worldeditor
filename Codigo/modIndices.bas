Attribute VB_Name = "modIndices"
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
' modIndices
'
' @remarks Funciones Especificas al Trabajo con Indices
' @author gshaxor@gmail.com
' @version 0.1.05
' @date 20060530

Option Explicit

Private Type tMoldeCuerpo
    X As Long
    y As Long
    Width As Long
    Height As Long
    DirCount(1 To 4) As Byte
    TotalGrhs As Long
End Type

Private MoldesBodies() As tMoldeCuerpo
Private BodiesHeading(1 To 4) As E_Heading

''
' Carga los indices de Graficos
'

Public Sub CargarIndicesDeGraficos()

    On Error GoTo ErrorHandler

    Dim Grh         As Long
    Dim Frame       As Long
    Dim grhCount    As Long
    Dim handle      As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "graficos.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de recurso!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "graficos.ind" For Binary Access Read As #handle
    #Else
        Open App.Path & "\..\Recursos\init\graficos.ind" For Binary Access Read As #handle
    #End If
    
    'Get file version
    Get #handle, , fileVersion
    
    'Get number of grhs
    Get #handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    MaxGrhs = grhCount
    Dim fin As Boolean
    fin = False

    While Not EOF(handle) And fin = False

        Get #handle, , Grh

        With GrhData(Grh)
        
            GrhData(Grh).active = True
            'Get number of frames
            Get #handle, , .NumFrames

            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get #handle, , .Frames(Frame)

                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler

                    End If

                Next Frame
                
                Get #handle, , GrhData(Grh).speed
                
                If .speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelWidth = GrhData(.Frames(1)).pixelWidth

                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .pixelHeight = GrhData(.Frames(1)).pixelHeight

                If .pixelHeight <= 0 Then GoTo ErrorHandler
                                                
                .TileWidth = GrhData(.Frames(1)).TileWidth

                If .TileWidth <= 0 Then GoTo ErrorHandler

                .TileHeight = GrhData(.Frames(1)).TileHeight

                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get #handle, , .FileNum

                If .FileNum <= 0 Then GoTo ErrorHandler
                                
                Get #handle, , GrhData(Grh).sX

                If .sX < 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(Grh).sY

                If .sY < 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(Grh).pixelWidth

                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get #handle, , GrhData(Grh).pixelHeight

                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth

                .Frames(1) = Grh

            End If

        End With

        If Grh = MaxGrhs Then fin = True
    Wend

    Close #handle
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "graficos.ind"
    #End If
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT", "minimap.bin", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de recursos (minimap.bin)!"
            GoTo ErrorHandler

        End If
    
        Open Windows_Temp_Dir & "minimap.bin" For Binary Access Read As #handle
    #Else
        Open App.Path & "\..\Recursos\init\minimap.bin" For Binary Access Read As #handle
    #End If

    Dim Count As Long

    For Count = 1 To MaxGrhs

        If GrhData(Count).active Then
            Get #handle, , GrhData(Count).MiniMap_color

        End If

    Next Count

    Close #handle
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "minimap.bin"
    #End If

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Description & " durante la carga de Grh.dat! La carga se ha detenido en GRH: " & Grh

End Sub

Sub CargarMoldes()

    BodiesHeading(1) = E_Heading.south
    BodiesHeading(2) = E_Heading.NORTH
    BodiesHeading(3) = E_Heading.WEST
    BodiesHeading(4) = E_Heading.EAST
    
    Dim Loader As clsIniManager
    Set Loader = New clsIniManager
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "moldes.ini", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de moldes.ini!"
            MsgBox Err.Description

        End If

        Call Loader.Initialize(Windows_Temp_Dir & "moldes.ini")
    #Else
        Call Loader.Initialize(App.Path & "\..\Recursos\init\moldes.ini")
    #End If
    
    Dim NumMoldes As Integer
    NumMoldes = Val(Loader.GetValue("INIT", "Moldes"))

    ReDim MoldesBodies(1 To NumMoldes)
    
    Dim i As Integer, MoldeKey As String
    
    For i = 1 To NumMoldes
        MoldeKey = "Molde" & i
    
        With MoldesBodies(i)
            .X = Val(Loader.GetValue(MoldeKey, "X"))
            .y = Val(Loader.GetValue(MoldeKey, "Y"))
            .Width = Val(Loader.GetValue(MoldeKey, "Width"))
            .Height = Val(Loader.GetValue(MoldeKey, "Height"))
            .DirCount(1) = Val(Loader.GetValue(MoldeKey, "Dir1"))
            .DirCount(2) = Val(Loader.GetValue(MoldeKey, "Dir2"))
            .DirCount(3) = Val(Loader.GetValue(MoldeKey, "Dir3"))
            .DirCount(4) = Val(Loader.GetValue(MoldeKey, "Dir4"))
            .TotalGrhs = .DirCount(1) + .DirCount(2) + .DirCount(3) + .DirCount(4) + 4
        End With
    Next
    
    Set Loader = Nothing
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "moldes.ini"
    #End If

End Sub

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************

    On Error GoTo ErrorHandler
   
    Dim FileDir As String
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT", "indices.ini", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de superficies (indices.ini)!"
            GoTo ErrorHandler

        End If
    
        FileDir = Windows_Temp_Dir & "indices.ini"
    #Else
        FileDir = App.Path & "\..\Recursos\init\indices.ini"
    #End If
    
    Dim Leer As New clsIniReader
    Dim i    As Long
    Leer.Initialize FileDir
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    FrmMain.lListado(0).Clear

    For i = 0 To MaxSup
        SupData(i).Name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        FrmMain.lListado(0).AddItem i & "- " & SupData(i).Name
    Next
    
    #If Compresion = 1 Then
        Delete_File FileDir
    #End If

    DoEvents
    Exit Sub
ErrorHandler:
    MsgBox "Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If FileExist(App.Path & "\..\Recursos\dat\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End

    End If

    Dim Obj  As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(App.Path & "\..\Recursos\dat\OBJ.dat")
    FrmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData

    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).grhindex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        FrmMain.lListado(3).AddItem Obj & "- " & ObjData(Obj).Name
    Next Obj

    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************

    On Error GoTo Fallo

    Dim FileDir As String
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT", "triggers.ini", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de triggers (triggers.ini)!"
            GoTo Fallo

        End If
    
        FileDir = Windows_Temp_Dir & "triggers.ini"
    #Else
        FileDir = App.Path & "\..\Recursos\init\triggers.ini"
    #End If
    
    Dim NumT As Integer
    Dim T    As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(FileDir)
    FrmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))

    For T = 1 To NumT
        FrmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1)
    Next T
    
    #If Compresion = 1 Then
        Delete_File FileDir
    #End If

    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
    
    On Error GoTo CargarIndicesDeCuerpos_Err
    
    Dim Loader       As clsIniManager

    Dim i            As Long
    
    Dim j            As Byte
    
    Dim k            As Integer
    
    Dim Heading      As Byte
    
    Dim BodyKey      As String
    
    Dim Std          As Byte

    Dim NumCuerpos   As Integer
    
    Dim LastGrh      As Long
    
    Dim AnimStart    As Long
    
    Dim X            As Long
    
    Dim y            As Long
    
    Dim FileNum      As Long
    
    Set Loader = New clsIniManager
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "cuerpos.dat", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de cuerpos.dat!"
            MsgBox Err.Description

        End If

        Call Loader.Initialize(Windows_Temp_Dir & "cuerpos.dat")
    #Else
        Call Loader.Initialize(App.Path & "\..\Recursos\init\cuerpos.dat")
    #End If
    
    NumCuerpos = Val(Loader.GetValue("INIT", "NumBodies"))
    
    'Resize array
    ReDim Preserve BodyData(0 To NumCuerpos)

    For i = 1 To NumCuerpos
        BodyKey = "BODY" & i
    
        Std = Val(Loader.GetValue(BodyKey, "Std"))
        BodyData(i).HeadOffset.X = Val(Loader.GetValue(BodyKey, "HeadOffsetX"))
        BodyData(i).HeadOffset.y = Val(Loader.GetValue(BodyKey, "HeadOffsetY"))

        If Std = 0 Then
            InitGrh BodyData(i).Walk(1), Val(Loader.GetValue(BodyKey, "Walk1")), 0
            InitGrh BodyData(i).Walk(2), Val(Loader.GetValue(BodyKey, "Walk2")), 0
            InitGrh BodyData(i).Walk(3), Val(Loader.GetValue(BodyKey, "Walk3")), 0
            InitGrh BodyData(i).Walk(4), Val(Loader.GetValue(BodyKey, "Walk4")), 0
            
        Else
            FileNum = Val(Loader.GetValue(BodyKey, "FileNum"))
        
            LastGrh = UBound(GrhData)

            ' Agrego espacio para meter el body en GrhData
            ReDim Preserve GrhData(1 To LastGrh + MoldesBodies(Std).TotalGrhs)
            
            MaxGrhs = UBound(GrhData)
            
            LastGrh = LastGrh + 1
            X = MoldesBodies(Std).X
            y = MoldesBodies(Std).y
            
            For j = 1 To 4
                AnimStart = LastGrh
            
                For k = 1 To MoldesBodies(Std).DirCount(j)
                    With GrhData(LastGrh)
                        .FileNum = FileNum
                        .NumFrames = 1
                        .sX = X
                        .sY = y
                        .pixelWidth = MoldesBodies(Std).Width
                        .pixelHeight = MoldesBodies(Std).Height
                        
                        .TileWidth = .pixelWidth / TilePixelHeight
                        .TileHeight = .pixelHeight / TilePixelWidth
        
                        ReDim .Frames(1)
                        .Frames(1) = LastGrh
                    End With
                    
                    LastGrh = LastGrh + 1
                    X = X + MoldesBodies(Std).Width
                Next
                
                X = MoldesBodies(Std).X
                y = y + MoldesBodies(Std).Height
                
                Heading = BodiesHeading(j)
                
                With GrhData(LastGrh)
                    .NumFrames = MoldesBodies(Std).DirCount(j)
                    .speed = .NumFrames / 0.018
                    
                    ReDim .Frames(1 To MoldesBodies(Std).DirCount(j))
                    
                    For k = 1 To MoldesBodies(Std).DirCount(j)
                        .Frames(k) = AnimStart + k - 1
                    Next
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                End With
                
                InitGrh BodyData(i).Walk(Heading), LastGrh, 0
                
                LastGrh = LastGrh + 1
            Next

        End If

    Next i

    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cuerpos.dat"
    #End If

    Set Loader = Nothing

    
    Exit Sub

CargarIndicesDeCuerpos_Err:
    Call RegistrarError(Err.Number, Err.Description, "modIndices.CargarIndicesDeCuerpos", Erl)
    Resume Next
    
End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
    
    On Error GoTo CargarIndicesDeCabezas_Err
    
    Dim n            As Integer
    Dim i            As Long
    Dim Numheads     As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "cabezas.ind", Windows_Temp_Dir, False) Then
            Err.Description = "¡No se puede cargar el archivo de Cabezas.ind!"
            MsgBox Err.Description

        End If

        Open Windows_Temp_Dir & "cabezas.ind" For Binary Access Read As #n
    #Else
        Open App.Path & "\..\Recursos\init\cabezas.ind" For Binary Access Read As #n
    #End If

    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #n
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "cabezas.ind"
    #End If
    
    
    Exit Sub

CargarIndicesDeCabezas_Err:
    Call RegistrarError(Err.Number, Err.Description, "modIndices.CargarIndicesDeCabezas", Erl)
    Resume Next
    
End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    On Error Resume Next

    'On Error GoTo Fallo
    If FileExist(App.Path & "\..\Recursos\dat\" & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & App.Path & "\..\Recursos\dat\", vbCritical
        End

    End If

    'If FileExist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
    '    End
    'End If
    Dim Trabajando As String
    Dim NPC        As Integer
    Dim Leer       As New clsIniReader
    FrmMain.lListado(1).Clear
    FrmMain.lListado(2).Clear
    Call Leer.Initialize(App.Path & "\..\Recursos\dat\" & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    ReDim NpcData(1000) As NpcData
    Trabajando = "Dats\NPCs.dat"

    For NPC = 1 To NumNPCs
        NpcData(NPC).Name = Leer.GetValue("NPC" & NPC, "Name")
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))

        If LenB(NpcData(NPC).Name) <> 0 Then FrmMain.lListado(1).AddItem NPC & "- " & NpcData(NPC).Name
    Next

    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
