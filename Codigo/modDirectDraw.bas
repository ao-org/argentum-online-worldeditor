Attribute VB_Name = "modDirectDraw"
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
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public BodyData() As BodyData
Public HeadData() As HeadData

Public COLOR_WHITE(3) As Long

''
' modDirectDraw
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)

    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    On Error Resume Next

'    If UserPos.X = 0 Then Exit Sub
'    If UserPos.y = 0 Then Exit Sub
    tX = UserPos.X + viewPortX \ 32 - FrmMain.renderer.ScaleWidth \ 64
    tY = UserPos.y + viewPortY \ 32 - FrmMain.renderer.ScaleHeight \ 64
    tX = tX - 1
    Debug.Print tX; tY

End Sub

Sub ConvertCPtoTPa(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ConvertCPtoTPa_Err
    
    Dim HWindowX As Integer
    Dim HWindowY As Integer

    cx = cx - StartPixelLeft
    cy = cy - StartPixelTop

    HWindowX = (WindowTileWidth \ 2)
    HWindowY = (WindowTileHeight \ 2)

    'Figure out X and Y tiles
    cx = (cx \ TilePixelWidth)
    cy = (cy \ TilePixelHeight)

    If cx > HWindowX Then
        cx = (cx - HWindowX)

    Else

        If cx < HWindowX Then
            cx = (0 - (HWindowX - cx))
        Else
            cx = 0

        End If

    End If

    If cy > HWindowY Then
        cy = (0 - (HWindowY - cy))
    Else

        If cy < HWindowY Then
            cy = (cy - HWindowY)
        Else
            cy = 0

        End If

    End If

    tX = UserPos.X + cx
    tY = UserPos.y + cy

    
    Exit Sub

ConvertCPtoTPa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.ConvertCPtoTPa", Erl)
    Resume Next
    
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal y As Integer)

    On Error Resume Next

    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With CharList(CharIndex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then NumChars = NumChars + 1
        
        '.iHead = Head
        '.iBody = Body
        
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.y = y
        
        'Make active
        .active = 1
      
    End With
    
    'Plot on map
    MapData(X, y).CharIndex = CharIndex
    bRefreshRadar = True ' GS
 
End Sub

Sub EraseChar(CharIndex As Integer)
    
    On Error GoTo EraseChar_Err
    

    CharList(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then

        Do Until CharList(LastChar).active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.y).CharIndex = 0
    
    'Remove char's dialog
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1

    bRefreshRadar = True

    
    Exit Sub

EraseChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.EraseChar", Erl)
    Resume Next
    
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    On Error Resume Next

    If CharIndex = 0 Then Exit Sub

    With CharList(CharIndex)
        .active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False

        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.y = 0

    End With

End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal grhindex As Long, Optional ByVal Started As Byte = 2)

    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    On Error Resume Next

    Grh.grhindex = grhindex

    If Grh.grhindex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.grhindex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0

        End If

    Else

        'Make sure the graphic can be started
        If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
        Grh.Started = Started

    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0

    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.grhindex).speed

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo MoveCharbyHead_Err
    
    Dim addx As Integer
    Dim addy As Integer
    Dim X    As Integer
    Dim y    As Integer
    Dim nX   As Integer
    Dim nY   As Integer

    X = CharList(CharIndex).Pos.X
    y = CharList(CharIndex).Pos.y

    'Figure out which way to move
    Select Case nHeading

        Case NORTH
            addy = -1

        Case EAST
            addx = 1

        Case SOUTH
            addy = 1
    
        Case WEST
            addx = -1
        
    End Select

    nX = X + addx
    nY = y + addy

    MapData(nX, nY).CharIndex = CharIndex
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.y = nY
    MapData(X, y).CharIndex = 0

    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nHeading

    
    Exit Sub

MoveCharbyHead_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.MoveCharbyHead", Erl)
    Resume Next
    
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim X        As Integer
    Dim y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        X = .Pos.X
        y = .Pos.y
        
               
        addx = nX - X
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.y = nY
        MapData(X, y).CharIndex = 0

        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)

    bRefreshRadar = True ' GS
    End With
    
End Sub

Function NextOpenChar() As Integer
    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    
    On Error GoTo NextOpenChar_Err
    
    Dim loopc As Integer

    loopc = 1

    Do While CharList(loopc).active
        loopc = loopc + 1
    Loop

    NextOpenChar = loopc

    
    Exit Function

NextOpenChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.NextOpenChar", Erl)
    Resume Next
    
End Function

Function LegalPos(X As Integer, y As Integer) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 28/05/06 - GS
    '*************************************************
    
    On Error GoTo LegalPos_Err
    
    LegalPos = True
    
    'Check to see if its out of bounds
    If X - 12 < YMinMapSize Or X - 12 > XMaxMapSize Or y - 9 < YMinMapSize Or y - 9 > YMaxMapSize Then
        LegalPos = False
        Exit Function
    End If
    
    'Check to see if its blocked
    If X > XMaxMapSize Or X < XMinMapSize Then Exit Function
    If y > YMaxMapSize Or y < YMinMapSize Then Exit Function
    
    'Check for character
    If MapData(X, y).CharIndex > 0 Then
        LegalPos = False
        Exit Function

    End If
    
    'Tile Bloqueado? (todo bloqueado)
    If MapData(X, y).Blocked > 0 Then
        LegalPos = False
        Exit Function

    End If

    
    Exit Function

LegalPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.LegalPos", Erl)
    Resume Next
    
End Function

Function InMapLegalBounds(X As Integer, y As Integer) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo InMapLegalBounds_Err
    

    If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        InMapLegalBounds = False
        Exit Function

    End If

    InMapLegalBounds = True

    
    Exit Function

InMapLegalBounds_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.InMapLegalBounds", Erl)
    Resume Next
    
End Function

Function InMapBounds(X As Integer, y As Integer) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo InMapBounds_Err
    

    If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        InMapBounds = False
        Exit Function

    End If

    InMapBounds = True

    
    Exit Function

InMapBounds_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.InMapBounds", Erl)
    Resume Next
    
End Function

Public Sub Grh_Render_To_Hdcok(ByRef pic As PictureBox, ByVal grhindex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False)
    
    On Error GoTo Grh_Render_To_Hdcok_Err
    

    If grhindex = 0 Then Exit Sub

    'Public Sub Draw_Grh_Picture(ByVal grh As Long, ByVal pic As PictureBox, _
     ByVal X As Integer, ByVal Y As Integer, _
     ByVal alpha As Boolean, ByVal angle As Single, _
     Optional ByVal ModSizeX2 As Byte = 0, Optional ByVal color As Long = -1)

    Static Piture As RECT

    With Piture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With

    Dim s(3) As Long
    s(0) = -1
    s(1) = -1
    s(2) = -1
    s(3) = -1

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
    engine.Device_Box_Textured_Render grhindex, screen_x, screen_y, GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, s, GrhData(grhindex).sX, GrhData(grhindex).sY, Alpha, 0
                           
    D3DDevice.EndScene
    D3DDevice.Present Piture, ByVal 0, pic.hWnd, ByVal 0

    
    Exit Sub

Grh_Render_To_Hdcok_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.Grh_Render_To_Hdcok", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_To_HdcPNG(ByRef pic As PictureBox, ByVal grhindex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False)
    
    On Error GoTo Grh_Render_To_HdcPNG_Err
    

    If grhindex = 0 Then Exit Sub

    'Public Sub Draw_Grh_Picture(ByVal grh As Long, ByVal pic As PictureBox, _
     ByVal X As Integer, ByVal Y As Integer, _
     ByVal alpha As Boolean, ByVal angle As Single, _
     Optional ByVal ModSizeX2 As Byte = 0, Optional ByVal color As Long = -1)

    Static Piture As RECT

    With Piture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With

    Dim s(3) As Long
    s(0) = -1
    s(1) = -1
    s(2) = -1
    s(3) = -1

    ' D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
    engine.Device_Box_Textured_Render grhindex, screen_x, screen_y, GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, s, GrhData(grhindex).sX, GrhData(grhindex).sY, Alpha, 0
                           
    D3DDevice.EndScene
    D3DDevice.Present Piture, ByVal 0, pic.hWnd, ByVal 0
    
    
    Exit Sub

Grh_Render_To_HdcPNG_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.Grh_Render_To_HdcPNG", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_To_Hdc(ByRef pic As PictureBox, ByVal grhindex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False, Optional ByVal ClearColor As Long = &O0)
    
    On Error GoTo Grh_Render_To_Hdc_Err
    

    If grhindex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With
    
    If GrhData(grhindex).NumFrames > 1 Then
        grhindex = GrhData(grhindex).Frames(1)
    End If

    Call D3DDevice.BeginScene
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, ClearColor, 1#, 0)
    
    engine.Device_Box_Textured_Render grhindex, screen_x, screen_y, GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, COLOR_WHITE, GrhData(grhindex).sX, GrhData(grhindex).sY, Alpha, 0

    Call D3DDevice.EndScene
    Call D3DDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    
    
    Exit Sub

Grh_Render_To_Hdc_Err:
    Call RegistrarError(Err.Number, Err.Description, "Grh_Render_To_Hdc", Erl)
    Resume Next

End Sub

Public Sub Grh_Render_To_HdcSinBorrar(ByRef pic As PictureBox, ByVal grhindex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False)
    
    On Error GoTo Grh_Render_To_HdcSinBorrar_Err
    

    If grhindex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With

    Call D3DDevice.BeginScene
    
    engine.Device_Box_Textured_Render grhindex, screen_x, screen_y, GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, COLOR_WHITE, GrhData(grhindex).sX, GrhData(grhindex).sY, Alpha, 0

    Call D3DDevice.EndScene
    Call D3DDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    
    
    Exit Sub

Grh_Render_To_HdcSinBorrar_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Render_To_HdcSinBorrar", Erl)
    Resume Next
    
End Sub

' [Loopzer]
Public Sub DePegar()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo DePegar_Err
    
    Dim X As Integer
    Dim y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            MapData(X + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(X, y)
        Next
    Next
 
    
    Exit Sub

DePegar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.DePegar", Erl)
    Resume Next
    
End Sub

Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo PegarSeleccion_Err
    
    Static UltimoX As Integer
    Static UltimoY As Integer
    'If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    Debug.Print SobreX
    Debug.Print SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
  
            DeSeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            DeSeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            DeSeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)

            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
            SeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            SeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            SeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
        
            MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC

        Next
    Next
    Seleccionando = False
    Call DibujarMiniMapa

    
    Exit Sub

PegarSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.PegarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub PegarSeleccionCasa() '(mx As Integer, my As Integer)
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo PegarSeleccionCasa_Err
    
    Static UltimoX As Integer
    Static UltimoY As Integer
    'If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    Debug.Print SobreX
    Debug.Print SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
  
            DeSeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            DeSeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            DeSeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)

            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC

        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
            SeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            SeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            SeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
        
            MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC

        Next
    Next
    Seleccionando = False
    Call DibujarMiniMapa

    
    Exit Sub

PegarSeleccionCasa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.PegarSeleccionCasa", Erl)
    Resume Next
    
End Sub

Public Sub AccionSeleccion()

    Working = True

    '*************************************************
    'Author: Loopzera
    'Last modified: 21/11/07
    '*************************************************
    On Error Resume Next

    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + y, False
        Next
    Next
    Seleccionando = False
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Tarea finalizada.", 255, 0, 0, False, True, False)
    Working = False
    Call DibujarMiniMapa

End Sub

Public Sub BlockearSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo BlockearSeleccion_Err
    
    Dim X     As Integer
    Dim y     As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1

            If MapData(X + SeleccionIX, y + SeleccionIY).Blocked > 0 Then
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 0
            Else
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = &HF

            End If

        Next
    Next
    Seleccionando = False

    
    Exit Sub

BlockearSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.BlockearSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub CortarSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo CortarSeleccion_Err
    
    CopiarSeleccion
    Dim X     As Integer
    Dim y     As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            MapData(X + SeleccionIX, y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
    Call DibujarMiniMapa

    
    Exit Sub

CortarSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.CortarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub CopiarSeleccionCasa()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo CopiarSeleccionCasa_Err
    
    SeleccionIX = 65
    SeleccionFX = 74
    
    SeleccionIY = 23
    SeleccionFY = 30
    
    Dim X As Integer
    Dim y As Integer
    Debug.Print
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    
    Debug.Print SeleccionIX
    Debug.Print SeleccionFX
    
    Debug.Print SeleccionIY
    Debug.Print SeleccionFY
    
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    MapInfo.Changed = 1

    
    Exit Sub

CopiarSeleccionCasa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.CopiarSeleccionCasa", Erl)
    Resume Next
    
End Sub

Public Sub CopiarSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo CopiarSeleccion_Err
    
    Dim X As Integer
    Dim y As Integer
    Debug.Print
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    
    Debug.Print SeleccionIX
    Debug.Print SeleccionFX
    
    Debug.Print SeleccionIY
    Debug.Print SeleccionFY
    
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    MapInfo.Changed = 1

    
    Exit Sub

CopiarSeleccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.CopiarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub GenerarVista()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    ' hacer una llamada a un seter o geter , es mas lento q una variable
    ' con esto hacemos q no este preguntando a el objeto cadavez
    ' q dibuja , Render mas rapido ;)
    
    On Error GoTo GenerarVista_Err
    
    VerBlockeados = FrmMain.cVerBloqueos.Value
    VerTriggers = FrmMain.cVerTriggers.Value
    VerCapa1 = FrmMain.mnuVerCapa1.Checked
    VerCapa2 = FrmMain.mnuVerCapa2.Checked
    VerCapa3 = FrmMain.mnuVerCapa3.Checked
    VerCapa4 = FrmMain.mnuVerCapa4.Checked
    VerTranslados = FrmMain.mnuVerTranslados.Checked
    VerObjetos = FrmMain.mnuVerObjetos.Checked
    VerNpcs = FrmMain.mnuVerNPCs.Checked
    VerParticulas = FrmMain.mnuVerParticulas.Checked
    VerLuces = FrmMain.mnuVerParticulas.Checked
    
    
    Exit Sub

GenerarVista_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.GenerarVista", Erl)
    Resume Next
    
End Sub

Function HayUserAbajo(X As Integer, y As Integer, grhindex) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo HayUserAbajo_Err
    
    HayUserAbajo = CharList(UserCharIndex).Pos.X >= X - (GrhData(grhindex).TileWidth \ 2) And CharList(UserCharIndex).Pos.X <= X + (GrhData(grhindex).TileWidth \ 2) And CharList(UserCharIndex).Pos.y >= y - (GrhData(grhindex).TileHeight - 1) And CharList(UserCharIndex).Pos.y <= y

    
    Exit Function

HayUserAbajo_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.HayUserAbajo", Erl)
    Resume Next
    
End Function

Function PixelPos(X As Integer) As Integer
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo PixelPos_Err
    

    PixelPos = (TilePixelWidth * X) - TilePixelWidth

    
    Exit Function

PixelPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.PixelPos", Erl)
    Resume Next
    
End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 15/10/06 by GS
    '*************************************************
    
    On Error GoTo InitTileEngine_Err
    

    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize

    '[GS] 02/10/2006
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)

    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    '****** INIT DirectDraw ******
    ' Create the root DirectDraw object

    frmCargando.X.Caption = "Iniciando Control de Superficies..."
    '    DoEvents

    InitTileEngine = True

    
    Exit Function

InitTileEngine_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDirectDraw.InitTileEngine", Erl)
    Resume Next
    
End Function
