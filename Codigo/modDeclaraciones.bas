Attribute VB_Name = "modDeclaraciones"
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
' modDeclaraciones
'
' @remarks Declaraciones
' @author ^[GS]^
' @version 0.1.12
' @date 20081218

Option Explicit
'Compresion

Public LightA           As New clsLight

Public Windows_Temp_Dir As String
Public NombreMapa       As String
Public UltimoClickX     As Byte
Public UltimoClickY     As Byte

Public map_base_light   As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long)

Public formatoAo             As Boolean

Public Llueve                As Byte
Public Nieba                 As Byte
Public nieblaV               As Byte

Public Mp3Music              As Integer
Public MidiMusic             As Integer
Public Ambiente              As String
Public AmbienteNoche         As Integer
Public ColorAmb              As Long

Public Const MSGMod          As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"
Public Const MSGDang         As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que desea continuar?"

Public Const ENDL            As String * 2 = vbCrLf
'[Loopzer]
Public SeleccionIX           As Integer
Public SeleccionFX           As Integer
Public SeleccionIY           As Integer
Public SeleccionFY           As Integer
Public SeleccionAncho        As Integer
Public SeleccionAlto         As Integer
Public Seleccionando         As Boolean
Public SeleccionMap()        As MapBlock

Public DeSeleccionOX         As Integer
Public DeSeleccionOY         As Integer
Public DeSeleccionIX         As Integer
Public DeSeleccionFX         As Integer
Public DeSeleccionIY         As Integer
Public DeSeleccionFY         As Integer
Public DeSeleccionAncho      As Integer
Public DeSeleccionAlto       As Integer
Public DeSeleccionando       As Boolean
Public DeSeleccionMap()      As MapBlock

Public VerBlockeados         As Boolean
Public VerTriggers           As Boolean
Public VerMarco              As Boolean ' Marco
Public VerGrilla             As Boolean ' grilla
Public VerCapa1              As Boolean
Public VerCapa2              As Boolean
Public VerCapa3              As Boolean
Public VerCapa4              As Boolean
Public VerTranslados         As Boolean
Public VerObjetos            As Boolean
Public VerNpcs               As Boolean
Public VerParticulas         As Boolean
Public VerLuces              As Boolean
'[/Loopzer]

' Objeto de Translado
Public Cfg_TrOBJ             As Integer

'Path
Public IniPath               As String
Public DirGraficos           As String
Public DirMidi               As String
Public DirIndex              As String
Public DirDats               As String

Public bAutoGuardarMapa      As Byte
Public bAutoGuardarMapaCount As Byte
Public HotKeysAllow          As Boolean  ' Control Automatico de HotKeys
Public vMostrando            As Byte
Public WORK                  As Boolean
Public PATH_Save             As String
Public NumMap_Save           As Integer
Public NameMap_Save          As String

' DX Config
Public PantallaX             As Integer
Public PantallaY             As Integer

' [GS] 02/10/06
' Client Config
Public ClienteHeight         As Integer
Public ClienteWidth          As Integer

Public Type tSetupMods

    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean

End Type

Public ClientSetup   As tSetupMods

Public SobreX        As Byte   ' Posicion X bajo el Cursor
Public SobreY        As Byte   ' Posicion Y bajo el Cursor

' Radar
Public MiRadarX      As Integer
Public MiRadarY      As Integer
Public bRefreshRadar As Boolean

Public Type GrhData

    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    active As Boolean
    MiniMap_color As Long

End Type

Public Type Position

    X As Integer
    y As Integer

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

    grhindex As Long
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single

End Type

Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

'Lista de cuerpos

Public GrhData()        As GrhData 'Guarda todos los grh

Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData()  As HeadData

Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

Type SupData

    Name As String
    Grh As Long
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte

End Type

Public MaxSup    As Integer
Public SupData() As SupData

Public Type NpcData

    Name As String
    Body As Integer
    Head As Integer
    Heading As Byte

End Type

Public NumNPCs   As Long
'Public NumNPCsHOST As Integer
Public NpcData() As NpcData

Public Type ObjData

    Name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    grhindex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
    Cerrada As Byte
    Subtipo As Byte

End Type

Public NumOBJs     As Long
Public ObjData()   As ObjData

Public Conexion    As New Connection
Public prgRun      As Boolean
Public CurrentGrh  As Grh
Public Play        As Boolean
Public MapaCargado As Boolean
Public cFPS        As Long
Public dTiempoGT   As Double
Public dLastWalk   As Double

'Hold info about each map
Public Type MapInfo

    Music As String
    Name As String
    MapVersion As Integer
    PK As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    NoEncriptarMP As Byte
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    Changed As Byte ' flag for WorldEditor
    Light As String

End Type

'********** CONSTANTS ***********
'Heading Constants

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'********** TYPES ***********
'Holds a local position

'Holds a world position
Public Type WorldPos

    Map As Integer
    X As Integer
    y As Integer

End Type

'Points to a grhData and keeps animation info

'Holds data about where a bmp can be found,
'How big it is and animation info

Rem Particle Groups
Public ParticulasTotales As Integer
Public StreamData()      As Stream

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

'Hold info about a character

Public Type char

    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    group_index As Integer
    
    Criminal As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte

End Type

'Holds info about a object
Public Type Obj

    objindex As Integer
    Amount As Integer

End Type

Public Type Light

    Rango As Integer
    color As Long

End Type

'Holds info about each tile position
Public Type MapBlock

    Graphic(1 To 8) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    light_value(3) As Long
    
    luz As Light
    color(3) As Long
    particle_group As Integer
    particle_Index As Integer

    Marcado As Boolean
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    cRojo As Byte
    cVerde As Byte
    cAzul As Byte
    
End Type

'********** Public VARS ***********
'Where the map borders are.. Set during load
Public MinXBorder                                                                                 As Byte
Public MaxXBorder                                                                                 As Byte
Public MinYBorder                                                                                 As Byte
Public MaxYBorder                                                                                 As Byte

'Object Constants
Public Const MAX_INVENORY_OBJS                                                                    As Integer = 10000

' Deshacer
Public Const maxDeshacer                                                                          As Integer = 30
Public MapData_Deshacer(1 To maxDeshacer, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

Type tDeshacerInfo

    Libre As Boolean
    Desc As String

End Type

Public MapData_Deshacer_Info(1 To maxDeshacer) As tDeshacerInfo

'********** Public ARRAYS ***********

Public MapData()                               As MapBlock 'Holds map data for current map
Public MapInfo                                 As MapInfo 'Holds map info for current map
Public CharList(1 To 10000)                    As char 'Holds info about all characters on map

'Encabezado bmp
Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Public LastSoundBufferUsed As Integer

Public gDespX              As Integer
Public gDespY              As Integer

'User status vars
Public CurMap              As Integer 'Current map loaded
Public UserIndex           As Integer
Global UserBody            As Integer
Global UserHead            As Integer
Public UserPos             As Position 'Holds current user pos
Public AddtoUserPos        As Position 'For moving user
Public UserCharIndex       As Integer

Public EngineRun           As Boolean
Public FramesPerSec        As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth     As Integer
Public WindowTileHeight    As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop         As Integer
Public MainViewLeft        As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize      As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd     As Long

'Tile size in pixels
Public TilePixelHeight     As Integer
Public TilePixelWidth      As Integer

'Map editor variables
Public WalkMode            As Boolean

'Totals
Public NumMaps             As Integer 'Number of maps
Public Numheads            As Integer
Public NumGrhFiles         As Long 'Number of bmps
Public MaxGrhs             As Long 'Number of Grhs
Public NumChars            As Integer
Public LastChar            As Integer

'********** Direct X ***********
Public MainViewRect        As RECT
Public MainDestRect        As RECT
Public MainViewWidth       As Integer
Public MainViewHeight      As Integer
Public BackBufferRect      As RECT

'********** OUTSIDE FUNCTIONS ***********
'Good old BitBlt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Sound stuff
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'For Get and Write Var
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long

Rem Particle Groups
Public TotalStreams As Integer

'RGB Type
Public Type RGB

    r As Long
    g As Long
    b As Long

End Type

Public Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type
