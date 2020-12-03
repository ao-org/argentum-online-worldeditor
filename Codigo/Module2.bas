Attribute VB_Name = "Module2"
Option Explicit

'--------------------------------------------
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com.ar
'Date: 01/11/2009
'--------------------------------------------
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long

Private Type GUID

    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte

End Type

Private Type EncoderParameter

    GUID            As GUID
    NumberOfValues  As Long

    type            As Long

    value           As Long

End Type

Private Type EncoderParameters

    Count           As Long
    Parameter(15)   As EncoderParameter

End Type

Private Type GdiplusStartupInput

    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long

End Type

Const ImageCodecBMP = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Const ImageCodecJPG = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Const ImageCodecGIF = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Const ImageCodecTIF = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
Const ImageCodecPNG = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"

Const EncoderQuality = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Const EncoderCompression = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"

Const TiffCompressionNone = 6
Const EncoderParameterValueTypeLong = 4

Public Function ConvertFileImage(ByVal SrcPath As String, ByVal DestPath As String, Optional ByVal JPG_Quality As Long = 85) As Boolean
                                 
    On Error Resume Next

    Dim GDIsi    As GdiplusStartupInput, gToken As Long, hBitmap As Long
    Dim tEncoder As GUID
    Dim tParams  As EncoderParameters
    Dim sExt     As String
    Dim lPos     As Long

    DestPath = Trim(DestPath)
        
    lPos = InStrRev(DestPath, ".")

    If lPos Then
        sExt = UCase(Right(DestPath, Len(DestPath) - lPos))

    End If

    Select Case sExt

        Case "PNG"
            CLSIDFromString StrPtr(ImageCodecPNG), tEncoder

        Case "TIF", "TIFF"
            CLSIDFromString StrPtr(ImageCodecTIF), tEncoder
            
            With tParams
                .Count = 1
                .Parameter(0).NumberOfValues = 1
                .Parameter(0).type = EncoderParameterValueTypeLong
                .Parameter(0).value = VarPtr(TiffCompressionNone)
                CLSIDFromString StrPtr(EncoderCompression), .Parameter(0).GUID

            End With
            
        Case "BMP", "DIB"
            CLSIDFromString StrPtr(ImageCodecBMP), tEncoder
        
        Case "GIF"
            CLSIDFromString StrPtr(ImageCodecGIF), tEncoder
        
        Case "JPG", "JPEG", "JPE", "JFIF"
            
            If JPG_Quality > 100 Then JPG_Quality = 100
            If JPG_Quality < 0 Then JPG_Quality = 0
            
            CLSIDFromString StrPtr(ImageCodecJPG), tEncoder
            
            With tParams
                .Count = 1
                .Parameter(0).NumberOfValues = 1
                .Parameter(0).type = EncoderParameterValueTypeLong
                .Parameter(0).value = VarPtr(JPG_Quality)
                CLSIDFromString StrPtr(EncoderQuality), .Parameter(0).GUID

            End With

        Case Else
            Exit Function
            
    End Select

    GDIsi.GdiplusVersion = 1&
    
    GdiplusStartup gToken, GDIsi

    If gToken Then
  
        If GdipLoadImageFromFile(StrPtr(SrcPath), hBitmap) = 0 Then
    
            If GdipSaveImageToFile(hBitmap, StrPtr(DestPath), tEncoder, ByVal tParams) = 0 Then
                ConvertFileImage = True

            End If
            
            GdipDisposeImage hBitmap
    
        End If
    
        GdiplusShutdown gToken

    End If
    
End Function

Public Function IsGdiPlusInstaled() As Boolean
    
    On Error GoTo IsGdiPlusInstaled_Err
    
    Dim hLib As Long
    
    hLib = LoadLibrary("gdiplus.dll")

    If hLib Then
        If GetProcAddress(hLib, "GdiplusStartup") Then
            IsGdiPlusInstaled = True

        End If

        FreeLibrary hLib

    End If
    
    
    Exit Function

IsGdiPlusInstaled_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module2.IsGdiPlusInstaled", Erl)
    Resume Next
    
End Function

Public Sub Engine_Convert_List(rgb_list() As Long, Long_Color As Long)
    
    On Error GoTo Engine_Convert_List_Err
    

    ' / Author: Dunkansdk
    ' / Note: Convierte en array's los D3DColorArgb

    rgb_list(0) = Long_Color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
    
    
    Exit Sub

Engine_Convert_List_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module2.Engine_Convert_List", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Box(ByVal X As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long)
    
    On Error GoTo Engine_Draw_Box_Err
    

    ' / Author: Ezequiel Juárez (Standelf)
    ' / Note: Extract to Blisse AO, modified by Dunkansdk

    Dim b_Rect           As RECT
    Dim b_Color(0 To 3)  As Long
    Dim b_Vertex(0 To 3) As TLVERTEX
    
    With b_Rect
        .Bottom = y + Height
        .Left = X
        .Right = X + Width
        .Top = y

    End With

    Engine_Convert_List b_Color(), color

    Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))

    
    Exit Sub

Engine_Draw_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module2.Engine_Draw_Box", Erl)
    Resume Next
    
End Sub

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long)
    
    On Error GoTo Geometry_Create_Box_Err
    

    ' / Author: Dunkansdk

    ' * v0      * v1
    ' |        /|
    ' |      /  |
    ' |    /    |
    ' |  /      |
    ' |/        |
    ' * v2      * v3

    Dim x_Cor As Single
    Dim y_Cor As Single
    
    ' * - - - - - - - Vertice 0 -
    x_Cor = dest.Left
    y_Cor = dest.Bottom
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom) / Textures_Height)
    Else
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)

    End If

    ' * - - - - - - - Vertice 0 -
    
    ' * - - - - - - - Vertice 1 -
    x_Cor = dest.Left
    y_Cor = dest.Top
       
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)

    End If

    ' * - - - - - - - Vertice 1 -

    ' * - - - - - - - Vertice 2 -
    x_Cor = dest.Right
    y_Cor = dest.Bottom
    
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / Textures_Width, (src.Bottom) / Textures_Height)
    Else
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)

    End If

    ' * - - - - - - - Vertice 2 -
    
    ' * - - - - - - - Vertice 3 -
    x_Cor = dest.Right
    y_Cor = dest.Top
    
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

    End If

    ' * - - - - - - - Vertice 3 -

    
    Exit Sub

Geometry_Create_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module2.Geometry_Create_Box", Erl)
    Resume Next
    
End Sub

Public Function CreateVertex(ByVal X As Single, ByVal y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, ByVal tv As Single) As TLVERTEX
    
    On Error GoTo CreateVertex_Err
    

    ' / Author: Aaron Perkins
    ' / Last Modify Date: 10/07/2002

    CreateVertex.X = X
    CreateVertex.y = y
    CreateVertex.Z = Z
    CreateVertex.rhw = rhw
    CreateVertex.color = color
    CreateVertex.Specular = Specular
    CreateVertex.tu = tu
    CreateVertex.tv = tv
    
    
    Exit Function

CreateVertex_Err:
    Call RegistrarError(Err.Number, Err.Description, "Module2.CreateVertex", Erl)
    Resume Next
    
End Function

