Attribute VB_Name = "ModLadder"
Public Working             As Boolean
Public MMiniMap_capa1      As Boolean
Public MMiniMap_capa2      As Boolean
Public MMiniMap_capa3      As Boolean
Public MMiniMap_capa4      As Boolean
Public MMiniMap_Npcs       As Boolean
Public MMiniMap_objetos    As Boolean
Public MMiniMap_Bloqueos   As Boolean
Public MMiniMap_particulas As Boolean
Public MMiniMap_Nombre     As Boolean

Public ToWorldMap2         As Boolean
Public Radio               As Byte

Public MapaActual          As Integer

'Compresion
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_LENGTH = 512
Option Explicit
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
    
    On Error GoTo SetTopMostWindow_Err
    

    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False

    End If

    
    Exit Function

SetTopMostWindow_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.SetTopMostWindow", Erl)
    Resume Next
    
End Function

Public Sub Obtener_RGB(ByVal color As Long, Rojo As Byte, Verde As Byte, Azul As Byte)
    
    On Error GoTo Obtener_RGB_Err
    
    
    Azul = (color And 16711680) / 65536
    Verde = (color And 65280) / 256
    Rojo = color And 255
  
    
    Exit Sub

Obtener_RGB_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.Obtener_RGB", Erl)
    Resume Next
    
End Sub

Public Function General_Get_Temp_Dir() As String
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    'Gets windows temporary directory
    '**************************************************************
    
    On Error GoTo General_Get_Temp_Dir_Err
    
    Dim s As String
    Dim c As Long
    s = Space$(MAX_LENGTH)
    c = GetTempPath(MAX_LENGTH, s)

    If c > 0 Then
        If c > Len(s) Then
            s = Space$(c + 1)
            c = GetTempPath(MAX_LENGTH, s)

        End If

    End If

    General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")

    
    Exit Function

General_Get_Temp_Dir_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.General_Get_Temp_Dir", Erl)
    Resume Next
    
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    
    On Error GoTo AddtoRichTextBox_Err
    

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    With RichTextBox

        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
    End With

    
    Exit Sub

AddtoRichTextBox_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.AddtoRichTextBox", Erl)
    Resume Next
    
End Sub

