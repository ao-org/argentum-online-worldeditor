Attribute VB_Name = "ModAmbientaciones"
Option Explicit
Public AmbientacionesTotal As Integer

Public Type Ambientacion

    Nombre As String
    tipo As Byte
    grhindex As Long

End Type

Public Ambientaciones(1 To 2000) As Ambientacion

Sub LeerAmbientaciones()
    
    On Error GoTo LeerAmbientaciones_Err
    
    Dim FilePath As String
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT", "ambientacion.ini", Windows_Temp_Dir, False) Then
            MsgBox "Se projudo un error al intentar cargar ambientaciones.ini"
            Exit Sub
        Else
            FilePath = Windows_Temp_Dir & "ambientacion.ini"

        End If

    #Else
        FilePath = App.Path & "\..\Recursos\init\ambientacion.ini"
    #End If
    
    AmbientacionesTotal = Val(GetVar(FilePath, "INIT", "AmbientacionTotales"))

    Dim i As Integer

    For i = 1 To AmbientacionesTotal
        Ambientaciones(i).Nombre = GetVar(FilePath, Val(i), "Nombre")
        Ambientaciones(i).tipo = GetVar(FilePath, Val(i), "Tipo")
        Ambientaciones(i).grhindex = GetVar(FilePath, Val(i), "GrhIndex")
    Next i
    
    #If Compresion = 1 Then
        Kill FilePath
    #End If
    
    
    Exit Sub

LeerAmbientaciones_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModAmbientaciones.LeerAmbientaciones", Erl)
    Resume Next
    
End Sub

