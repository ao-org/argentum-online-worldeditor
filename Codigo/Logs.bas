Attribute VB_Name = "Logs"
Option Explicit

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type

Private HistorialError As UltimoError

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
On Error GoTo EH:

    'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
    If Componente = HistorialError.Componente And _
       Numero = HistorialError.ErrorCode Then
        
        'Agregamos el error al historial.
        HistorialError.Contador = HistorialError.Contador + 1
        
        'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
        'x lo que no hace falta registrar el error.
        If HistorialError.Contador = 10 Then Exit Sub
        
    Else 'Si NO es igual, reestablecemos el contador.

        HistorialError.Contador = 0
        HistorialError.ErrorCode = Numero
        HistorialError.Componente = Componente
            
    End If
    
    'Registramos el error en Errores.log
    Dim File As Integer: File = FreeFile
        
    Open App.Path & "\logs\Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
    
    Exit Sub
    
EH:
    
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
    
End Sub

