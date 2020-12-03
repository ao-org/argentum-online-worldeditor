Attribute VB_Name = "ModBLoques"
Public NTIPOS      As Byte
Public RepetirSup  As Boolean

Public tipo()      As String
Public NRO()       As Byte
Public TILESX()    As Byte
Public TILESY()    As Byte
Public LAYER()     As Byte
Public Nombre()    As String
Public Grh()       As Long
Public TIPOOK()    As Byte

Public DesdeBloq   As Boolean
Public RenderX     As Byte
Public RenderY     As Byte
Public RenderGrh   As Long
Public RenderLayer As Byte

Public Sub InsertarBloque()
    
    On Error GoTo InsertarBloque_Err
    

    If FrmBloques.List1.ListIndex + 1 > 0 Then
        DesdeBloq = True
        RenderGrh = Grh(FrmBloques.List1.ListIndex + 1)
        RenderX = Val(TILESX(TIPOOK(FrmBloques.List1.ListIndex + 1)))
        RenderY = Val(TILESY(TIPOOK(FrmBloques.List1.ListIndex + 1)))
        RenderLayer = Val(LAYER(TIPOOK(FrmBloques.List1.ListIndex + 1)))

    End If

    
    Exit Sub

InsertarBloque_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModBLoques.InsertarBloque", Erl)
    Resume Next
    
End Sub

Public Sub PonerGrh()
    
    On Error GoTo PonerGrh_Err
    

    Dim j    As Integer, i As Integer, grafico As Integer
            
    Dim Cont As Integer
        
    Dim grha As Long
    grha = RenderGrh
    Sobre = RenderGrh

    If RenderLayer = 1 Then
        Suma = 0

    End If

    If RenderLayer = 2 Then
        Suma = 1

    End If

    If RenderLayer = 3 Then
        Suma = 1

    End If

    If RenderLayer = 4 Then
        Suma = 1

    End If

    For i = UltimoClickY To UltimoClickY + RenderY - 1
        For j = UltimoClickX To UltimoClickX + RenderX - 1
            MapData(j, i).Marcado = False
            MapData(j - Suma, i - Suma).Graphic(RenderLayer).grhindex = grha
            InitGrh MapData(j - Suma, i - Suma).Graphic(RenderLayer), grha
            MapData(j, i).Graphic(RenderLayer + 4).grhindex = 1
            InitGrh MapData(j, i).Graphic(RenderLayer + 4), 1

            If Cont < RenderY * RenderX Then
                Cont = Cont + 1
                grha = grha + 1

            End If

        Next
    Next

    
    Exit Sub

PonerGrh_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModBLoques.PonerGrh", Erl)
    Resume Next
    
End Sub
 
Public Sub CargarBloq()
    
    On Error GoTo CargarBloq_Err
    
   
    FrmBloques.List1.Clear
    FrmBloques.Combo1.Clear
    
    ' frmbloques.combo1.AddItem A

    NTIPOS = Val(GetVar(App.Path & "\bloq_database.txt", "CABECERA", "NTIPOS"))
    Dim i As Integer

    ReDim tipo(NTIPOS) As String
    ReDim NRO(NTIPOS) As Byte
    ReDim TILESX(NTIPOS) As Byte
    ReDim TILESY(NTIPOS) As Byte
    ReDim LAYER(NTIPOS) As Byte

    For i = 1 To NTIPOS
        tipo(i) = GetVar(App.Path & "\bloq_database.txt", "CABECERA", "Tipo" & i)
        NRO(i) = Val(GetVar(App.Path & "\bloq_database.txt", "CABECERA", "NRO" & i))
        TILESX(i) = Val(GetVar(App.Path & "\bloq_database.txt", "CABECERA", "TILESX" & i))
        TILESY(i) = Val(GetVar(App.Path & "\bloq_database.txt", "CABECERA", "TILESY" & i))
        LAYER(i) = Val(GetVar(App.Path & "\bloq_database.txt", "CABECERA", "LAYER" & i))
        Debug.Print TILESX(i)
        FrmBloques.Combo1.AddItem tipo(i)
    Next i

    FrmBloques.Show , FrmMain

    Exit Sub

    
    Exit Sub

CargarBloq_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModBLoques.CargarBloq", Erl)
    Resume Next
    
End Sub

Public Sub CargarTipo(ByVal Indice As Byte)
    
    On Error GoTo CargarTipo_Err
    

    Dim NITEMS As Byte
   
    FrmBloques.List1.Clear

    NITEMS = Val(GetVar(App.Path & "\bloq_database.txt", tipo(Indice), "NITEMS"))
    Dim i As Integer

    ReDim Nombre(NITEMS) As String
    ReDim Grh(NITEMS) As Long
    ReDim TIPOOK(NITEMS) As Byte

    For i = 1 To NITEMS
        Nombre(i) = GetVar(App.Path & "\bloq_database.txt", tipo(Indice), "NOMBRE" & i)
        Grh(i) = Val(GetVar(App.Path & "\bloq_database.txt", tipo(Indice), "GRH" & i))
        TIPOOK(i) = Val(GetVar(App.Path & "\bloq_database.txt", tipo(Indice), "TIPO" & i))
        FrmBloques.List1.AddItem Nombre(i)
    Next i

    Exit Sub

    
    Exit Sub

CargarTipo_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModBLoques.CargarTipo", Erl)
    Resume Next
    
End Sub
