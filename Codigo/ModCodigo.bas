Attribute VB_Name = "ModCodigo"
Public MapFile As String


Public Translando As Boolean
Public TempMapa As Integer
Public TempIndice As Integer

Public MouseBoton As Long
Public MouseShift As Long
Public HuboCambios As Boolean

Public TotalWorlds As Byte
Public MundoSeleccionado As Byte


Public Type WorldMap
    MapIndice() As Integer
    Ancho As Integer
    Alto As Integer
End Type

Public Mundo() As WorldMap


Public Sub CargarWorldData()

On Error GoTo Cargarmapsworlddata_Err
    

    'Ladder
    

    Dim i       As Integer
    Dim j       As Byte


    MapFile = App.Path & "\..\Recursos\init\mapsworlddata.dat"

    Dim Leer As New clsIniManager
    Call Leer.Initialize(MapFile)
    
    
    TotalWorlds = Val(Leer.GetValue("INIT", "TotalWorlds"))
       
    ReDim Mundo(1 To TotalWorlds) As WorldMap
   
    For j = 1 To TotalWorlds
        FrmMundo.Combo1.AddItem (j)
        Mundo(j).Alto = Val(Leer.GetValue("WORLDMAP" & j, "Alto"))
        Mundo(j).Ancho = Val(Leer.GetValue("WORLDMAP" & j, "Ancho"))

        ReDim Mundo(j).MapIndice(1 To Mundo(j).Alto * Mundo(j).Ancho) As Integer
         
         For i = 1 To Mundo(j).Alto * Mundo(j).Ancho
             Mundo(j).MapIndice(i) = Val(Leer.GetValue("WORLDMAP" & j, i))
         Next i
         
     Next j
     
     
     MundoSeleccionado = 1
     FrmMundo.Combo1.ListIndex = 0
     FrmMundo.txtancho = Mundo(MundoSeleccionado).Ancho
     FrmMundo.txtalto = Mundo(MundoSeleccionado).Alto
     
      FrmMundo.cmdCommand1_Click
    Exit Sub

Cargarmapsworlddata_Err:
   ' Call RegistrarError(Err.Number, Err.Description, "Recursos.Cargarmapsworlddata", Erl)
    Resume Next
    
End Sub
