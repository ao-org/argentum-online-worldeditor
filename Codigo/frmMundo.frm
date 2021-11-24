VERSION 5.00
Begin VB.Form FrmMundo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reenderizador de mundo por Ladder"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTransladar 
      Caption         =   "Transladar"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   300
      Width           =   1095
   End
   Begin VB.CheckBox vacio 
      Caption         =   "Vacio"
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   300
      Width           =   735
   End
   Begin VB.CheckBox chkAgua 
      Caption         =   "Agua"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton cmdExportarPlano 
      Caption         =   "Exportar mapa"
      Height          =   360
      Left            =   5160
      TabIndex        =   10
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabarWorldData 
      Caption         =   "Grabar WorldData"
      Height          =   360
      Left            =   5160
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Left            =   120
      ScaleHeight     =   594
      ScaleMode       =   0  'User
      ScaleWidth      =   576
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   7680
      Begin VB.Image imgSwitchWorld 
         Height          =   435
         Index           =   1
         Left            =   8790
         Tag             =   "0"
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lblPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   8220
         Width           =   60
      End
      Begin VB.Shape lblAllies 
         BorderColor     =   &H000000C0&
         FillColor       =   &H0000FFFF&
         Height          =   405
         Left            =   1920
         Top             =   2880
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.TextBox txtalto 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtancho 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMundo.frx":0000
      Left            =   120
      List            =   "frmMundo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Renderizar"
      Height          =   360
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblMapanumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mapanumero"
      Height          =   195
      Left            =   3960
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblmapa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   75
      Width           =   2115
   End
   Begin VB.Label lblAlto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblAncho 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblMundo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mundo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmMundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const TILE_SIZE = 27

Public Sub cmdCommand1_Click()

Dim precarga As Picture
Dim Pos_x As Integer
Dim Pos_y As Integer

'de donde saco los mapas?

Dim i As Integer


For i = 1 To Mundo(MundoSeleccionado).Alto * Mundo(MundoSeleccionado).Ancho

    Set precarga = LoadPicture(App.Path & "\render\mapa" & Mundo(MundoSeleccionado).MapIndice(i) & ".bmp")
    
    picture1.PaintPicture precarga, Pos_x, Pos_y
    Pos_x = Pos_x + TILE_SIZE
    

    If i Mod Mundo(MundoSeleccionado).Ancho = 0 Then
        Pos_y = Pos_y + TILE_SIZE
        Pos_x = 0
    End If
Next i

End Sub

Private Sub cmdExportarPlano_Click()

Dim Imagen As IPictureDisp
Set Imagen = picture1.image

SavePicture Imagen, App.Path & "\recursos\exportacion\mapa" & MundoSeleccionado & ".bmp"
SavePicture Imagen, App.Path & "\..\Recursos\interface\mapa" & MundoSeleccionado & ".bmp"

MsgBox ("Se han guardado los cambios.")

Set Imagen = Nothing
End Sub

Private Sub cmdGrabarWorldData_Click()
Dim i As Integer
Dim j As Byte
Dim Manager  As clsIniManager
Set Manager = New clsIniManager
Call Manager.Initialize(MapFile)

Call Manager.ChangeValue("INIT", "TotalWorlds", TotalWorlds)

For j = 1 To TotalWorlds

    Call Manager.ChangeValue("WORLDMAP" & j, "Ancho", Mundo(j).Ancho)
    Call Manager.ChangeValue("WORLDMAP" & j, "Alto", Mundo(j).Alto)
    
    For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
    
        Call Manager.ChangeValue("WORLDMAP" & j, i, Mundo(j).MapIndice(i))
    
    Next i
Next j

MsgBox "¡Archivo guardado!"
Call Manager.DumpFile(MapFile)


HuboCambios = False
    
Set Manager = Nothing
End Sub

Private Sub Combo1_Click()
MundoSeleccionado = FrmMundo.Combo1.ListIndex + 1
FrmMundo.txtancho = Mundo(MundoSeleccionado).Ancho
FrmMundo.txtalto = Mundo(MundoSeleccionado).Alto

picture1.BackColor = vbBlack
'picture1.Refresh
picture1.ScaleMode = 3
picture1.Height = Mundo(MundoSeleccionado).Alto * TILE_SIZE
picture1.Width = Mundo(MundoSeleccionado).Ancho * TILE_SIZE
FrmMundo.cmdCommand1_Click
End Sub

Private Sub Form_Load()
Call CargarWorldData
picture1.ScaleMode = 3
picture1.Height = Mundo(MundoSeleccionado).Alto * TILE_SIZE
picture1.Width = Mundo(MundoSeleccionado).Ancho * TILE_SIZE

picture1.ScaleHeight = Mundo(MundoSeleccionado).Alto * TILE_SIZE
picture1.ScaleWidth = Mundo(MundoSeleccionado).Ancho * TILE_SIZE
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift

    

    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If HuboCambios Then
        If MsgBox("¡No se han guardado los cambios!" & vbNewLine & "¿Seguro desea salir?", vbYesNo, "¡NO SE GUARDO!") = vbNo Then
             Cancel = True
        End If
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblAllies.Visible = True
    Dim POSX As Integer
    Dim PosY As Integer
    Dim Mapa As Integer
    
    MouseBoton = Button
    MouseShift = Shift
    
    
    ' Para obtener las coordenadas (x, y) del "slot" divido la posición del cursor
    ' por el tamaño de los tiles y me quedo solo con la parte entera
    POSX = Int(X / TILE_SIZE) ' PosX = Valor entero entre 0 y (MAPAS_ANCHO - 1)
    PosY = Int(y / TILE_SIZE) ' PosY = Valor entero entre 0 y (MAPAS_ALTO - 1)
    
    ' Uso estas coordeandas para calcular el índice del mapa
    Mapa = POSX + PosY * Mundo(MundoSeleccionado).Ancho + 1  ' +1 porque los mapas empiezan en 1
    
    ' Luego multiplico por TILE_SIZE para tener la posición final en donde poner el indicador
    POSX = POSX * TILE_SIZE
    PosY = PosY * TILE_SIZE
    
    lblAllies.Top = PosY
    lblAllies.Left = POSX
    
    
    lblmapa.Caption = "Mundo Nº: " & MundoSeleccionado & " - Mapa: " & Mundo(MundoSeleccionado).MapIndice(Mapa)

    lblMapanumero.Caption = Mundo(MundoSeleccionado).MapIndice(Mapa)
    
    If MouseBoton = 4 Then
        
        Mapa = lblMapanumero.Caption
        FileName = PATH_Save & NameMap_Save & Mapa & ".csm"

        If FileExist(FileName, vbArchive) = False Then Exit Sub
        Call modMapIO.NuevoMapa
        DoEvents
        modMapIO.AbrirMapa FileName
        EngineRun = True
        Exit Sub
    End If
    
    If chkTransladar And Not Translando Then
        Translando = True
        TempMapa = Mundo(MundoSeleccionado).MapIndice(Mapa)
        TempIndice = Mapa
        lblmapa.Caption = "Seleccione destino..."
        lblAllies.BorderColor = vbYellow
        Exit Sub
    End If
    
    If Translando Then
            NuevoIndice = TempMapa
            Mundo(MundoSeleccionado).MapIndice(TempIndice) = Mundo(MundoSeleccionado).MapIndice(Mapa)
            Mundo(MundoSeleccionado).MapIndice(Mapa) = TempMapa
            Call cmdCommand1_Click
            HuboCambios = True
            Translando = False
            lblAllies.BorderColor = vbRed
    End If
    
    If vacio Then
            NuevoIndice = 300
            Mundo(MundoSeleccionado).MapIndice(Mapa) = NuevoIndice
            Call cmdCommand1_Click
            HuboCambios = True
        Exit Sub
    End If
    
        
    If chkAgua Then
            NuevoIndice = 0
            Mundo(MundoSeleccionado).MapIndice(Mapa) = NuevoIndice
            Call cmdCommand1_Click
            HuboCambios = True
        Exit Sub
    End If
    
    
    
    If Button = 2 And MouseShift = 1 Then
        NuevoIndice = 300
            Mundo(MundoSeleccionado).MapIndice(Mapa) = NuevoIndice
            Call cmdCommand1_Click
            HuboCambios = True
    
        Exit Sub
    End If
    
    
    If Button = 2 Then
        NuevoIndice = InputBox("Ingrese nuevo indice de mapa:")
        If IsNumeric(NuevoIndice) Then
            Mundo(MundoSeleccionado).MapIndice(Mapa) = NuevoIndice
            Call cmdCommand1_Click
            HuboCambios = True
        Else
            MsgBox "No valido"
        End If
    End If
    

End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   ' picture1.ScaleWidth = 432
   ' picture1.ScaleHeight = 594
    
   ' picture1.Width = 432
   ' picture1.Height = 594
End Sub


