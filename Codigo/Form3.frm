VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Translados Automaticos"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4080
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Izquierdo 
      Caption         =   "Izquierdo"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox Derecho 
      Caption         =   "Derecho"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox Inferior 
      Caption         =   "Inferior"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Superior 
      Caption         =   "Superior"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comenzar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblMapaActual 
      Alignment       =   2  'Center
      Caption         =   "Mapa Actual"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Command1_Click()
    'FrmMain.Timer4.Enabled = True
    
    On Error GoTo Command1_Click_Err
    
    FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm"
    FrmMain.mnuGuardarMapa_Click
    Label5.Caption = MapaActual
    Call HacerTranslados
    Label5.Caption = 0

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form3.Command1_Click", Erl)
    Resume Next
    
End Sub

Public Sub LogError(Desc As String)

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub HacerTranslados()
    
    On Error GoTo HacerTranslados_Err
    
    Label5.Caption = MapaActual

    Dim X As Integer
    Dim Y As Integer

    'Izquierda
    X = 13

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label1.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'arriba
    Y = 10

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label2.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Derecha
    X = 88

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label3.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Abajo
    Y = 91

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label4.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next

    If Superior.Value = vbChecked Then
        If CLng(Label2.Caption) = 0 Then
            Label2.Caption = MapData(49, 10).TileExit.Map

            If CLng(Label2.Caption) = 0 Then
                Call LogError("Mapa " & Label5.Caption & " sin translado")
                MsgBox "arriba cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 1
        SeleccionFX = 100
        SeleccionIY = 10
        SeleccionFY = 21
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(Label2.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label2.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName

            FrmMain.mnuReAbrirMapa.Enabled = True

        End If
    
        SobreX = 1
        SobreY = 90
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        FrmMain.mnuGuardarMapa_Click

        'Call Form_Load
    
        If CLng(Label4.Caption) = 0 Then
            Label4.Caption = MapData(49, 91).TileExit.Map

            If CLng(Label4.Caption) = 0 Then
                Call LogError("Mapa " & Label5.Caption & " sin translado")
                Exit Sub

            End If

        End If
    
        If FileExist(PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
    
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If

    'Izquierda
    X = 13

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label1.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'arriba
    Y = 10

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label2.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Derecha
    X = 88

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label3.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Abajo
    Y = 91

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label4.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next

    If Inferior.Value = vbChecked Then
        If CLng(Label4.Caption) = 0 Then
            Label4.Caption = MapData(49, 91).TileExit.Map

            If CLng(Label4.Caption) = 0 Then
                Call LogError("Mapa " & Label5.Caption & " sin translado")
                MsgBox "Abajo cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 1
        SeleccionFX = 100
        SeleccionIY = 81
        SeleccionFY = 89

        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(Label4.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label4.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        FrmMain.mnuGuardarMapa_Click
            
        Call Form_Load
                
        If FileExist(PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If

    'Izquierda
    X = 13

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label1.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'arriba
    Y = 10

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label2.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Derecha
    X = 88

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label3.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Abajo
    Y = 91

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label4.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next

    If Derecho.Value = vbChecked Then
        If CLng(Label3.Caption) = 0 Then
            Label3.Caption = MapData(88, 49).TileExit.Map

            If CLng(Label3.Caption) = 0 Then
                Call LogError("Mapa " & Label5.Caption & " sin translado")
                MsgBox "Derecha cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 76
        SeleccionFX = 87
        SeleccionIY = 1
        SeleccionFY = 100
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(Label3.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label3.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        FrmMain.mnuGuardarMapa_Click

        If FileExist(PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If
            
    'Izquierda
    X = 13

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label1.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'arriba
    Y = 10

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label2.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Derecha
    X = 88

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label3.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    
    'Abajo
    Y = 91

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label4.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next

    If Izquierdo.Value = vbChecked Then
        If CLng(Label1.Caption) = 0 Then
            Label1.Caption = MapData(12, 49).TileExit.Map

            If CLng(Label1.Caption) = 0 Then
                Call LogError("Mapa " & Label5.Caption & " sin translado")
                MsgBox "Izquierda cancelado con dos intentos"
                Exit Sub

            End If

        End If

        'ver ReyarB
        SeleccionIX = 13
        SeleccionFX = 25
        SeleccionIY = 1
        SeleccionFY = 100
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(Label1.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label1.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 88
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        FrmMain.mnuGuardarMapa_Click

        If FileExist(PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(Label5.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
                
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If
            
    Debug.Print "TERMINADO"
    Unload Me

    
    Exit Sub

HacerTranslados_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form3.HacerTranslados", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    Dim X           As Integer
    Dim Y           As Integer
    
    Dim ObtenerMapa As String
    
    ObtenerMapa = FrmMain.MapPest(4).Caption
    
    Label5.Caption = ReadField(3, ObtenerMapa, Asc("a"))
    

    X = 13

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label1.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    

    Y = 10

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label2.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    

    X = 88

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label3.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next
    

    Y = 91

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map <> 0 Then
            Label4.Caption = MapData(X, Y).TileExit.Map
            Exit For

        End If

    Next

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form3.Form_Load", Erl)
    Resume Next
    
End Sub

