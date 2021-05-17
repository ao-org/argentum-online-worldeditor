VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Translados manual"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pegar"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Pegar"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pegar"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pegar"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copiar al Sur"
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar al Norte"
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copiar al Este"
      Height          =   2055
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copiar al Oeste"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblMapaOeste 
      Caption         =   "Oeste"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMapaEste 
      Alignment       =   1  'Right Justify
      Caption         =   "Este"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblMapaSur 
      Alignment       =   2  'Center
      Caption         =   "Sur"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblMapaNorte 
      Alignment       =   2  'Center
      Caption         =   "Norte"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPIAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Command1_Click()
    'Superior
    
    On Error GoTo Command1_Click_Err
    
    'Call VerMapaTraslado

    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 11
    SeleccionFY = 22
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = True
    Command6.Visible = False
    Command7.Visible = False
    Command8.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    FrmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command1_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command2_Click()
    'copiar izquierdo
    
    On Error GoTo Command2_Click_Err
    
    SeleccionIX = 14
    SeleccionFX = 27
    SeleccionIY = 1
    SeleccionFY = 100
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    Command6.Visible = False
    Command7.Visible = True
    Command8.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    FrmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command2_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command3_Click()
    'copiar derecho
    
    On Error GoTo Command3_Click_Err
    
    SeleccionIX = 75
    SeleccionFX = 87
    SeleccionIY = 1
    SeleccionFY = 100
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    Command6.Visible = False
    Command7.Visible = False
    Command8.Visible = True
    Call CopiarSeleccion
    MapInfo.Changed = 1
    FrmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command3_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command4_Click()
    'Copiar inferior ok!
    
    On Error GoTo Command4_Click_Err
    
    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 81
    SeleccionFY = 90
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    Command6.Visible = True
    Command7.Visible = False
    Command8.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    FrmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command4_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command5_Click()
    
    On Error GoTo Command5_Click_Err
    

    'Pegar Inferior OK!
    If lblMapaNorte.Caption <> "Norte" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
    
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 91
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.y = 87
        Unload Me

    End If

    
    Exit Sub

Command5_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command5_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command6_Click()
    'pegar superior
    
    On Error GoTo Command6_Click_Err
    

    If lblMapaSur.Caption <> "Sur" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
    
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.y = 14
        Unload Me

    End If

    
    Exit Sub

Command6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command6_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command7_Click()
    
    On Error GoTo Command7_Click_Err
    

    'pegar derecho OK!
    If lblMapaOeste.Caption <> "Oeste" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
    
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 88
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.X = 83
        Unload Me

    End If

    
    Exit Sub

Command7_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command7_Click", Erl)
    Resume Next
    
End Sub

Public Sub Command8_Click()
    
    On Error GoTo Command8_Click_Err
    

    'pegar izquierdo OK!
    If lblMapaEste.Caption <> "Este" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            FrmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm"
            modMapIO.AbrirMapa FrmMain.Dialog.FileName
    
            FrmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.X = 19
        Unload Me

    End If

    
    Exit Sub

Command8_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command8_Click", Erl)
    Resume Next
    
End Sub

Private Sub VerMapaTraslado()
    
    On Error GoTo VerMapaTraslado_Err
    
    Dim X As Integer
    Dim y As Integer

    'Izquierda
    X = 13

    For y = (11) To (90)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaOeste.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'arriba
    y = 10

    For X = (14) To (87)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaNorte.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'Derecha
    X = 88

    For y = (11) To (90)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaEste.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'Abajo
    y = 91

    For X = (14) To (87)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaSur.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next

    
    Exit Sub

VerMapaTraslado_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.VerMapaTraslado", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call VerMapaTraslado

    If lblMapaSur.Caption = "Sur" Then Form2.Command4.Visible = False
    If lblMapaEste.Caption = "Este" Then Form2.Command3.Visible = False
    If lblMapaOeste.Caption = "Oeste" Then Form2.Command2.Visible = False
    If lblMapaOeste.Caption = "Norte" Then Form2.Command1.Visible = False

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Form_Load", Erl)
    Resume Next
    
End Sub
