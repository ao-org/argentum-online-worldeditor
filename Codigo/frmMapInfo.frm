VERSION 5.00
Begin VB.Form frmMapInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del Mapa"
   ClientHeight    =   9360
   ClientLeft      =   11505
   ClientTop       =   3030
   ClientWidth     =   4395
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   4395
   Begin VB.TextBox txtMapVersion 
      Height          =   285
      Left            =   3240
      TabIndex        =   139
      Text            =   "MapVersion"
      Top             =   8280
      Width           =   615
   End
   Begin VB.CheckBox chkMapPK 
      Caption         =   "MapPK"
      Height          =   210
      Left            =   240
      TabIndex        =   138
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdIraZona 
      Caption         =   "Ir a Zona"
      Height          =   375
      Left            =   120
      TabIndex        =   137
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtIrZona 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1200
      TabIndex        =   136
      Text            =   "1"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtMapNivelMinimo 
      Height          =   285
      Left            =   3240
      TabIndex        =   135
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox txtMapNivelMaximo 
      Height          =   285
      Left            =   3240
      TabIndex        =   133
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox tAX2 
      Height          =   285
      Left            =   6720
      TabIndex        =   127
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox tAY2 
      Height          =   285
      Left            =   6720
      TabIndex        =   126
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox tAY1 
      Height          =   285
      Left            =   5040
      TabIndex        =   125
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox tAX1 
      Height          =   285
      Left            =   5040
      TabIndex        =   124
      Top             =   2760
      Width           =   615
   End
   Begin VB.CheckBox chkTieneNpcInvocacion 
      Caption         =   "Tiene Invocacion de NPCs"
      Height          =   210
      Left            =   240
      TabIndex        =   123
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CheckBox chkMapNoEncriptarMP 
      Caption         =   "No Encriptar MP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   122
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtMapMusica 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   120
      Text            =   "0"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   4
      Left            =   6360
      Picture         =   "frmMapInfo.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   6360
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   117
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command39 
      Height          =   255
      Left            =   7200
      Picture         =   "frmMapInfo.frx":67E6
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command38 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":6D42
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command37 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":729E
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   18
      Left            =   6720
      Picture         =   "frmMapInfo.frx":77FA
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   2
      Left            =   6360
      Picture         =   "frmMapInfo.frx":7D56
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   5640
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   110
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command36 
      Height          =   255
      Left            =   7200
      Picture         =   "frmMapInfo.frx":82B2
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command35 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":880E
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command34 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":8D6A
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   16
      Left            =   6720
      Picture         =   "frmMapInfo.frx":92C6
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   3
      Left            =   6360
      Picture         =   "frmMapInfo.frx":9822
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   103
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Height          =   255
      Left            =   7200
      Picture         =   "frmMapInfo.frx":9D7E
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":A2DA
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":A836
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   14
      Left            =   6720
      Picture         =   "frmMapInfo.frx":AD92
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   0
      Left            =   6360
      Picture         =   "frmMapInfo.frx":B2EE
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   94
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Height          =   255
      Left            =   7200
      Picture         =   "frmMapInfo.frx":B84A
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":BDA6
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":C302
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   12
      Left            =   6720
      Picture         =   "frmMapInfo.frx":C85E
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   1
      Left            =   6360
      Picture         =   "frmMapInfo.frx":CDBA
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox tSonidos 
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   87
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Height          =   255
      Left            =   7200
      Picture         =   "frmMapInfo.frx":D316
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Height          =   255
      Left            =   8160
      Picture         =   "frmMapInfo.frx":D872
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command17 
      Height          =   255
      Left            =   7680
      Picture         =   "frmMapInfo.frx":DDCE
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton btnwav 
      Height          =   255
      Index           =   10
      Left            =   6720
      Picture         =   "frmMapInfo.frx":E32A
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   7680
      TabIndex        =   77
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   4800
      TabIndex        =   76
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox tNPC 
      Height          =   285
      Left            =   5160
      TabIndex        =   75
      Text            =   "1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Listar"
      Height          =   375
      Left            =   4800
      TabIndex        =   74
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3720
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   7800
      TabIndex        =   72
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox lstNpc 
      Height          =   1425
      Left            =   4560
      TabIndex        =   71
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Guardar Area"
      Height          =   375
      Left            =   5880
      TabIndex        =   70
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox tCant 
      Height          =   285
      Left            =   6480
      TabIndex        =   69
      Text            =   "1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Modifica"
      Height          =   375
      Left            =   7800
      TabIndex        =   68
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":E886
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":EDE2
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":F33E
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   9
      Left            =   2520
      Picture         =   "frmMapInfo.frx":F89A
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":FDF6
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":10352
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":108AE
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   8
      Left            =   2520
      Picture         =   "frmMapInfo.frx":10E0A
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":11366
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":118C2
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":11E1E
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   7
      Left            =   2520
      Picture         =   "frmMapInfo.frx":1237A
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":128D6
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":12E32
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":1338E
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   6
      Left            =   2520
      Picture         =   "frmMapInfo.frx":138EA
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   5
      Left            =   2520
      Picture         =   "frmMapInfo.frx":13E46
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   3480
      Picture         =   "frmMapInfo.frx":143A2
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMapInfo.frx":148FE
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   3000
      Picture         =   "frmMapInfo.frx":14E5A
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Guardar Zona"
      Height          =   375
      Left            =   2400
      TabIndex        =   32
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000004&
      Caption         =   "Segura"
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtMapNombre 
      Height          =   285
      Left            =   1680
      TabIndex        =   30
      Text            =   "Nombre"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   29
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   28
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   27
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   26
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   0
      Left            =   2160
      Picture         =   "frmMapInfo.frx":153B6
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   1
      Left            =   2160
      Picture         =   "frmMapInfo.frx":15912
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   2
      Left            =   2160
      Picture         =   "frmMapInfo.frx":15E6E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   3
      Left            =   2160
      Picture         =   "frmMapInfo.frx":163CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Height          =   255
      Index           =   4
      Left            =   2160
      Picture         =   "frmMapInfo.frx":16926
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   255
   End
   Begin VB.TextBox tZY2 
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox tZX2 
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox tZY1 
      Height          =   285
      Left            =   600
      TabIndex        =   17
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox tZX1 
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox chkOcultarNombre 
      BackColor       =   &H80000004&
      Caption         =   "OcultarNombre"
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   2640
      TabIndex        =   15
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtNiebla 
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TxtR 
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox TxtG 
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "0"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox TxtB 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CheckBox chkMapResuSinEfecto 
      Caption         =   "ResuSinEfecto"
      Height          =   210
      Left            =   2640
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox chkMapInviSinEfecto 
      Caption         =   "InviSinEfecto"
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   6720
      Width           =   1455
   End
   Begin VB.PictureBox cmdCerrar 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   7200
      ScaleHeight     =   225
      ScaleWidth      =   1305
      TabIndex        =   8
      Top             =   480
      Width           =   1365
   End
   Begin VB.ComboBox txtMapRestringir 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":16E82
      Left            =   1680
      List            =   "frmMapInfo.frx":16E8C
      TabIndex        =   6
      Text            =   "No"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox txtMapTerreno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      HelpContextID   =   1
      ItemData        =   "frmMapInfo.frx":16E98
      Left            =   1680
      List            =   "frmMapInfo.frx":16EA5
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox txtMapZona 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":16EC2
      Left            =   1680
      List            =   "frmMapInfo.frx":16EDE
      TabIndex        =   4
      Text            =   "1"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   7920
      Width           =   855
   End
   Begin VB.CheckBox chkMapMagiaSinEfecto 
      Caption         =   "Magia Sin Efecto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2640
      TabIndex        =   0
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox cmdMusica 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   7200
      ScaleHeight     =   225
      ScaleWidth      =   1305
      TabIndex        =   121
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label lblMusicaZona 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Musica Zona"
      Height          =   195
      Left            =   120
      TabIndex        =   141
      Top             =   3840
      Width           =   930
   End
   Begin VB.Label lblMapaVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Version"
      Height          =   195
      Left            =   1800
      TabIndex        =   140
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label lblMapNivelMinimo 
      Caption         =   "Nivel Minimo"
      Height          =   255
      Left            =   1800
      TabIndex        =   134
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label lblMapNivelMaximo 
      Caption         =   "NivelMaximo"
      Height          =   210
      Left            =   1800
      TabIndex        =   132
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Y2"
      Height          =   255
      Left            =   6240
      TabIndex        =   131
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "X2"
      Height          =   255
      Left            =   6240
      TabIndex        =   130
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "Y1"
      Height          =   375
      Left            =   4560
      TabIndex        =   129
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "X1"
      Height          =   255
      Left            =   4560
      TabIndex        =   128
      Top             =   2760
      Width           =   375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   2760
      X2              =   120
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   2760
      X2              =   2760
      Y1              =   2160
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   2760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 5"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   119
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 3"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   112
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 4"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   105
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "FX del Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   98
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label11 
      Caption         =   "Musica del Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   97
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 1"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   12
      Left            =   4440
      TabIndex        =   96
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido 2"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   89
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblPos 
      BackColor       =   &H80000004&
      Caption         =   "(0,0)"
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   6000
      TabIndex        =   82
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblNum 
      BackColor       =   &H80000004&
      Caption         =   "Area Nº: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   4560
      TabIndex        =   81
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "Npc Areas"
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4560
      TabIndex        =   80
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "NPC:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4560
      TabIndex        =   79
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Cant:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5880
      TabIndex        =   78
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "B"
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "G"
      Height          =   255
      Left            =   3240
      TabIndex        =   45
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R"
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 1"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   43
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 2"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   42
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 3"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 4"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   40
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 5"
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   39
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   38
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   37
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   36
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   35
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label tNumZona 
      BackColor       =   &H80000004&
      Caption         =   "Zona Nº:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   2040
      TabIndex        =   34
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "Niebla"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2880
      TabIndex        =   33
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   3960
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label Label5 
      Caption         =   "Restringir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Zona:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Terreno Tipo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private Sub btnMidi_Click(Index As Integer)

    Dim ret As Integer

    Dim Num As Integer

    Num = Val(frmMapInfo.tMusica(Index).Text)

    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
    Else
        ret = mciSendString("open " & """" & App.Path & "\..\Recursos\midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If

End Sub

Private Sub btnwav_Click(Index As Integer)

'Dim ret As Integer
'
'Dim Num As Integer
'
'Num = Val(frmMapInfo.tSonidos(Index).Text)
'
'If IsPlaying Then
'   ret = mciSendString("close mus", 0&, 0, 0)
'   IsPlaying = False
'Else
'   ret = mciSendString("open " & """" & App.Path & "\Wav\" & Num & ".wav" & """" & " type sequencer alias mus", 0&, 0, 0)
'   ret = mciSendString("play mus", 0&, 0, 0)
'   IsPlaying = True
'End If


   ' Close CANYON.MID file and sequencer device
End Sub

Private Sub chkMapBackup_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.BackUp = chkMapBackup.Value
'MapInfo.Changed = 1
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.MagiaSinEfecto = chkMapMagiaSinEfecto.Value
'MapInfo.Changed = 1
End Sub

Private Sub chkMapInviSinEfecto_LostFocus()
'*************************************************
'Author:
'Last modified:
'*************************************************
MapInfo.InviSinEfecto = chkMapInviSinEfecto.Value
'MapInfo.Changed = 1

End Sub

Private Sub chkMapResuSinEfecto_LostFocus()
'*************************************************
'Author:
'Last modified:
'*************************************************
MapInfo.ResuSinEfecto = chkMapResuSinEfecto.Value
'MapInfo.Changed = 1

End Sub

Private Sub chkMapNoEncriptarMP_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.NoEncriptarMP = chkMapNoEncriptarMP.Value
'MapInfo.Changed = 1
End Sub

Private Sub chkMapPK_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'MapInfo.PK = chkMapPK.value
'MapInfo.Changed = 1
End Sub

Private Sub chkTieneNpcInvocacion_Click()
'MapInfo.TieneNpcInvocacion = chkTieneNpcInvocacion.Value
'MapInfo.Changed = 1
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

Public Sub cmdIraZona_Click()
    Dim i As Integer
    Dim e As Integer

    If txtIrZona.Text > NumZonas Then
        Exit Sub
    Else
        i = txtIrZona.Text
    End If
        
    frmMapInfo.txtMapNombre.Text = Zonas(i).Nombre
    frmMapInfo.txtMapZona.ListIndex = Zonas(i).Terreno
    frmMapInfo.Check4.Value = IIf(Zonas(i).Segura, vbChecked, vbUnchecked)
    frmMapInfo.chkMapMagiaSinEfecto.Value = IIf(Zonas(i).MagiaSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.chkMapInviSinEfecto.Value = IIf(Zonas(i).InviSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.chkTieneNpcInvocacion.Value = IIf(Zonas(i).TieneNpcInvocacion, vbChecked, vbUnchecked)
    frmMapInfo.chkOcultarNombre.Value = IIf(Zonas(i).OcultarNombre, vbChecked, vbUnchecked)
    frmMapInfo.txtMapRestringir.Text = Zonas(i).Restringir
    frmMapInfo.txtMapNivelMaximo.Text = Zonas(i).NivelMaximo
    frmMapInfo.txtMapNivelMinimo.Text = Zonas(i).NivelMinimo
    frmMapInfo.txtMapMusica.Text = Zonas(i).Musica
        
    For e = 1 To 5
        tMusica(e - 1).Text = Zonas(i).Musicas(e)
    Next e

    For e = 1 To 5
        frmMapInfo.tMusica(e - 1).Text = Zonas(i).Musicas(e)
    Next e

    For e = 1 To 5
        frmMapInfo.tSonidos(e - 1).Text = Zonas(i).Sonido(e)
    Next e
    
    frmMapInfo.tZX1.Text = Zonas(i).X1
    frmMapInfo.tZY1.Text = Zonas(i).Y1
    frmMapInfo.tZX2.Text = Zonas(i).X2
    frmMapInfo.tZY2.Text = Zonas(i).Y2
    frmMapInfo.tNumZona.Caption = "Zona N°: " & i
    frmMapInfo.txtNiebla.Text = Zonas(i).niebla
    frmMapInfo.TxtR.Text = Zonas(i).NieblaR
    frmMapInfo.TxtG.Text = Zonas(i).NieblaG
    frmMapInfo.TxtB.Text = Zonas(i).NieblaB
    'ReyarB Info Mapa
    txtMapNombre.Text = Zonas(i).Nombre
    Check4.Value = IIf(Zonas(i).Segura, vbChecked, vbUnchecked)
    frmMapInfo.chkMapMagiaSinEfecto.Value = IIf(Zonas(i).MagiaSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.chkMapInviSinEfecto.Value = IIf(Zonas(i).InviSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.chkTieneNpcInvocacion.Value = IIf(Zonas(i).TieneNpcInvocacion, vbChecked, vbUnchecked)
    chkOcultarNombre.Value = IIf(Zonas(i).OcultarNombre, vbChecked, vbUnchecked)
    tZX1.Text = Zonas(i).X1
    tZY1.Text = Zonas(i).Y1
    tZX2.Text = Zonas(i).X2
    tZY2.Text = Zonas(i).Y2
    tNumZona.Caption = "Zona N°: " & i

End Sub

Private Sub cmdMusica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMusica.Show
End Sub

Private Sub Command1_Click()
    Dim ret As Integer

    Dim Num As Integer
    frmMapInfo.tMusica(0) = frmMapInfo.tMusica(0) - 1
    Num = frmMapInfo.tMusica(0)
    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
        ret = mciSendString("open " & """" & App.Path & "\..\Recursos\Midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If
End Sub

Private Sub Command2_Click()
    Dim ret As Integer

    Dim Num As Integer
    frmMapInfo.tMusica(0) = frmMapInfo.tMusica(0) + 1
    Num = frmMapInfo.tMusica(0)
    If IsPlaying Then
        ret = mciSendString("close mus", 0&, 0, 0)
        IsPlaying = False
        ret = mciSendString("open " & """" & App.Path & "\..\Recursos\Midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)
        ret = mciSendString("play mus", 0&, 0, 0)
        IsPlaying = True
    End If
End Sub

Private Sub Command24_Click()
AgregarZona = 0
SelZona = 1
Dim i As Integer
For i = 0 To 4
tMusica(i).Text = "0"
Next i
txtMapNombre.Text = ""
Check4.Value = vbUnchecked
CargarZonas
End Sub

Private Sub Command25_Click()
'Call mnuQuitarFunciones_Click
SelZona = 0
AgregarZona = 1
ZonaR.Left = 0
ZonaR.Top = 0
ZonaR.Right = 0
ZonaR.Bottom = 0
Me.Hide
End Sub

Private Sub Command26_Click()
Dim i As Integer
SelZona = txtIrZona.Text

If AgregarZona = 3 Then
    NumZonas = NumZonas + 1
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Config", "Cantidad", CStr(NumZonas))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Nombre", frmMapInfo.txtMapNombre.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Mapa", CStr(UserMap))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Terreno", CStr(frmMapInfo.txtMapZona.ListIndex))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "X1", frmMapInfo.tZX1.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Y1", frmMapInfo.tZY1.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "X2", frmMapInfo.tZX2.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Y2", frmMapInfo.tZY2.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Segura", IIf(frmMapInfo.Check4.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "MagiaSinEfecto", IIf(frmMapInfo.chkMapMagiaSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "InviSinEfecto", IIf(frmMapInfo.chkMapInviSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "TieneNpcInvocacion", IIf(frmMapInfo.chkTieneNpcInvocacion.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "ResuSinEfecto", IIf(frmMapInfo.chkMapResuSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Niebla", CStr(frmMapInfo.txtNiebla.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "NieblaR", CStr(frmMapInfo.TxtR.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "NieblaG", CStr(frmMapInfo.TxtG.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "NieblaB", CStr(frmMapInfo.TxtB.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Restringir", CStr(frmMapInfo.txtMapRestringir.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "NivelMaximo", CStr(frmMapInfo.txtMapNivelMaximo.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "NivelMinimo", CStr(frmMapInfo.txtMapNivelMinimo.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Musica", CStr(frmMapInfo.txtMapMusica.Text))
    
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "OcultarNombre", IIf(frmMapInfo.chkOcultarNombre.Value = vbChecked, 1, 0))
    For i = 0 To 4
        If Val(frmMapInfo.tMusica(i).Text) > 0 Then Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Musica" & (i + 1), frmMapInfo.tMusica(i).Text)
    Next i
    For i = 0 To 4
        If Val(frmMapInfo.tSonidos(i).Text) > 0 Then Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "Sonido" & (i + 1), frmMapInfo.tSonidos(i).Text)
    Next i
    
'    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & NumZonas, "************** NUEVA ZONA **********************", "")
    
    frmMapInfo.tNumZona.Caption = "Zona N°: " & NumZonas
    CargarZonas
ElseIf SelZona > 0 Then
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Nombre", frmMapInfo.txtMapNombre.Text)
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Segura", IIf(frmMapInfo.Check4.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "MagiaSinEfecto", IIf(frmMapInfo.chkMapMagiaSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "InviSinEfecto", IIf(frmMapInfo.chkMapInviSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "TieneNpcInvocacion", IIf(frmMapInfo.chkTieneNpcInvocacion.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "ResuSinEfecto", IIf(frmMapInfo.chkMapResuSinEfecto.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Terreno", CStr(frmMapInfo.txtMapZona.ListIndex))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "OcultarNombre", IIf(frmMapInfo.chkOcultarNombre.Value = vbChecked, 1, 0))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "X1", CStr(frmMapInfo.tZX1.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Y1", CStr(frmMapInfo.tZY1.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "X2", CStr(frmMapInfo.tZX2.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Y2", CStr(frmMapInfo.tZY2.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Niebla", CStr(frmMapInfo.txtNiebla.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NieblaR", CStr(frmMapInfo.TxtR.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NieblaG", CStr(frmMapInfo.TxtG.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NieblaB", CStr(frmMapInfo.TxtB.Text))
    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Restringir", CStr(frmMapInfo.txtMapRestringir.Text))
    
    If frmMapInfo.txtMapNivelMaximo.Text = "0" Or frmMapInfo.txtMapNivelMaximo.Text = "" Then
        Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NivelMaximo", "")
    Else
        Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NivelMaximo", CStr(frmMapInfo.txtMapNivelMaximo.Text))
    End If
    If frmMapInfo.txtMapNivelMinimo.Text = "0" Or frmMapInfo.txtMapNivelMinimo.Text = "" Then
        Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NivelMinimo", "")
    Else
        Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "NivelMinimo", CStr(frmMapInfo.txtMapNivelMinimo.Text))
    End If
    
    For i = 0 To 4
        If Val(frmMapInfo.tMusica(i).Text) > 0 Then
            Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Musica" & (i + 1), frmMapInfo.tMusica(i).Text)
        Else
            Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Musica" & (i + 1), "")
        End If
    Next
    
    For i = 0 To 4
        If Val(frmMapInfo.tSonidos(i).Text) > 0 Then
            Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Sonido" & (i + 1), frmMapInfo.tSonidos(i).Text)
        Else
            Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "Sonido" & (i + 1), "")
        End If
    Next i
    
'    Call WriteVar(App.Path & "\..\Recursos\Dat\zonas.dat", "Zona" & SelZona, "************** NUEVA ZONA **********************", "")

    frmMapInfo.tNumZona.Caption = "Zona N°: " & SelZona
    CargarZonas
End If
AgregarZona = 0
End Sub

Private Sub Command27_Click()
Dim i As Integer
Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Mapa", CStr(Areas(SelArea).Mapa))
Areas(SelArea).X1 = Val(frmMapInfo.tAX1.Text)
Areas(SelArea).Y1 = Val(frmMapInfo.tAY1.Text)
Areas(SelArea).X2 = Val(frmMapInfo.tAX2.Text)
Areas(SelArea).Y2 = Val(frmMapInfo.tAY2.Text)
Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "X1", CStr(frmMapInfo.tAX1.Text))
Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Y1", CStr(frmMapInfo.tAY1.Text))
Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "X2", CStr(frmMapInfo.tAX2.Text))
Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Y2", CStr(frmMapInfo.tAY2.Text))

Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Npcs", CStr(Areas(SelArea).NPCs))
For i = 1 To Areas(SelArea).NPCs
    Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Npc" & i, CStr(Areas(SelArea).NPC(i).NPCIndex))
    Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "Cant" & i, CStr(Areas(SelArea).NPC(i).cantidad))
Next i
'Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "", "")
'Call WriteVar(App.Path & "\..\Recursos\Dat\areas.dat", "Area" & SelArea, "----------  Nueva Area ---------------", "")
AgregarArea = 0
End Sub


Private Sub Command3_Click()
Dim Num As Integer
Dim ret As Integer

Num = frmMapInfo.tMusica(0)

If IsPlaying Then
   ret = mciSendString("close mus", 0&, 0, 0)
   IsPlaying = False
   ret = mciSendString("play mus", 0&, 0, 0)
   IsPlaying = True
End If
End Sub

Private Sub Command30_Click()
Dim i As Integer
Areas(SelArea).NPCs = Areas(SelArea).NPCs - 1
For i = lstNpc.ListIndex + 1 To Areas(SelArea).NPCs
    Areas(SelArea).NPC(i) = Areas(SelArea).NPC(i + 1)
Next i

lstNpc.RemoveItem (lstNpc.ListIndex)
End Sub

Private Sub Command31_Click()
Areas(SelArea).NPCs = Areas(SelArea).NPCs + 1
ReDim Preserve Areas(SelArea).NPC(1 To Areas(SelArea).NPCs)
Areas(SelArea).NPC(Areas(SelArea).NPCs).NPCIndex = Val(frmMapInfo.tNPC.Text)
Areas(SelArea).NPC(Areas(SelArea).NPCs).cantidad = Val(frmMapInfo.tCant.Text)

End Sub

Private Sub Command32_Click()
    Dim TempInt As Integer
    Dim y As Integer
    Dim X As Integer
    Dim tmpByte As Byte
    Dim FreeFileInf As Long
    DoEvents
    
    FreeFileInf = FreeFile
    'Open App.Path & "\Mapas\Mapa1.inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    frmMapInfo.List1.Clear
    
    'Load arrays
    For y = 1 To 100
        For X = 1 To 100
    
            
            '.inf file
            Get FreeFileInf, , tmpByte
            
            If tmpByte And 1 Then
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
            Dim i As Integer
            If tmpByte And 2 Then
                'Get and make NPC
                Get FreeFileInf, , TempInt
                
                For i = 0 To frmMapInfo.List1.ListCount - 1
                    If ReadField(1, frmMapInfo.List1.List(i), Asc("#")) = TempInt Then
                        frmMapInfo.List1.List(i) = ReadField(1, frmMapInfo.List1.List(i), Asc("#")) & "#" & ReadField(2, frmMapInfo.List1.List(i), Asc("#")) + 1 & "#" & ReadField(3, frmMapInfo.List1.List(i), Asc("#"))
                        i = -1
                        Exit For
                    End If
                Next i
                If i <> -1 Then
'                    frmMapInfo.List1.AddItem TempInt & "#1#" & FrmMain.NombreNPC(TempInt)
                End If
            End If
    
            If tmpByte And 4 Then
                'Get and make Object
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
    
        Next X
    Next y

    Close FreeFileInf

End Sub

Private Sub Command33_Click()
Dim Texto As String
If List1.ListIndex = -1 Then Exit Sub
Texto = List1.List(List1.ListIndex)

Areas(SelArea).NPCs = Areas(SelArea).NPCs + 1
ReDim Preserve Areas(SelArea).NPC(1 To Areas(SelArea).NPCs)

Areas(SelArea).NPC(Areas(SelArea).NPCs).NPCIndex = ReadField(1, Texto, Asc("#"))
Areas(SelArea).NPC(Areas(SelArea).NPCs).cantidad = ReadField(2, Texto, Asc("#"))
lstNpc.AddItem "(" & Areas(SelArea).NPC(Areas(SelArea).NPCs).cantidad & ") " & ReadField(3, Texto, Asc("#")) & "#" & Areas(SelArea).NPC(Areas(SelArea).NPCs).NPCIndex

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub txtMapNombre_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Name = txtMapNombre.Text
'frmMain.lblMapNombre.Caption = MapInfo.name
'MapInfo.Changed = 1
End Sub

Private Sub txtMapMusica_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Music = txtMapMusica.Text
'frmMain.lblMapMusica.Caption = MapInfo.Music
'MapInfo.Changed = 1
End Sub

Private Sub txtMapVersion_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
'MapInfo.MapVersion = txtMapVersion.text
''frmMain.lblMapVersion.Caption = MapInfo.MapVersion
'MapInfo.Changed = 1
End Sub

Private Sub txtMapNombre_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Name = txtMapNombre.Text
'frmMain.lblMapNombre.Caption = MapInfo.name
'MapInfo.Changed = 1
End Sub

Private Sub lstNpc_Click()
frmMapInfo.tNPC.Text = Areas(SelArea).NPC(lstNpc.ListIndex + 1).NPCIndex
tCant.Text = Areas(SelArea).NPC(lstNpc.ListIndex + 1).cantidad
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapRestringir_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Restringir = txtMapRestringir.Text
'MapInfo.Changed = 1
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapTerreno_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Terreno = frmMapInfo.txtMapZona.ListIndex
'MapInfo.Changed = 1
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapZona_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Zona = txtMapZona.Text
'MapInfo.Changed = 1
End Sub

