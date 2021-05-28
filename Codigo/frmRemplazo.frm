VERSION 5.00
Begin VB.Form frmRemplazo 
   Caption         =   "Remplazo de graficos"
   ClientHeight    =   2025
   ClientLeft      =   9360
   ClientTop       =   6195
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   2025
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   840
      Left            =   6960
      TabIndex        =   28
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Frame FraNpc 
      Caption         =   "Npc"
      Height          =   1695
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
      Begin VB.OptionButton OptTodoEl 
         Caption         =   "Todo el Mundo"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptEsteMapa 
         Caption         =   "Este mapa"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   22
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   9
         Left            =   720
         TabIndex        =   21
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   25
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Remplazo NPC"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Deshacer"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   27
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npc Nº"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame FraObjetos 
      Caption         =   "Objetos"
      Height          =   1695
      Index           =   1
      Left            =   4800
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   15
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptEsteMapa 
         Caption         =   "Este mapa"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptTodoEl 
         Caption         =   "Todo el Mundo"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Remplazo Obj"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Deshacer"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   18
         Top             =   840
         Width           =   90
      End
   End
   Begin VB.Frame FraObjetos 
      Caption         =   "Grafico"
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton OptTodoEl 
         Caption         =   "Mismo Grafico"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptEsteMapa 
         Caption         =   "Este Mapa"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   3
         Text            =   "1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   1
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Remplazo Grh"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBDeshacer 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Deshacer"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grafico"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblDeCapa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de Capa"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmRemplazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Remplazografico()
    
    On Error GoTo Remplazografico_Err
    

    Dim Y As Integer
    Dim X As Integer
    Dim c As Long
    Dim D As Long
    
    
'    c = txtCapaN.Text
'    D = txtCapaD.Text
    
'    For y = YMinMapSize To YMaxMapSize
'        For X = XMinMapSize To XMaxMapSize
    
            ' If MapData(X, y).OBJInfo.objindex > 0 Then
            '  If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 4 Then
            '   If MapData(X, y).Graphic(3).grhindex = MapData(X, y).ObjGrh.grhindex Then MapData(X, y).Graphic(3).grhindex = 0
            '   MapData(X, y).OBJInfo.objindex = 0
            '   MapData(X, y).OBJInfo.Amount = 0
            '   MapData(X, y).Blocked = 0
            ' End If
            '  End If
        
'            If MapData(X, y).Graphic(c).grhindex = txtGRH.Text Then
'                InitGrh MapData(X, y).Graphic(c), 0  '.grhindex = 0
'                InitGrh MapData(X, y).Graphic(D), TxtGrh2.Text
            
'                'InitGrh MapData(X, y).Graphic(2), 0
'                MapData(X, y).Graphic(2).grhindex = TxtGrh.Text
'                InitGrh MapData(X, y).Graphic(2), TxtGrh2.Text
            
'            End If
        
            '        If MapData(X, y).Graphic(3).grhindex = 12445 Then
            '            MapData(X, y).Graphic(3).grhindex = 0
            '            'InitGrh MapData(X, y).Graphic(2), 0
            '            MapData(X, y).Graphic(2).grhindex = 12445
            '            InitGrh MapData(X, y).Graphic(2), 12445
            '        End If
        
            ' Dim num As Long
        
            ' For num = 943 To 950
            '   If MapData(X, y).Graphic(3).grhindex = num Then
            ' MapData(X, y).Graphic(3).grhindex = 0
            'InitGrh MapData(X, y).Graphic(2), 0
            'MapData(X, y).Graphic(2).grhindex = num
            ' InitGrh MapData(X, y).Graphic(2), num
            ' End If
            ' Next num
        
'        Next X
'    Next y

    
    Exit Sub

Remplazografico_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Remplazografico", Erl)
    Resume Next
    
End Sub






Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub LvBDeshacer_Click(Index As Integer)
    Dim Y As Integer
    Dim X As Integer
    Dim O As Long
    Dim D As Long
    

    
    Select Case Index

        Case 0

        
        Case 1
        
        
        Case 2
            O = txt(7).Text
            D = txt(5).Text
            
            For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
    
                    'If MapData(X, y).OBJInfo.objindex = 0 Then
                        If MapData(X, Y).Graphic(O).grhindex = txt(3).Text Then
                            InitGrh MapData(X, Y).Graphic(O), 0
                            InitGrh MapData(X, Y).Graphic(D), txt(0).Text
                            MapInfo.Changed = 1
                        End If
                   'End If
                Next X
            Next Y

        Case 3
            O = txt(7).Text
            D = txt(5).Text
            
            For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
    
                    If MapData(X, Y).OBJInfo.objindex = 0 Then
                        If MapData(X, Y).Graphic(D).grhindex = txt(0).Text Then
                            InitGrh MapData(X, Y).Graphic(D), 0
                            InitGrh MapData(X, Y).Graphic(O), txt(3).Text
                            MapInfo.Changed = 1
                        End If
                   End If
                Next X
            Next Y
            
        End Select
End Sub

Private Sub OptEsteMapa_Click(Index As Integer)
    
    If OptEsteMapa(0).Value = True Then
'        txt(0).Visible = True
'        lblX(0).Visible = True
    End If

End Sub

Private Sub OptTodoEl_Click(Index As Integer)

    If OptTodoEl(0).Value = True Then
        txt(0).Text = txt(3).Text
'        txt(0).Visible = False
'        lblX(0).Visible = False
      Else
'        txt(0).Visible = True
'        lblX(0).Visible = True
    End If

End Sub

Private Sub txt_Change(Index As Integer)
     If OptTodoEl(0).Value = True Then txt(0).Text = txt(3).Text

End Sub
