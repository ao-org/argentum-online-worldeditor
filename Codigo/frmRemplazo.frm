VERSION 5.00
Begin VB.Form frmRemplazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remplazo de graficos"
   ClientHeight    =   2535
   ClientLeft      =   9345
   ClientTop       =   6180
   ClientWidth     =   8895
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   Begin VB.Timer IterateMaps 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   2040
   End
   Begin VB.CheckBox AllMaps 
      Caption         =   "Cambiar en todos los mapas"
      Height          =   375
      Left            =   4800
      TabIndex        =   29
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Triggers"
      Height          =   855
      Left            =   4680
      TabIndex        =   22
      Top             =   1080
      Width           =   4095
      Begin VB.TextBox TriggerReplaceFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   24
         Text            =   "1"
         Top             =   330
         Width           =   495
      End
      Begin VB.TextBox TriggerReplaceTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "1"
         Top             =   330
         Width           =   495
      End
      Begin WorldEditor.lvButtons_H Reemplazar 
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Remplazar"
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
         Caption         =   "Trigger Nº"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame FraNpc 
      Caption         =   "Npc"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   4455
      Begin VB.TextBox NpcReplaceTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Text            =   "1"
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox NpcReplaceFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Text            =   "1"
         Top             =   330
         Width           =   735
      End
      Begin WorldEditor.lvButtons_H Reemplazar 
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Remplazar"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npc Nº"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame FraObjetos 
      Caption         =   "Objetos"
      Height          =   855
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   4095
      Begin VB.TextBox ObjReplaceFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Text            =   "1"
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox ObjReplaceTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Text            =   "1"
         Top             =   330
         Width           =   735
      End
      Begin WorldEditor.lvButtons_H Reemplazar 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Remplazar"
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
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame FraObjetos 
      Caption         =   "Grafico"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox SameGrh 
         Caption         =   "Mismo gráfico"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox GrhReplaceTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "1"
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox GrhReplaceFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "1"
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox HastaCapa 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "1"
         Top             =   330
         Width           =   495
      End
      Begin VB.TextBox DesdeCapa 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "1"
         Top             =   330
         Width           =   495
      End
      Begin WorldEditor.lvButtons_H Reemplazar 
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Remplazar"
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
         Left            =   1800
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
         Caption         =   "De capa"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmRemplazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Replacing As Integer

Private Sub DesdeCapa_Change()
    DesdeCapa.Text = Val(DesdeCapa.Text)
End Sub

Private Sub GrhReplaceFrom_Change()
    GrhReplaceFrom.Text = Val(GrhReplaceFrom.Text)
    
    If SameGrh.Value = vbChecked Then
        GrhReplaceTo.Text = GrhReplaceFrom.Text
    End If
End Sub

Private Sub GrhReplaceTo_Change()
    GrhReplaceTo.Text = Val(GrhReplaceTo.Text)
End Sub

Private Sub HastaCapa_Change()
    HastaCapa.Text = Val(HastaCapa.Text)
End Sub

Private Sub IterateMaps_Timer()
    Call ReplaceOnMap(Replacing)

    Call modMapIO.GuardarMapa(PATH_Save & FrmMain.MapPest(4).Caption)

    If Not FrmMain.MapPest(5).Visible Then
        IterateMaps.Enabled = False
        Exit Sub
    End If

    Call FrmMain.NextMap
End Sub

Private Sub NpcReplaceFrom_Change()
    NpcReplaceFrom.Text = Val(NpcReplaceFrom.Text)
End Sub

Private Sub NpcReplaceTo_Change()
    NpcReplaceTo.Text = Val(NpcReplaceTo.Text)
End Sub

Private Sub ObjReplaceFrom_Change()
    ObjReplaceFrom.Text = Val(ObjReplaceFrom.Text)
End Sub

Private Sub ObjReplaceTo_Change()
    ObjReplaceTo.Text = Val(ObjReplaceTo.Text)
End Sub

Private Sub Reemplazar_Click(Index As Integer)

    If AllMaps.Value = vbChecked Then
        If MsgBox("¿Está seguro que desea remplazar en todos los mapas?", vbYesNo, "Reemplazar") = vbYes Then
            Dim FileName As String
            FileName = PATH_Save & NameMap_Save & "1.csm"
        
            If FileExist(FileName, vbArchive) = False Then
                Unload Me
                MsgBox "Primero abre algún mapa de la carpeta a convertir.", vbOKOnly, "Error"
                Exit Sub
            End If
            
            Call AbrirMapa(FileName)

            Replacing = Index

            IterateMaps.Enabled = True
        End If
    Else
        Call ReplaceOnMap(Index)
    End If

End Sub

Private Sub SameGrh_Click()
    If SameGrh.Value = vbChecked Then
        GrhReplaceTo.Text = GrhReplaceFrom.Text
        GrhReplaceTo.Enabled = False
    Else
        GrhReplaceTo.Enabled = True
    End If
End Sub

Private Sub TriggerReplaceFrom_Change()
    TriggerReplaceFrom.Text = Val(TriggerReplaceFrom.Text)
End Sub

Private Sub TriggerReplaceTo_Change()
    TriggerReplaceTo.Text = Val(TriggerReplaceTo.Text)
End Sub

Private Sub ReplaceOnMap(ByVal Index As Integer)
    Dim ValueFrom As Long, ValueTo As Long
    Dim X As Integer, y As Integer

    Select Case Index

        Case 0
            ValueFrom = Val(GrhReplaceFrom.Text)
            ValueTo = Val(GrhReplaceTo.Text)
            
            Dim LayerFrom As Integer, LayerTo As Integer
            LayerFrom = Val(DesdeCapa.Text)
            LayerTo = Val(HastaCapa.Text)

            For y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize

                    With MapData(X, y)
                        If .Graphic(LayerFrom).grhindex = ValueFrom Then
                            .Graphic(LayerFrom).grhindex = 0
                            Call InitGrh(.Graphic(LayerTo), ValueTo)
                        End If
                    End With

                Next X
            Next y
            
        Case 1
            ValueFrom = Val(ObjReplaceFrom.Text)
            ValueTo = Val(ObjReplaceTo.Text)

            For y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize

                    With MapData(X, y)
                        If .OBJInfo.objindex = ValueFrom Then
                            If ValueTo <> 0 Then
                                Call InitGrh(.ObjGrh, ObjData(ValueTo).grhindex)
                            Else
                                .ObjGrh.grhindex = 0
                            End If
                            .OBJInfo.objindex = ValueTo
                        End If
                    End With

                Next X
            Next y
            
        Case 2
            ValueFrom = Val(NpcReplaceFrom.Text)
            ValueTo = Val(NpcReplaceTo.Text)

            For y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize

                    With MapData(X, y)
                        If .NPCIndex = ValueFrom Then
                            If ValueTo <> 0 Then
                                Call MakeChar(NextOpenChar(), NpcData(ValueTo).Body, NpcData(ValueTo).Head, NpcData(ValueTo).Heading, X, y)
                            Else
                                Call EraseChar(.CharIndex)
                            End If
                            .NPCIndex = ValueTo
                        End If
                    End With

                Next X
            Next y
            
        Case 3
            ValueFrom = Val(TriggerReplaceFrom.Text)
            ValueTo = Val(TriggerReplaceTo.Text)

            For y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize

                    With MapData(X, y)
                        If .Trigger = ValueFrom Then
                            .Trigger = ValueTo
                        End If
                    End With

                Next X
            Next y

    End Select

    MapInfo.Changed = 1
End Sub
