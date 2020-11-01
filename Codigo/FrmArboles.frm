VERSION 5.00
Begin VB.Form FrmArboles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insertar Arboles"
   ClientHeight    =   2640
   ClientLeft      =   15765
   ClientTop       =   1755
   ClientWidth     =   2055
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
   ScaleHeight     =   2640
   ScaleWidth      =   2055
   Begin VB.CheckBox Check1 
      Caption         =   "Bloquear"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "49"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "149"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "147"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "30"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Insertar 
      Caption         =   "Insertar arboles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Index n°4:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Index n°3:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Index n°2:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Index n°1:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "FrmArboles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      Dim lR As Long
         lR = SetTopMostWindow(FrmArboles.hwnd, True)
End Sub

Private Sub Insertar_Click()
Dim cantidad As Long
Dim bloquear As Byte
Dim objeto As Long
Dim X As Byte
Dim Y As Byte
Dim i As Long

Dim BuscarX As Byte
Dim BuscarY As Byte
Dim poner As Boolean

cantidad = Text1.Text
If cantidad <= 0 Then Exit Sub
bloquear = check1


Dim minx As Byte
Dim maxx As Byte
Dim miny As Byte
Dim maxy As Byte
maxy = 8
miny = 8
minx = 8
maxx = 8



If Text2 > 0 Then

objeto = Text2
For i = 1 To cantidad
            X = RandomNumber(14, 86)
            Y = RandomNumber(12, 89)
            If MapData(X, Y).Graphic(1).grhindex < 1505 Or MapData(X, Y).Graphic(1).grhindex > 1520 Then
            poner = True
            For BuscarX = X - minx To X + maxx
                For BuscarY = Y - miny To Y + maxy
                If MapData(BuscarX, BuscarY).OBJInfo.objindex <> 0 Then
                    If ObjData(MapData(BuscarX, BuscarY).OBJInfo.objindex).ObjType = 4 Then
                        poner = False
                    End If
                End If
                Next BuscarY
            Next BuscarX
            If poner Then
                    MapData(X, Y).Blocked = bloquear
                    InitGrh MapData(X, Y).ObjGrh, ObjData(objeto).grhindex
                    MapData(X, Y).OBJInfo.objindex = Text2
                    MapData(X, Y).OBJInfo.Amount = 1
                    End If

        
            If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
            Else
                i = i - 1
            End If
    Next i
End If
i = 0

If Text3 > 0 Then
objeto = Text3
For i = 1 To cantidad
            X = RandomNumber(14, 86)
            Y = RandomNumber(12, 89)
            If MapData(X, Y).Graphic(1).grhindex < 1505 Or MapData(X, Y).Graphic(1).grhindex > 1520 Then
            poner = True
            For BuscarX = X - minx To X + maxx
                For BuscarY = Y - miny To Y + maxy
                If MapData(BuscarX, BuscarY).OBJInfo.objindex <> 0 Then
                    If ObjData(MapData(BuscarX, BuscarY).OBJInfo.objindex).ObjType = 4 Then
                        poner = False
                    End If
                End If
                Next BuscarY
            Next BuscarX
            If poner Then
                    MapData(X, Y).Blocked = bloquear
                    InitGrh MapData(X, Y).ObjGrh, ObjData(objeto).grhindex
                    MapData(X, Y).OBJInfo.objindex = Text3
                    MapData(X, Y).OBJInfo.Amount = 1
                    End If

        
            If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
            Else
                i = i - 1
            End If
    Next i
End If
        
            
i = 0

If Text4 > 0 Then
objeto = Text4
For i = 1 To cantidad
            X = RandomNumber(14, 86)
            Y = RandomNumber(12, 89)
            If MapData(X, Y).Graphic(1).grhindex < 1505 Or MapData(X, Y).Graphic(1).grhindex > 1520 Then
            poner = True
            For BuscarX = X - minx To X + maxx
                For BuscarY = Y - miny To Y + maxy
                If MapData(BuscarX, BuscarY).OBJInfo.objindex <> 0 Then
                    If ObjData(MapData(BuscarX, BuscarY).OBJInfo.objindex).ObjType = 4 Then
                        poner = False
                    End If
                End If
                Next BuscarY
            Next BuscarX
            If poner Then
                    MapData(X, Y).Blocked = bloquear
                    InitGrh MapData(X, Y).ObjGrh, ObjData(objeto).grhindex
                    MapData(X, Y).OBJInfo.objindex = Text4
                    MapData(X, Y).OBJInfo.Amount = 1
                    End If

        
            If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
            Else
                i = i - 1
            End If
    Next i
End If
            
            
            
If Text5 > 0 Then
objeto = Text5
For i = 1 To cantidad
            X = RandomNumber(14, 86)
            Y = RandomNumber(12, 89)
            If MapData(X, Y).Graphic(1).grhindex < 1505 Or MapData(X, Y).Graphic(1).grhindex > 1520 Then
            poner = True
            For BuscarX = X - minx To X + maxx
                For BuscarY = Y - miny To Y + maxy
                If MapData(BuscarX, BuscarY).OBJInfo.objindex <> 0 Then
                    If ObjData(MapData(BuscarX, BuscarY).OBJInfo.objindex).ObjType = 4 Then
                        poner = False
                    End If
                End If
                Next BuscarY
            Next BuscarX
            If poner Then
                    MapData(X, Y).Blocked = bloquear
                    InitGrh MapData(X, Y).ObjGrh, ObjData(objeto).grhindex
                    MapData(X, Y).OBJInfo.objindex = Text5
                    MapData(X, Y).OBJInfo.Amount = 1
                    End If

        
            If MapData(X, Y).Graphic(3).grhindex <> MapData(X, Y).ObjGrh.grhindex Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
            Else
                i = i - 1
            End If
    Next i
End If
            
            
            
 Call AddtoRichTextBox(FrmMain.RichTextBox1, "Se agregaron " & cantidad & " arboles al mapa.", 255, 255, 255, False, True, False)
Call DibujarMiniMapa
DibujarMiniMapaParaMAPA
End Sub

