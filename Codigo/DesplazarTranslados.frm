VERSION 5.00
Begin VB.Form DesplazarTranslados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desplazar Translados"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7140
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DestinoDerecha 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Text            =   "13"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox ActualXDerecha 
      Height          =   285
      Left            =   4560
      TabIndex        =   18
      Text            =   "92"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox DesplazadaXDerecha 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Text            =   "88"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox DestinoIzquierda 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Text            =   "87"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox DesplazadaXIzquierda 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Text            =   "12"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox ActualXIzquierda 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Text            =   "9"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox DestinoYInferior 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Text            =   "11"
      Top             =   3900
      Width           =   495
   End
   Begin VB.TextBox DestinoSuperior 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "90"
      Top             =   340
      Width           =   495
   End
   Begin VB.TextBox DesplazadaYInferior 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Text            =   "91"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox DesplazadaYSuperior 
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Text            =   "10"
      Top             =   920
      Width           =   495
   End
   Begin VB.TextBox ActualYInferior 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "94"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox ActualYSuperior 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "7"
      Top             =   920
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Actual X"
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Actual X"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Destino Y="
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Destino Y="
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   3135
      Left            =   1800
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Desplazar a Y="
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3405
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Desplazar a Y="
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Actual Y="
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   3405
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Actual Y="
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   1920
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "DesplazarTranslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
Dim X As Byte
Dim Y As Byte

For X = 13 To 87
    For Y = ActualYSuperior To ActualYSuperior
        If MapData(X, Y).TileExit.Map > 0 Then
            MapData(X, DesplazadaYSuperior).TileExit.Map = MapData(X, Y).TileExit.Map
            MapData(X, DesplazadaYSuperior).TileExit.X = MapData(X, Y).TileExit.X
            MapData(X, DesplazadaYSuperior).TileExit.Y = DestinoSuperior
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
        Next Y
Next X

For X = 13 To 87
    For Y = ActualYInferior To ActualYInferior
        If MapData(X, Y).TileExit.Map > 0 Then
            MapData(X, DesplazadaYInferior).TileExit.Map = MapData(X, Y).TileExit.Map
            MapData(X, DesplazadaYInferior).TileExit.X = MapData(X, Y).TileExit.X
            MapData(X, DesplazadaYInferior).TileExit.Y = DestinoYInferior
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
        Next Y
Next X

For X = ActualXIzquierda To ActualXIzquierda
    For Y = 11 To 90
        If MapData(X, Y).TileExit.Map > 0 Then
            MapData(DesplazadaXIzquierda, Y).TileExit.Map = MapData(X, Y).TileExit.Map
            MapData(DesplazadaXIzquierda, Y).TileExit.X = DestinoIzquierda
            MapData(DesplazadaXIzquierda, Y).TileExit.Y = MapData(X, Y).TileExit.Y
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
        Next Y
Next X


For X = ActualXDerecha To ActualXDerecha
    For Y = 11 To 90
        If MapData(X, Y).TileExit.Map > 0 Then
            MapData(DesplazadaXDerecha, Y).TileExit.Map = MapData(X, Y).TileExit.Map
            MapData(DesplazadaXDerecha, Y).TileExit.X = DestinoDerecha
            MapData(DesplazadaXDerecha, Y).TileExit.Y = MapData(X, Y).TileExit.Y
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
        Next Y
Next X



            
        
End Sub
