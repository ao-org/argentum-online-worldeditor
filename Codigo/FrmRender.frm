VERSION 5.00
Begin VB.Form FrmRender 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form4"
   ClientHeight    =   13830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20070
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   922
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12435
      Left            =   240
      ScaleHeight     =   829
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1093
      TabIndex        =   2
      Top             =   600
      Width           =   16395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Renderizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Smallpic 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FrmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'*************************************************************
' Capturar la imagen de controles
       
'  1 - Colocar un picturebox llamado picture1, un Command1 y un Command2 _
   2 - Agragar algunos controles _
   3 - Indicar en la Sub " Capturar_Imagen " .. el control a capturar
'*************************************************************
      
' Declaraciones del Api
      
'*************************************************************
' Función BitBlt para copiar la imagen del control en un picturebox
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
      
' Recupera la imagen del área del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Call engine.MapCapture(False, False)
    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmRender.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

'*************************************************************
' Sub que copia la imagen del control en un picturebox
'*************************************************************
Public Sub Capturar_Imagen(Control As Control, Destino As Object)
          
    Dim hdc             As Long
    Dim Escala_Anterior As Integer
    Dim Ancho           As Long
    Dim Alto            As Long
          
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
          
    On Error Resume Next

    ' Si da error es por que el control está dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
          
    If Err.Number = 438 Then
        ' Si el control está en un Frame, convierte la escala
        Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
        Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
        ' Si no cambia la escala del  contenedor a pixeles
        Control.Container.ScaleMode = vbPixels
        Ancho = Control.Width
        Alto = Control.Height

    End If
          
    ' limpia el error
    On Error GoTo 0

    ' Captura el área de pantalla correspondiente al control
    hdc = GetWindowDC(Control.hWnd)
    
    ' Copia esa área al picturebox
    If ToWorldMap2 Then
        Call BitBlt(Destino.hdc, 0 - 50, 0 - 50, Ancho - 50, Alto - 50, hdc, 0, 0, vbSrcCopy) '
        Call BitBlt(Destino.hdc, 0, 0, Ancho, Alto, hdc, 0, 0, vbSrcCopy)
    Else
        Call BitBlt(Destino.hdc, 0, 0, 3000, 3000, hdc, 0, 0, vbSrcCopy)
        

    End If
    
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.image
    
    ' Borra la imagen ya que ahora usa el Picture
    Call Destino.Cls
          
    On Error Resume Next

    If Err.Number = 0 Then
        ' Si el control no está en un  Frame, restaura la escala del contenedor
        Control.Container.ScaleMode = Escala_Anterior

    End If
          
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Unload Me

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmRender.Command1_Click", Erl)
    Resume Next
    
End Sub

