VERSION 5.00
Begin VB.Form FrmBloques 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bloques"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Verdana"
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
   Picture         =   "FrmBloques.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   2280
      ScaleHeight     =   195.097
      ScaleMode       =   0  'User
      ScaleWidth      =   192
      TabIndex        =   4
      Top             =   120
      Width           =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insertar en Ultimo click"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmBloques.frx":0342
      Left            =   120
      List            =   "FrmBloques.frx":0344
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arrastrar desde el picture a la ubicacion y tocar la barra espaciado para repetir"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "FrmBloques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grafico As Long

Private Sub Combo1_Click()
    
    On Error GoTo Combo1_Click_Err
    

    List1.Clear
    Call CargarTipo(Combo1.ListIndex + 1)

    
    Exit Sub

Combo1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmBloques.Combo1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()

    On Error Resume Next

    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
    tXX = UltimoClickX
    tYY = UltimoClickY
    desptile = 0

    For i = 1 To CInt(Val(TILESX(TIPOOK(List1.ListIndex + 1))))
        For j = 1 To CInt(Val(TILESY(TIPOOK(List1.ListIndex + 1))))
        
            aux = Val(Grh(List1.ListIndex + 1)) + desptile

            If tYY > 100 Then Exit Sub
            If tXX > 100 Then Exit Sub
            MapData(tXX, tYY).Graphic(CInt(Val(LAYER(TIPOOK(List1.ListIndex + 1))))).grhindex = aux
            InitGrh MapData(tXX, tYY).Graphic(CInt(Val(LAYER(TIPOOK(List1.ListIndex + 1))))), aux
            tXX = tXX + 1
            desptile = desptile + 1
        Next j

        tXX = UltimoClickX
        tYY = tYY + 1
    Next i

    tYY = y
    MapInfo.Changed = 1

End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call CargarBloq

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmBloques.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    
    HotKeysAllow = False
    DesdeBloq = False
    picture1.BackColor = vbBlack
    picture1.Refresh

    grafico = Grh(List1.ListIndex + 1)

    'For x = 1 To TILESX(TIPOOK(List1.ListIndex))
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0

    'Call Grh_Render_To_Hdc(FrmBloques.Picture1.hdc, grafico, 0, 0, False)
    Dim SR  As RECT, DR As RECT
    Dim aux As Long
    Dim Tem As Long

    Debug.Print grafico

    If grafico = 0 Then Exit Sub

    RenderX = Val(TILESX(TIPOOK(List1.ListIndex + 1)))
    RenderY = Val(TILESY(TIPOOK(List1.ListIndex + 1)))
    Dim X    As Integer, y As Integer, j As Integer, i As Integer
    Dim Cont As Integer
        
    For i = 1 To CInt(Val(TILESY(TIPOOK(List1.ListIndex + 1))))
        For j = 1 To CInt(Val(TILESX(TIPOOK(List1.ListIndex + 1))))
               
            Call Grh_Render_To_HdcPNG(FrmBloques.picture1, (grafico), j * 32 - 32, i * 32 - 32, False)

            If Cont < CInt(Val(TILESY(TIPOOK(List1.ListIndex + 1)))) * CInt(Val(TILESX(TIPOOK(List1.ListIndex + 1)))) Then Cont = Cont + 1
            grafico = grafico + 1
        Next
    Next

    'Next x
    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmBloques.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo Picture1_MouseUp_Err
    

    If List1.ListIndex = -1 Then Exit Sub

    DesdeBloq = True
    RenderGrh = Grh(List1.ListIndex + 1)
    RenderX = Val(TILESX(TIPOOK(List1.ListIndex + 1)))
    RenderY = Val(TILESY(TIPOOK(List1.ListIndex + 1)))
    RenderLayer = Val(LAYER(TIPOOK(List1.ListIndex + 1)))

    
    Exit Sub

Picture1_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmBloques.Picture1_MouseUp", Erl)
    Resume Next
    
End Sub

