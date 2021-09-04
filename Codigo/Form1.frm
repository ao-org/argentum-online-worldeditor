VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de graficos"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListGraficosind 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   3810
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Dim lR As Long
    lR = SetTopMostWindow(Form1.hWnd, True)
         
    Dim i As Long

    For i = 1 To MaxGrh
        Rem    If GrhData(i).NumFrames > 0 Then
        Form1.ListGraficosind.AddItem i
        Rem   End If
    Next i

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form1.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub ListGraficosind_Click()
    
    On Error GoTo ListGraficosind_Click_Err
    
    picture1.Refresh
    Call Grh_Render_To_Hdc(Form1.picture1, str$(ListGraficosind.ListIndex And &HFFFF&) + 1, 0, 0, False)
    frmConfigSup.MOSAICO.Value = vbUnchecked
    frmConfigSup.mAncho.Text = "0"
    frmConfigSup.mLargo.Text = "0"
    HotKeysAllow = False

    
    Exit Sub

ListGraficosind_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form1.ListGraficosind_Click", Erl)
    Resume Next
    
End Sub

Private Sub ListGraficosind_DblClick()
    
    On Error GoTo ListGraficosind_DblClick_Err
    
    FrmMain.cGrh.Text = str$(ListGraficosind.ListIndex And &HFFFF&) + 1

    
    Exit Sub

ListGraficosind_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form1.ListGraficosind_DblClick", Erl)
    Resume Next
    
End Sub
