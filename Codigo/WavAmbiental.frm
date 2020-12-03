VERSION 5.00
Begin VB.Form WavAmbiental 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Wavs"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2445
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      ItemData        =   "WavAmbiental.frx":0000
      Left            =   120
      List            =   "WavAmbiental.frx":003D
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "WavAmbiental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    
    FrmMain.TxtWav.Text = List1.ListIndex + 500
    Unload Me

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "WavAmbiental.List1_Click", Erl)
    Resume Next
    
End Sub
