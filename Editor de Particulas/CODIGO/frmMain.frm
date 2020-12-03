VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Particular - Revolucion-Ao"
   ClientHeight    =   11220
   ClientLeft      =   4905
   ClientTop       =   2385
   ClientWidth     =   15780
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   748
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1052
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Opciones"
      Height          =   2655
      Left            =   8280
      TabIndex        =   109
      Top             =   8400
      Width           =   3015
      Begin VB.OptionButton Option4 
         Caption         =   "Arena"
         Height          =   195
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   116
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Nieve"
         Height          =   195
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   115
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pasto"
         Height          =   195
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   114
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nada"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   113
         Top             =   1440
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar Body"
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Particulas Centradas"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Fondo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   112
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   5880
      TabIndex        =   106
      Top             =   8040
      Width           =   495
   End
   Begin VB.Frame frameGravity 
      BorderStyle     =   0  'None
      Caption         =   "Gravity Settings"
      Height          =   1095
      Left            =   480
      TabIndex        =   95
      Top             =   8760
      Width           =   1935
      Begin VB.CheckBox chkGravity 
         Caption         =   "Influencia de gravedad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtBounceStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   97
         Text            =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtGravStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   96
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bounce Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   705
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   465
         Width           =   1185
      End
   End
   Begin VB.Frame frameMovement 
      BorderStyle     =   0  'None
      Caption         =   "Movement Settings"
      Height          =   1935
      Left            =   480
      TabIndex        =   84
      Top             =   8790
      Width           =   1935
      Begin VB.CheckBox chkXMove 
         Caption         =   "X Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkYMove 
         Caption         =   "Y Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox move_y2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   88
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox move_y1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   87
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox move_x2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   86
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox move_x1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   85
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   1605
         Width           =   1035
      End
   End
   Begin VB.Frame frameSpinSettings 
      BorderStyle     =   0  'None
      Caption         =   "Spin Settings"
      Height          =   1095
      Left            =   480
      TabIndex        =   78
      Top             =   8790
      Width           =   1935
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   255
         Left            =   105
         TabIndex        =   81
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox spin_speedH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   80
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox spin_speedL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   79
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (H):"
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (L):"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Particle Duration"
      Height          =   855
      Left            =   480
      TabIndex        =   74
      Top             =   8880
      Width           =   1935
      Begin VB.TextBox life 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   76
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkNeverDies 
         Caption         =   "Never Dies"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Life:"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame frmSettings 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   480
      TabIndex        =   41
      Top             =   8775
      Width           =   6600
      Begin VB.TextBox txtrx 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   58
         Text            =   "0"
         Top             =   1395
         Width           =   495
      End
      Begin VB.TextBox txtPCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   57
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtX1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   56
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtX2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   55
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   54
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtY2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   53
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "0"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox vecx1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   51
         Text            =   "-10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox vecx2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   50
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox vecy1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   49
         Text            =   "-50"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox vecy2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   48
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox life1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   47
         Text            =   "10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox life2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   46
         Text            =   "50"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox fric 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   45
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   3930
         TabIndex        =   44
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkresize 
         Caption         =   "Resize"
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   1920
         Width           =   1245
      End
      Begin VB.TextBox txtry 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   42
         Text            =   "0"
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   72
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X2:"
         Height          =   195
         Left            =   1950
         TabIndex        =   71
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   70
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y2"
         Height          =   195
         Left            =   1950
         TabIndex        =   69
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (L):"
         Height          =   195
         Left            =   3915
         TabIndex        =   68
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (H):"
         Height          =   195
         Left            =   3915
         TabIndex        =   67
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fricción:"
         Height          =   195
         Left            =   4320
         TabIndex        =   66
         Top             =   885
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label lblPCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Particles:"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   525
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Y:"
         Height          =   195
         Left            =   1950
         TabIndex        =   60
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize X:"
         Height          =   195
         Left            =   1950
         TabIndex        =   59
         Top             =   1440
         Width           =   675
      End
   End
   Begin VB.Frame frameColorSettings 
      BorderStyle     =   0  'None
      Caption         =   "Color Tint Settings"
      Height          =   2175
      Left            =   375
      TabIndex        =   29
      Top             =   8730
      Width           =   3975
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         ScaleHeight     =   795
         ScaleWidth      =   2355
         TabIndex        =   34
         Top             =   240
         Width           =   2415
      End
      Begin VB.ListBox lstColorSets 
         Height          =   840
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.HScrollBar BScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   32
         Top             =   1200
         Width           =   3015
      End
      Begin VB.HScrollBar GScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   31
         Top             =   1500
         Width           =   3015
      End
      Begin VB.HScrollBar RScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   30
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   165
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Particle Speed"
      Height          =   855
      Left            =   435
      TabIndex        =   26
      Top             =   8865
      Width           =   1935
      Begin VB.TextBox speed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Text            =   "0.5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Render Delay:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame frmfade 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2235
      Left            =   360
      TabIndex        =   20
      Top             =   8760
      Width           =   7680
      Begin VB.TextBox txtfin 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Text            =   "0"
         Top             =   90
         Width           =   630
      End
      Begin VB.TextBox txtfout 
         Height          =   300
         Left            =   1320
         TabIndex        =   21
         Text            =   "0"
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label29 
         Caption         =   "Fade in time"
         Height          =   180
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label30 
         Caption         =   "Fade out time"
         Height          =   300
         Left            =   60
         TabIndex        =   24
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Note: The time a particle remains alive is set in the Duration Tab"
         Height          =   585
         Left            =   90
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.ListBox lstStreamType 
      Height          =   4545
      Left            =   11520
      TabIndex        =   18
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame frameGrhs 
      Caption         =   "Parametros de graficos"
      Height          =   4515
      Left            =   11400
      TabIndex        =   9
      Top             =   5280
      Width           =   4050
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   120
         ScaleHeight     =   131.097
         ScaleMode       =   0  'User
         ScaleWidth      =   240
         TabIndex        =   107
         Top             =   2400
         Width           =   3600
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Limpiar"
         Height          =   255
         Left            =   1590
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.ListBox lstGrhs 
         Height          =   1620
         Left            =   45
         TabIndex        =   13
         Top             =   450
         Width           =   1500
      End
      Begin VB.ListBox lstSelGrhs 
         Height          =   1620
         Left            =   2385
         TabIndex        =   12
         Top             =   450
         Width           =   1530
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   1590
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Quitar"
         Height          =   255
         Left            =   1590
         TabIndex        =   10
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de graficos"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   255
         Width           =   1170
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graficos de Particulas"
         Height          =   195
         Left            =   2370
         TabIndex        =   16
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label28 
         Caption         =   "Vista Previa"
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   2145
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opciones"
      Height          =   4815
      Left            =   13530
      TabIndex        =   3
      Top             =   480
      Width           =   2040
      Begin VB.CommandButton Command5 
         Caption         =   "Copiar como nueva"
         Height          =   375
         Left            =   120
         TabIndex        =   108
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cambiar Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox namelbl 
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdClearParticleGroups 
         Caption         =   "&Clear Particle Groups"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdEngineStats 
         Caption         =   "Toggle &Engine Stats"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdNewDoubleParticle 
         Caption         =   "New &Double Particle"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdNewParticle 
         Caption         =   "Crear nueva Particula"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Guardar Particulas"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label35 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   960
         TabIndex        =   105
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "FPS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   104
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vista Previa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   2
      Top             =   9960
      Width           =   2175
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   14280
      Top             =   720
   End
   Begin VB.Timer SpoofCheck 
      Interval        =   1000
      Left            =   14280
      Top             =   360
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   420
      Left            =   14040
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":4282
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7920
      Left            =   120
      ScaleHeight     =   528
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   1
      Top             =   120
      Width           =   11040
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2670
      Left            =   240
      TabIndex        =   101
      Top             =   8400
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Opciones de Particula"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravedad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimiento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spin "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Velocidad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Duración"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Desvanecimiento"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStreamType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de particulas"
      Height          =   195
      Left            =   11880
      TabIndex        =   19
      Top             =   360
      Width           =   1290
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6240
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   8190
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Dim PuedeMacrear As Boolean


Private Sub Check1_Click()
MostrarBody = Not MostrarBody
End Sub

Private Sub Check2_Click()
ParticulasCentradas = Not ParticulasCentradas
End Sub

Private Sub cmdAdd_Click()
Dim loopc As Long
If lstGrhs.ListIndex >= 0 Then lstSelGrhs.AddItem lstGrhs.List(lstGrhs.ListIndex)

StreamData(lstStreamType.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

ReDim StreamData(lstStreamType.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)

For loopc = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    StreamData(lstStreamType.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
Next loopc

End Sub

Private Sub cmdDelete_Click()
Dim loopc As Long

If lstSelGrhs.ListIndex >= 0 Then lstSelGrhs.RemoveItem lstSelGrhs.ListIndex

StreamData(lstStreamType.ListIndex + 1).NumGrhs = lstSelGrhs.ListCount

If StreamData(lstStreamType.ListIndex + 1).NumGrhs = 0 Then
    Erase StreamData(lstStreamType.ListIndex + 1).grh_list
Else
    ReDim StreamData(lstStreamType.ListIndex + 1).grh_list(1 To lstSelGrhs.ListCount)
End If

For loopc = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    StreamData(lstStreamType.ListIndex + 1).grh_list(loopc) = lstSelGrhs.List(loopc - 1)
Next loopc
End Sub

Private Sub cmdNewParticle_Click()
Call cmdNewStream_Click
End Sub

Private Sub Command1_Click()
engine.Engine_Meteo_Particle_Set (lstStreamType.ListIndex + 1)
        
End Sub

Private Sub Command2_Click()
Call cmdSaveAll_Click
End Sub


Private Sub Command3_Click()
Dim n As Integer
Dim namesito As String
namesito = InputBox("Por favor ingrese el nuevo nombre", "Cambiar Nombre")
If namesito = "" Then Exit Sub
StreamData(lstStreamType.ListIndex + 1).Name = namesito


End Sub



Private Sub Command5_Click()
Dim Name As String
Dim NewStreamNumber As Integer


'Get name for new stream
Name = InputBox("Please enter a Stream Name", "New Stream")
If Name = "" Then Exit Sub

'Set new stream #
NewStreamNumber = lstStreamType.ListCount + 1

'Add stream to combo box
lstStreamType.AddItem NewStreamNumber & " - " & Name

'Add 1 to TotalStreams
TotalStreams = TotalStreams + 1

'ReDim StreamData(1 To NewStreamNumber) As Stream

'Add stream data to StreamData array


StreamData(NewStreamNumber) = StreamData(lstStreamType.ListIndex + 1)


End Sub

Private Sub Form_Load()
Me.Caption = "Editor de Particulas por Ladder - Revolucion-Ao"
lstColorSets.AddItem "Bottom Left"
    lstColorSets.AddItem "Top Left"
    lstColorSets.AddItem "Bottom Right"
    lstColorSets.AddItem "Top Right"
    frmSettings.Visible = True
    frmfade.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
CurStreamFile = DirInits & "Particles.ini"
engine.Fill_Grh_List lstGrhs
End Sub

Private Sub Option1_Click()
Fondo = 1
End Sub

Private Sub Option2_Click()
Fondo = 0
End Sub

Private Sub Option3_Click()
Fondo = 2
End Sub

Private Sub Option4_Click()
Fondo = 3
End Sub

Private Sub speed_Change()
On Error Resume Next
DataChanged = True

'Arrange decimal separator
Dim temp As String
temp = General_Field_Read(1, speed.Text, 44)
If Not temp = "" Then
    speed.Text = temp & "." & Right(speed.Text, Len(speed.Text) - Len(temp) - 1)
    speed.SelStart = Len(speed.Text)
     speed.SelLength = 0
End If
StreamData(frmMain.lstStreamType.ListIndex + 1).speed = speed.Text

End Sub

Private Sub speed_GotFocus()

speed.SelStart = 0
speed.SelLength = Len(speed.Text)

End Sub
Private Sub BScroll_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).b = BScroll.value
txtB.Text = BScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub
Private Sub lstStreamType_Click()
Dim loopc As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.Text = StreamData(lstStreamType.ListIndex + 1).NumOfParticles
txtX1.Text = StreamData(lstStreamType.ListIndex + 1).x1
txtY1.Text = StreamData(lstStreamType.ListIndex + 1).y1
txtX2.Text = StreamData(lstStreamType.ListIndex + 1).x2
txtY2.Text = StreamData(lstStreamType.ListIndex + 1).y2
txtAngle.Text = StreamData(lstStreamType.ListIndex + 1).angle
vecx1.Text = StreamData(lstStreamType.ListIndex + 1).vecx1
vecx2.Text = StreamData(lstStreamType.ListIndex + 1).vecx2
vecy1.Text = StreamData(lstStreamType.ListIndex + 1).vecy1
vecy2.Text = StreamData(lstStreamType.ListIndex + 1).vecy2
life1.Text = StreamData(lstStreamType.ListIndex + 1).life1
life2.Text = StreamData(lstStreamType.ListIndex + 1).life2
speed.Text = StreamData(lstStreamType.ListIndex + 1).speed
fric.Text = StreamData(lstStreamType.ListIndex + 1).friction
chkSpin.value = StreamData(lstStreamType.ListIndex + 1).spin
spin_speedL.Text = StreamData(lstStreamType.ListIndex + 1).spin_speedL
spin_speedH.Text = StreamData(lstStreamType.ListIndex + 1).spin_speedH
txtGravStrength.Text = StreamData(lstStreamType.ListIndex + 1).grav_strength
txtBounceStrength.Text = StreamData(lstStreamType.ListIndex + 1).bounce_strength
chkAlphaBlend.value = StreamData(lstStreamType.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(lstStreamType.ListIndex + 1).gravity
txtrx.Text = StreamData(lstStreamType.ListIndex + 1).grh_resizex
txtry.Text = StreamData(lstStreamType.ListIndex + 1).grh_resizey
chkXMove.value = StreamData(lstStreamType.ListIndex + 1).XMove
chkYMove.value = StreamData(lstStreamType.ListIndex + 1).YMove
move_x1.Text = StreamData(lstStreamType.ListIndex + 1).move_x1
move_x2.Text = StreamData(lstStreamType.ListIndex + 1).move_x2
move_y1.Text = StreamData(lstStreamType.ListIndex + 1).move_y1
move_y2.Text = StreamData(lstStreamType.ListIndex + 1).move_y2

If StreamData(lstStreamType.ListIndex + 1).grh_resize = True Then
    chkresize = vbChecked
Else
    chkresize = vbUnchecked
End If

If StreamData(lstStreamType.ListIndex + 1).life_counter = -1 Then
    life.Enabled = False
    chkNeverDies.value = vbChecked
Else
    life.Enabled = True
    life.Text = StreamData(lstStreamType.ListIndex + 1).life_counter
    chkNeverDies.value = vbUnchecked
End If

speed.Text = StreamData(lstStreamType.ListIndex + 1).speed

lstSelGrhs.Clear

For loopc = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(lstStreamType.ListIndex + 1).grh_list(loopc)
Next loopc

DataChanged = DataTemp

 
If StreamData(lstStreamType.ListIndex + 1).NumGrhs = 0 Then Exit Sub
Call Command1_Click
End Sub

Private Sub lstStreamType_KeyUp(KeyCode As Integer, Shift As Integer)

Dim loopc As Long
Dim DataTemp As Boolean
DataTemp = DataChanged

'Set the values
txtPCount.Text = StreamData(lstStreamType.ListIndex + 1).NumOfParticles
txtX1.Text = StreamData(lstStreamType.ListIndex + 1).x1
txtY1.Text = StreamData(lstStreamType.ListIndex + 1).y1
txtX2.Text = StreamData(lstStreamType.ListIndex + 1).x2
txtY2.Text = StreamData(lstStreamType.ListIndex + 1).y2
txtAngle.Text = StreamData(lstStreamType.ListIndex + 1).angle
vecx1.Text = StreamData(lstStreamType.ListIndex + 1).vecx1
vecx2.Text = StreamData(lstStreamType.ListIndex + 1).vecx2
vecy1.Text = StreamData(lstStreamType.ListIndex + 1).vecy1
vecy2.Text = StreamData(lstStreamType.ListIndex + 1).vecy2
life1.Text = StreamData(lstStreamType.ListIndex + 1).life1
life2.Text = StreamData(lstStreamType.ListIndex + 1).life2
speed.Text = StreamData(lstStreamType.ListIndex + 1).speed
fric.Text = StreamData(lstStreamType.ListIndex + 1).friction
chkSpin.value = StreamData(lstStreamType.ListIndex + 1).spin
spin_speedL.Text = StreamData(lstStreamType.ListIndex + 1).spin_speedL
spin_speedH.Text = StreamData(lstStreamType.ListIndex + 1).spin_speedH
txtGravStrength.Text = StreamData(lstStreamType.ListIndex + 1).grav_strength
txtBounceStrength.Text = StreamData(lstStreamType.ListIndex + 1).bounce_strength
chkAlphaBlend.value = StreamData(lstStreamType.ListIndex + 1).AlphaBlend
chkGravity.value = StreamData(lstStreamType.ListIndex + 1).gravity
txtrx.Text = StreamData(lstStreamType.ListIndex + 1).grh_resizex
txtry.Text = StreamData(lstStreamType.ListIndex + 1).grh_resizey
chkXMove.value = StreamData(lstStreamType.ListIndex + 1).XMove
chkYMove.value = StreamData(lstStreamType.ListIndex + 1).YMove
move_x1.Text = StreamData(lstStreamType.ListIndex + 1).move_x1
move_x2.Text = StreamData(lstStreamType.ListIndex + 1).move_x2
move_y1.Text = StreamData(lstStreamType.ListIndex + 1).move_y1
move_y2.Text = StreamData(lstStreamType.ListIndex + 1).move_y2

If StreamData(lstStreamType.ListIndex + 1).grh_resize = True Then
   chkresize = vbChecked
 Else
    chkresize = vbUnchecked
  End If

If StreamData(lstStreamType.ListIndex + 1).life_counter = -1 Then
    life.Enabled = False
    chkNeverDies.value = vbChecked
Else
    life.Enabled = True
    life.Text = StreamData(lstStreamType.ListIndex + 1).life_counter
    chkNeverDies.value = vbUnchecked
End If

speed.Text = StreamData(lstStreamType.ListIndex + 1).speed

lstSelGrhs.Clear

For loopc = 1 To StreamData(lstStreamType.ListIndex + 1).NumGrhs
    lstSelGrhs.AddItem StreamData(lstStreamType.ListIndex + 1).grh_list(loopc)
Next loopc

DataChanged = DataTemp



Call Command1_Click

End Sub

Private Sub cmdClearParticleGroups_Click()

 Rem Particle_Engine.Particle_Group_Remove_All

End Sub

Private Sub cmdengineStats_Click()

Rem engine.Engine_Stats_Show_Toggle

End Sub

Private Sub cmdNewStream_Click()
Dim Name As String
Dim NewStreamNumber As Integer


'Get name for new stream
Name = InputBox("Please enter a Stream Name", "New Stream")
If Name = "" Then Exit Sub

'Set new stream #
NewStreamNumber = lstStreamType.ListCount + 1

'Add stream to combo box
lstStreamType.AddItem NewStreamNumber & " - " & Name

'Add 1 to TotalStreams
TotalStreams = TotalStreams + 1

'ReDim StreamData(1 To NewStreamNumber) As Stream

'Add stream data to StreamData array
StreamData(NewStreamNumber).Name = Name

StreamData(NewStreamNumber).NumOfParticles = 20
StreamData(NewStreamNumber).x1 = 0
StreamData(NewStreamNumber).y1 = 0
StreamData(NewStreamNumber).x2 = 0
StreamData(NewStreamNumber).y2 = 0
StreamData(NewStreamNumber).angle = 0
StreamData(NewStreamNumber).vecx1 = -20
StreamData(NewStreamNumber).vecx2 = 20
StreamData(NewStreamNumber).vecy1 = -20
StreamData(NewStreamNumber).vecy2 = 20
StreamData(NewStreamNumber).life1 = 10
StreamData(NewStreamNumber).life2 = 50
StreamData(NewStreamNumber).friction = 8
StreamData(NewStreamNumber).spin_speedL = 0.1
StreamData(NewStreamNumber).spin_speedH = 0.1
StreamData(NewStreamNumber).grav_strength = 2
StreamData(NewStreamNumber).bounce_strength = -5
StreamData(NewStreamNumber).life_counter = -1
StreamData(NewStreamNumber).NumGrhs = 0

StreamData(NewStreamNumber).AlphaBlend = 1
StreamData(NewStreamNumber).gravity = 0


        


'Select the new stream type in the combo box
Rem lstStreamType.ListIndex = NewStreamNumber - 1

End Sub

Private Sub cmdSaveAll_Click()


Call GuardarParticulasBinaria

Exit Sub

Dim loopc As Long
Dim StreamFile As String
Dim Bypass As Boolean
Dim retval

If General_File_Exists(CurStreamFile, vbNormal) = True Then
    retval = MsgBox("The file " & CurStreamFile & " already exists!" & vbCrLf & "Would you like to overwrite it?", vbYesNoCancel Or vbQuestion)
    If retval = vbNo Then
        Bypass = False
    ElseIf retval = vbCancel Then
        Exit Sub
    ElseIf retval = vbYes Then
        StreamFile = CurStreamFile
        Bypass = True
    End If
End If

If Bypass = False Then
    With ComDlg
        .Filter = "*.ini (Stream Data Files)|*.ini"
        .ShowSave
        StreamFile = .FileName
    End With
    
    If General_File_Exists(StreamFile, vbNormal) = True Then
        retval = MsgBox("The file " & StreamFile & " already exists!" & vbCrLf & "Would you like to overwrite it?", vbYesNo Or vbQuestion)
        If retval = vbNo Then
            Exit Sub
        End If
    End If
End If

Dim GrhListing As String
Dim i As Long

'Check for existing data file and kill it
If General_File_Exists(StreamFile, vbNormal) Then Kill StreamFile

Dim n
Dim Datos$

n = FreeFile
Open StreamFile For Binary Access Write As n
Put n, , "[INIT]" & vbCrLf & "Total=" & Val(TotalStreams) & vbCrLf & vbCrLf


Put n, , "[Graphics]" & vbCrLf


'Write particle data to Particles.ini
'General_Var_Write StreamFile, "INIT", "Total", Val(TotalStreams)

For loopc = 1 To TotalStreams
    Put n, , "[" & Val(loopc) & "]" & vbCrLf
    
    Put n, , "Name=" & StreamData(loopc).Name & vbCrLf
    Put n, , "NumOfParticles=" & Val(StreamData(loopc).NumOfParticles) & vbCrLf
    Put n, , "X1=" & Val(StreamData(loopc).x1) & vbCrLf
    Put n, , "Y1=" & Val(StreamData(loopc).y1) & vbCrLf
    Put n, , "X2=" & Val(StreamData(loopc).x2) & vbCrLf
    Put n, , "Y2=" & Val(StreamData(loopc).y2) & vbCrLf
    Put n, , "Angle=" & Val(StreamData(loopc).angle) & vbCrLf
    Put n, , "VecX1=" & Val(StreamData(loopc).vecx1) & vbCrLf
    Put n, , "Vecy1=" & Val(StreamData(loopc).vecy1) & vbCrLf
    Put n, , "VecX2=" & Val(StreamData(loopc).vecx2) & vbCrLf
    Put n, , "Vecy2=" & Val(StreamData(loopc).vecy2) & vbCrLf
    Put n, , "Life1=" & Val(StreamData(loopc).life1) & vbCrLf
    Put n, , "Life2=" & Val(StreamData(loopc).life2) & vbCrLf
    Put n, , "Friction=" & Val(StreamData(loopc).friction) & vbCrLf
    Put n, , "Spin=" & Val(StreamData(loopc).spin) & vbCrLf
    Put n, , "Spin_SpeedL=" & Val(StreamData(loopc).spin_speedL) & vbCrLf
    Put n, , "Spin_SpeedH=" & Val(StreamData(loopc).spin_speedH) & vbCrLf
    Put n, , "Grav_Strength=" & Val(StreamData(loopc).grav_strength) & vbCrLf
    Put n, , "Bounce_Strength=" & Val(StreamData(loopc).bounce_strength) & vbCrLf
    Put n, , "AlphaBlend=" & Val(StreamData(loopc).AlphaBlend) & vbCrLf
    Put n, , "Gravity=" & Val(StreamData(loopc).gravity) & vbCrLf
    Put n, , "XMove=" & Val(StreamData(loopc).XMove) & vbCrLf
    Put n, , "YMove=" & Val(StreamData(loopc).YMove) & vbCrLf
    Put n, , "move_x1=" & Val(StreamData(loopc).move_x1) & vbCrLf
    Put n, , "move_x2=" & Val(StreamData(loopc).move_x2) & vbCrLf
    Put n, , "move_y1=" & Val(StreamData(loopc).move_y1) & vbCrLf
    Put n, , "move_y2=" & Val(StreamData(loopc).move_y2) & vbCrLf
    Put n, , "life_counter=" & Val(StreamData(loopc).life_counter) & vbCrLf
    Put n, , "Speed=" & str(StreamData(loopc).speed) & vbCrLf
    Put n, , "resize=" & CInt(StreamData(loopc).grh_resize) & vbCrLf
    Put n, , "rx=" & StreamData(loopc).grh_resizex & vbCrLf
    Put n, , "ry=" & StreamData(loopc).grh_resizey & vbCrLf
    Put n, , "NumGrhs=" & Val(StreamData(loopc).NumGrhs) & vbCrLf
    
    GrhListing = vbNullString
    For i = 1 To StreamData(loopc).NumGrhs
        GrhListing = GrhListing & StreamData(loopc).grh_list(i) & ","
    Next i

    Put n, , "Grh_List=" & GrhListing & vbCrLf
    Put n, , "ColorSet1=" & StreamData(loopc).colortint(0).r & "," & StreamData(loopc).colortint(0).g & "," & StreamData(loopc).colortint(0).b & vbCrLf
    Put n, , "ColorSet2=" & StreamData(loopc).colortint(1).r & "," & StreamData(loopc).colortint(1).g & "," & StreamData(loopc).colortint(1).b & vbCrLf
    Put n, , "ColorSet3=" & StreamData(loopc).colortint(2).r & "," & StreamData(loopc).colortint(2).g & "," & StreamData(loopc).colortint(2).b & vbCrLf
    Put n, , "ColorSet4=" & StreamData(loopc).colortint(3).r & "," & StreamData(loopc).colortint(3).g & "," & StreamData(loopc).colortint(3).b & vbCrLf
    
Next loopc

Close #n


Call GuardarParticulasBinaria

'Report the results
If TotalStreams > 1 Then
    MsgBox TotalStreams & " particle stream types saved to: " & vbCrLf & StreamFile, vbInformation
Else
    MsgBox TotalStreams & " particle stream type saved to: " & vbCrLf & StreamFile, vbInformation
End If

'Set DataChanged variable to false
DataChanged = False
CurStreamFile = StreamFile

End Sub

Sub GuardarParticulasBinaria()
Dim loopc As Long
Dim StreamFile As String
Dim Bypass As Boolean
Dim retval

Dim GrhListing As String
Dim i As Long

StreamFile = App.Path & "\..\recursos\INIT\Particles.ind"

'Check for existing data file and kill it
If General_File_Exists(StreamFile, vbNormal) Then Kill StreamFile

Dim n
Dim Datos$

n = FreeFile
Open StreamFile For Binary Access Write As #n
Put #n, , TotalStreams


For loopc = 1 To TotalStreams
    Put #n, , StreamData(loopc)
Next loopc

Close #n

'Report the results
If TotalStreams > 1 Then
    MsgBox TotalStreams & " particle stream types saved to: " & vbCrLf & StreamFile, vbInformation
Else
    MsgBox TotalStreams & " particle stream type saved to: " & vbCrLf & StreamFile, vbInformation
End If

'Set DataChanged variable to false
DataChanged = False
CurStreamFile = StreamFile
End Sub

Private Sub chkNeverDies_Click()

DataChanged = True

If chkNeverDies.value = vbChecked Then
    life.Enabled = False
    StreamData(frmMain.lstStreamType.ListIndex + 1).life_counter = -1
Else
    life.Enabled = True
    StreamData(frmMain.lstStreamType.ListIndex + 1).life_counter = life.Text
End If
End Sub

Private Sub chkSpin_Click()

DataChanged = True
StreamData(frmMain.lstStreamType.ListIndex + 1).spin = chkSpin.value

If chkSpin.value = vbChecked Then
    spin_speedL.Enabled = True
    spin_speedH.Enabled = True
Else
    spin_speedL.Enabled = False
    spin_speedH.Enabled = False
End If

End Sub



Private Sub GScroll_Change()
On Error Resume Next
DataChanged = True


StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).g = GScroll.value
txtG.Text = GScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub life_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).life_counter = life.Text
End Sub

Private Sub life_GotFocus()

life.SelStart = 0
life.SelLength = Len(life.Text)

End Sub

Private Sub lstColorSets_Click()

Dim DataTemp As Boolean
DataTemp = DataChanged

RScroll.value = StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).r
GScroll.value = StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).g
BScroll.value = StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).b

DataChanged = DataTemp
If DataChanged = True Then

Else
End If

End Sub

Private Sub RScroll_Change()
On Error Resume Next
DataChanged = True


StreamData(frmMain.lstStreamType.ListIndex + 1).colortint(lstColorSets.ListIndex).r = RScroll.value
txtR.Text = RScroll.value

picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub SpoofCheck_Timer()
If engine.bRunning Then engine.Engine_ActFPS
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.index
Case 1:
    frmSettings.Visible = True
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 2:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = True
    frmfade.Visible = False
Case 3:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = True
    frameGravity.Visible = False
    frmfade.Visible = False
Case 4:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = True
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 5:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = True
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 6:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 7:
    frmSettings.Visible = False
    frameColorSettings.Visible = True
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = False
Case 8:
    frmSettings.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False
    frmfade.Visible = True
End Select
End Sub
Private Sub chkresize_Click()
If chkresize.value = vbChecked Then
    StreamData(frmMain.lstStreamType.ListIndex + 1).grh_resize = True
Else
   StreamData(frmMain.lstStreamType.ListIndex + 1).grh_resize = False
End If
End Sub
Private Sub txtrx_Change()
On Error Resume Next
StreamData(frmMain.lstStreamType.ListIndex + 1).grh_resizex = txtrx.Text
End Sub

Private Sub txtry_Change()
On Error Resume Next
StreamData(frmMain.lstStreamType.ListIndex + 1).grh_resizey = txtry.Text
End Sub

Private Sub vecx1_GotFocus()

vecx1.SelStart = 0
vecx1.SelLength = Len(vecx1.Text)

End Sub

Private Sub vecx1_Change()
On Error Resume Next
DataChanged = True
Rem
StreamData(frmMain.lstStreamType.ListIndex + 1).vecx1 = vecx1.Text
End Sub

Private Sub vecx2_GotFocus()

vecx2.SelStart = 0
vecx2.SelLength = Len(vecx2.Text)

End Sub

Private Sub vecx2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).vecx2 = vecx2.Text
End Sub

Private Sub vecy1_GotFocus()

vecy1.SelStart = 0
vecy1.SelLength = Len(vecy1.Text)

End Sub

Private Sub vecy1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).vecy1 = vecy1.Text
End Sub

Private Sub vecy2_GotFocus()

vecy2.SelStart = 0
vecy2.SelLength = Len(vecy2.Text)

End Sub

Private Sub vecy2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).vecy2 = vecy2.Text
End Sub

Private Sub life1_GotFocus()

life1.SelStart = 0
life1.SelLength = Len(life1.Text)

End Sub

Private Sub life1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).life1 = life1.Text
End Sub

Private Sub life2_GotFocus()

life2.SelStart = 0
life2.SelLength = Len(life2.Text)

End Sub

Private Sub life2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).life2 = life2.Text
End Sub

Private Sub fric_GotFocus()

fric.SelStart = 0
fric.SelLength = Len(fric.Text)

End Sub

Private Sub fric_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).friction = fric.Text
End Sub

Private Sub spin_speedL_GotFocus()

spin_speedL.SelStart = 0
spin_speedL.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedL_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).spin_speedL = spin_speedL.Text
End Sub

Private Sub spin_speedH_GotFocus()

spin_speedH.SelStart = 0
spin_speedH.SelLength = Len(spin_speedH.Text)

End Sub

Private Sub spin_speedH_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).spin_speedH = spin_speedH.Text
End Sub

Private Sub txtPCount_GotFocus()

txtPCount.SelStart = 0
txtPCount.SelLength = Len(txtPCount.Text)

End Sub

Private Sub txtPCount_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).NumOfParticles = txtPCount.Text
End Sub

Private Sub txtX1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).x1 = txtX1.Text
End Sub

Private Sub txtX1_GotFocus()

txtX1.SelStart = 0
txtX1.SelLength = Len(txtX1.Text)

End Sub

Private Sub txtY1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).y1 = txtY1.Text
End Sub

Private Sub txtY1_GotFocus()

txtY1.SelStart = 0
txtY1.SelLength = Len(txtY1.Text)

End Sub

Private Sub txtX2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).x2 = txtX2.Text
End Sub

Private Sub txtX2_GotFocus()

txtX2.SelStart = 0
txtX2.SelLength = Len(txtX2.Text)

End Sub

Private Sub txtY2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).y2 = txtY2.Text
End Sub

Private Sub txtY2_GotFocus()

txtY2.SelStart = 0
txtY2.SelLength = Len(txtY2.Text)

End Sub

Private Sub txtAngle_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).angle = txtAngle.Text
End Sub

Private Sub txtAngle_GotFocus()

txtAngle.SelStart = 0
txtAngle.SelLength = Len(txtAngle.Text)

End Sub

Private Sub txtGravStrength_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).grav_strength = txtGravStrength.Text
End Sub

Private Sub txtGravStrength_GotFocus()

txtGravStrength.SelStart = 0
txtGravStrength.SelLength = Len(txtGravStrength.Text)

End Sub

Private Sub txtBounceStrength_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).bounce_strength = txtBounceStrength.Text
End Sub

Private Sub txtBounceStrength_GotFocus()

txtBounceStrength.SelStart = 0
txtBounceStrength.SelLength = Len(txtBounceStrength.Text)

End Sub

Private Sub move_x1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).move_x1 = move_x1.Text
End Sub

Private Sub move_x1_GotFocus()

move_x1.SelStart = 0
move_x1.SelLength = Len(move_x1.Text)

End Sub

Private Sub move_x2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).move_x2 = move_x2.Text
End Sub

Private Sub move_x2_GotFocus()

move_x2.SelStart = 0
move_x2.SelLength = Len(move_x2.Text)

End Sub

Private Sub move_y1_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).move_y1 = move_y1.Text
End Sub

Private Sub move_y1_GotFocus()

move_y1.SelStart = 0
move_y1.SelLength = Len(move_y1.Text)

End Sub

Private Sub move_y2_Change()
On Error Resume Next
DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).move_y2 = move_y2.Text
End Sub

Private Sub move_y2_GotFocus()

move_y2.SelStart = 0
move_y2.SelLength = Len(move_y2.Text)

End Sub


Private Sub chkAlphaBlend_Click()

DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).AlphaBlend = chkAlphaBlend.value
End Sub

Private Sub chkGravity_Click()

DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).gravity = chkGravity.value

If chkGravity.value = vbChecked Then
    txtGravStrength.Enabled = True
    txtBounceStrength.Enabled = True
Else
    txtGravStrength.Enabled = False
    txtBounceStrength.Enabled = False
End If

End Sub

Private Sub chkXMove_Click()

DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).XMove = chkXMove.value

If chkXMove.value = vbChecked Then
    move_x1.Enabled = True
    move_x2.Enabled = True
Else
    move_x1.Enabled = False
    move_x2.Enabled = False
End If

End Sub

Private Sub chkYMove_Click()

DataChanged = True

StreamData(frmMain.lstStreamType.ListIndex + 1).YMove = chkYMove.value

If chkYMove.value = vbChecked Then
    move_y1.Enabled = True
    move_y2.Enabled = True
Else
    move_y1.Enabled = False
    move_y2.Enabled = False
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub


Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub





Private Sub lstGrhs_Click()
picPreview.Refresh
Call Grh_Render_To_Hdc(picPreview, (lstGrhs.List(lstGrhs.ListIndex)), 0, 0)
End Sub

Private Sub lstSelGrhs_Click()
picPreview.Refresh
Call Grh_Render_To_Hdc(picPreview, (lstSelGrhs.List(lstSelGrhs.ListIndex)), 0, 0)


End Sub











Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + (32 * 17) Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + (32 * 13) Then Exit Function
    
    InGameArea = True
End Function


