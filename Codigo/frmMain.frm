VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "TxtWav.Text = ""508-509"""
   ClientHeight    =   14610
   ClientLeft      =   2085
   ClientTop       =   750
   ClientWidth     =   24330
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":628A
   ScaleHeight     =   974
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1622
   Visible         =   0   'False
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12975
      Left            =   4800
      ScaleHeight     =   865
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1145
      TabIndex        =   192
      Top             =   1560
      Width           =   17175
   End
   Begin VB.CommandButton BloqAll 
      Caption         =   "X"
      Height          =   255
      Left            =   2040
      TabIndex        =   178
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "N"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   177
      Top             =   4680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Caption         =   "O"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   176
      Top             =   5040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "S"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   175
      Top             =   5400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkBloqueo 
      BackColor       =   &H80000000&
      Caption         =   "E"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   174
      Top             =   5040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir bloqueos"
      Height          =   375
      Left            =   20280
      TabIndex        =   173
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Opción Grh"
      Height          =   1095
      Left            =   16200
      TabIndex        =   168
      Top             =   0
      Width           =   5775
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   18
         Left            =   240
         TabIndex        =   169
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Grh Normal"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   19
         Left            =   240
         TabIndex        =   170
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Dia / Noche"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   21
         Left            =   2400
         TabIndex        =   172
         Top             =   600
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   22
         Left            =   2400
         TabIndex        =   193
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Caption         =   "Limpiar Luz,Particula,Trigger's"
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
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H80000000&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   22080
      TabIndex        =   144
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdDM 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmMain.frx":6ECC
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   720
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":71B3
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":74A2
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmMain.frx":7792
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   240
         Width           =   240
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   150
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":7A84
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   151
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":86D6
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   152
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":9328
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   153
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":9F7A
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   154
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "1"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   155
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "2"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   156
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "3"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   157
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "4"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   158
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Magic Mapas"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   159
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Ins Traslados"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   160
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Copy  manual"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   161
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Ir Map"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   162
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Ambientacion"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   163
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Copy Norte"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   855
         Index           =   14
         Left            =   240
         TabIndex        =   164
         Top             =   4920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Copy Oeste"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   855
         Index           =   15
         Left            =   1200
         TabIndex        =   165
         Top             =   4920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Copy  Este"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   166
         Top             =   5880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Copy  Sur"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   167
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Bloq"
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   20
         Left            =   240
         TabIndex        =   171
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Sup x Bloques"
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
   End
   Begin VB.CommandButton cmdCovertitMap 
      Caption         =   "Convertir Mapa"
      Height          =   375
      Left            =   18720
      TabIndex        =   143
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Techos transparentes"
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
      Left            =   16440
      TabIndex        =   141
      Top             =   1080
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   140
      Top             =   10440
      Width           =   4335
   End
   Begin VB.PictureBox MiniMapas2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000B&
      Height          =   1200
      Left            =   2760
      ScaleHeight     =   91.954
      ScaleMode       =   0  'User
      ScaleWidth      =   85.556
      TabIndex        =   136
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   125
      Top             =   120
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":ABCC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "Información del Mapa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   22080
      TabIndex        =   107
      Top             =   6480
      Width           =   2175
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "Sin mascotas/invocar"
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
         Index           =   6
         Left            =   120
         TabIndex        =   194
         Top             =   5880
         Width           =   1935
      End
      Begin VB.TextBox OutY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   191
         Text            =   "0"
         Top             =   7155
         Width           =   495
      End
      Begin VB.TextBox OutX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   190
         Text            =   "0"
         Top             =   7155
         Width           =   495
      End
      Begin VB.TextBox OutMap 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   189
         Text            =   "0"
         Top             =   7155
         Width           =   495
      End
      Begin VB.TextBox lvlMax 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   186
         Text            =   "0"
         Top             =   6600
         Width           =   495
      End
      Begin VB.TextBox lvlMin 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   185
         Text            =   "0"
         Top             =   6600
         Width           =   495
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "Solo para clanes"
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
         Index           =   5
         Left            =   120
         TabIndex        =   184
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "Sin Invi / ocultarse"
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
         Index           =   4
         Left            =   120
         TabIndex        =   183
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "NoCiud"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   182
         Top             =   5160
         Width           =   855
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "NoPKs"
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
         Index           =   2
         Left            =   1200
         TabIndex        =   181
         Top             =   4920
         Width           =   825
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "Newbie"
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
         TabIndex        =   180
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CheckBox MapFlags 
         BackColor       =   &H80000000&
         Caption         =   "Sin Magia"
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
         Index           =   1
         Left            =   120
         TabIndex        =   179
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox Seguro 
         BackColor       =   &H80000000&
         Caption         =   "Seguro"
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
         TabIndex        =   135
         Top             =   4680
         Width           =   975
      End
      Begin VB.CheckBox BackUp 
         BackColor       =   &H80000000&
         Caption         =   "Backup"
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
         Left            =   1200
         TabIndex        =   134
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtnamemapa 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   133
         Top             =   7680
         Width           =   1935
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
         ItemData        =   "frmMain.frx":AC43
         Left            =   840
         List            =   "frmMain.frx":AC50
         TabIndex        =   129
         Top             =   3960
         Width           =   1095
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
         ItemData        =   "frmMain.frx":AC6C
         Left            =   840
         List            =   "frmMain.frx":AC79
         TabIndex        =   128
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000000&
         Caption         =   "Luz base"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   1935
         Begin VB.TextBox LuzMapa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   124
            Top             =   580
            Width           =   1095
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   123
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H80000000&
            Caption         =   "Luz climatica"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            MaskColor       =   &H00404040&
            TabIndex        =   122
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000000&
         Caption         =   "Audio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   111
         Top             =   2040
         Width           =   1935
         Begin VB.CommandButton ProbarMidi 
            Caption         =   "Probar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   120
            Top             =   370
            Width           =   855
         End
         Begin VB.TextBox TxtMidi 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   119
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton ProbarMp3 
            Caption         =   "Probar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   117
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtMp3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   116
            Text            =   "0"
            Top             =   940
            Width           =   375
         End
         Begin VB.CommandButton ProbarAmbiental 
            Caption         =   "Elegir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1150
            TabIndex        =   114
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox TxtWav 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   113
            Text            =   "0"
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Musica Midi:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   190
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Musica MP3:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Wav Ambiental:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   112
            Top             =   1250
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000000&
         Caption         =   "Estado Climatico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   108
         Top             =   1320
         Width           =   1935
         Begin VB.CheckBox niebla 
            BackColor       =   &H80000000&
            Caption         =   "Niebla"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            MaskColor       =   &H00404040&
            TabIndex        =   126
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000000&
            Caption         =   "Lluvia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            MaskColor       =   &H00404040&
            TabIndex        =   110
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H80000000&
            Caption         =   "Nieve"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1080
            MaskColor       =   &H00404040&
            TabIndex        =   109
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Lbl_fuera_del_mapa 
         BackStyle       =   0  'Transparent
         Caption         =   "Afuera del mapa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Label Lbl_maplvl 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel mapa (min/max)     (0 ignorar)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   187
         Top             =   6150
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del mapa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   132
         Top             =   7500
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000000&
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
         Left            =   240
         TabIndex        =   131
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000000&
         Caption         =   "Terreno:"
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
         TabIndex        =   130
         Top             =   4320
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Complear"
      Height          =   375
      Left            =   22200
      TabIndex        =   106
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000B&
      Height          =   1500
      Left            =   120
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   101.01
      TabIndex        =   101
      Top             =   120
      Width           =   1500
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         Top             =   480
         Width           =   375
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   11040
      TabIndex        =   40
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":AC96
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   10080
      TabIndex        =   39
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":B25C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   9120
      TabIndex        =   38
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":B75D
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   8160
      TabIndex        =   37
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":BB11
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   7200
      TabIndex        =   36
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":BE92
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   6240
      TabIndex        =   35
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":F4F2
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   435
      Left            =   13920
      TabIndex        =   34
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   1440
      Top             =   2400
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3435
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmMain.frx":12A38
      Top             =   6360
      Width           =   4395
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   4
      Top             =   1800
      Width           =   4395
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   1320
         Top             =   120
      End
      Begin WorldEditor.lvButtons_H insertarParticula 
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Insertar Particula"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   240
         Top             =   3120
      End
      Begin VB.TextBox ColorLuz 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   98
         Text            =   "0"
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox LuzColor 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   97
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin WorldEditor.lvButtons_H quitarparticula 
         Height          =   375
         Left            =   2280
         TabIndex        =   95
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Quitar Particula"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   120
      End
      Begin VB.TextBox RangoLuz 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   92
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox numerodeparticula 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   91
         Text            =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   8
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   9
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   43
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   54
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   61
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   62
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   63
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   76
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":12A78
         Left            =   3360
         List            =   "frmMain.frx":12A7A
         TabIndex        =   73
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H insertarLuz 
         Height          =   375
         Left            =   240
         TabIndex        =   93
         Top             =   1800
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Luz"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H QuitarLuz 
         Height          =   375
         Left            =   240
         TabIndex        =   94
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Luz"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":12A7C
         Left            =   1080
         List            =   "frmMain.frx":12A8C
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":12A9C
         Left            =   840
         List            =   "frmMain.frx":12A9E
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   66
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":12AA0
         Left            =   840
         List            =   "frmMain.frx":12AA2
         TabIndex        =   51
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":12AA4
         Left            =   840
         List            =   "frmMain.frx":12AA6
         TabIndex        =   70
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   58
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":12AA8
         Left            =   3360
         List            =   "frmMain.frx":12AAA
         TabIndex        =   60
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":12AAC
         Left            =   4440
         List            =   "frmMain.frx":12AAE
         Sorted          =   -1  'True
         TabIndex        =   72
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   49
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   69
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   82
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":12AB0
         Left            =   3360
         List            =   "frmMain.frx":12AB2
         TabIndex        =   50
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   77
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   78
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   79
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   975
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1720
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   1800
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   84
         Top             =   2880
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":12AB4
         Left            =   120
         List            =   "frmMain.frx":12AB6
         Sorted          =   -1  'True
         TabIndex        =   64
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":12AB8
         Left            =   120
         List            =   "frmMain.frx":12ABA
         Sorted          =   -1  'True
         TabIndex        =   59
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":12ABC
         Left            =   120
         List            =   "frmMain.frx":12ABE
         Sorted          =   -1  'True
         TabIndex        =   48
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":12AC0
         Left            =   120
         List            =   "frmMain.frx":12AC2
         TabIndex        =   47
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   2040
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1085
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   375
         Left            =   2400
         TabIndex        =   46
         Top             =   3480
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H TiggerEspecial 
         Height          =   375
         Left            =   2400
         TabIndex        =   137
         Top             =   3840
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         Caption         =   "&Trigger Especial"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox ListaParticulas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3630
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Si el rango es mayor a 100 la luz se convierte en redonda."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   142
         Top             =   3120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   87
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   86
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   85
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   3660
      Left            =   120
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4395
      Begin VB.PictureBox PreviewGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3300
         Left            =   0
         ScaleHeight     =   220
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   293
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   4395
         Begin VB.Shape Cual 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            FillColor       =   &H0080FF80&
            Height          =   495
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   10080
      TabIndex        =   88
      Top             =   360
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   3
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":12AC4
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   12000
      TabIndex        =   89
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Particulas"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":12E78
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   12960
      TabIndex        =   90
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1826
      Caption         =   "Luces"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":13237
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
      Height          =   375
      Left            =   13920
      TabIndex        =   127
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "&Información del Mapa"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Label15"
      Height          =   255
      Left            =   5040
      TabIndex        =   139
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   138
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   255
      Left            =   22320
      TabIndex        =   105
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   22320
      TabIndex        =   104
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   22320
      TabIndex        =   103
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   22320
      TabIndex        =   102
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label POSX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   99
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0E0FF&
      Height          =   10965
      Left            =   4680
      Top             =   1440
      Width           =   11325
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   14340
      TabIndex        =   42
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   13575
      TabIndex        =   41
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5925
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6690
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7455
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8220
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9750
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   10515
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   11280
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   12045
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   12810
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAbrirMapaLong 
         Caption         =   "&Abrir Mapa Long"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu render_mapa 
         Caption         =   "Reenderizar"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPegarCasa 
         Caption         =   "&Pegar House"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu Npcalazarpormapa 
            Caption         =   "&Npc al azar por mapa"
         End
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu Todas_las_Particulas 
            Caption         =   "Todas las Particulas"
         End
         Begin VB.Menu Todas_las_luces 
            Caption         =   "Todas las luces"
         End
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerMarco 
         Caption         =   "...Marco"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuVerLuces 
         Caption         =   "...Luces"
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Particulas"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditarIndices 
         Caption         =   "Editar Indices.ini"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "Actualizar índices..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
   Begin VB.Menu ladder 
      Caption         =   "L&adder"
      Begin VB.Menu minimapSave 
         Caption         =   "Guardar MiniMap"
      End
      Begin VB.Menu SaveAllMiniMap 
         Caption         =   "Save All MiniMap"
      End
      Begin VB.Menu Stopminimap 
         Caption         =   "Stop Save All MiniMap"
      End
      Begin VB.Menu openminimap 
         Caption         =   "Abrir Mapa del mundo"
      End
      Begin VB.Menu borrarnegros 
         Caption         =   "Borrar bordes negros"
      End
      Begin VB.Menu abrirmapn 
         Caption         =   "Abrir mapa N°"
      End
      Begin VB.Menu vergraficoslistado 
         Caption         =   "Ver Graficos"
      End
      Begin VB.Menu Ambientacones 
         Caption         =   "Ambientaciones"
      End
      Begin VB.Menu copyborder 
         Caption         =   "Copiar Bordes Manual"
      End
      Begin VB.Menu copyauto 
         Caption         =   "Copiar Bordes Aut."
      End
      Begin VB.Menu desptranslados 
         Caption         =   "Desplazar Translados"
      End
   End
   Begin VB.Menu MiniMap_ 
      Caption         =   "MiniMap"
      Begin VB.Menu MiniMap_capa1 
         Caption         =   "Capa 1"
         Checked         =   -1  'True
      End
      Begin VB.Menu MiniMap_capa2 
         Caption         =   "Capa 2"
         Checked         =   -1  'True
      End
      Begin VB.Menu MiniMap_capa3 
         Caption         =   "Capa 3"
      End
      Begin VB.Menu MiniMap_capa4 
         Caption         =   "Capa 4"
      End
      Begin VB.Menu MiniMap_Npcs 
         Caption         =   "Npcs"
      End
      Begin VB.Menu MiniMap_objetos 
         Caption         =   "Objetos"
      End
      Begin VB.Menu MiniMap_Bloqueos 
         Caption         =   "Bloqueos"
      End
      Begin VB.Menu MiniMap_particulas 
         Caption         =   "Particulas"
      End
      Begin VB.Menu MiniMap_ndemapa 
         Caption         =   "N° de mapa"
      End
      Begin VB.Menu Dibujarmini 
         Caption         =   "Dibujar"
      End
   End
   Begin VB.Menu mapppear 
      Caption         =   "Mapear"
      Begin VB.Menu agua 
         Caption         =   "Agua"
      End
      Begin VB.Menu pasto 
         Caption         =   "Pasto"
      End
      Begin VB.Menu arena 
         Caption         =   "Arena"
      End
      Begin VB.Menu hielo 
         Caption         =   "Hielo"
      End
      Begin VB.Menu ins_ladder 
         Caption         =   "Insertar"
         Begin VB.Menu objalazar 
            Caption         =   "Objeto al Azar"
         End
         Begin VB.Menu arbolazar 
            Caption         =   "Arboles al azar"
         End
      End
      Begin VB.Menu blqq 
         Caption         =   "Bloquear"
         Begin VB.Menu blqspaciosvacios 
            Caption         =   "Espacios vacios"
         End
      End
      Begin VB.Menu BloquesOpen 
         Caption         =   "Bloques"
      End
   End
End
Attribute VB_Name = "FrmMain"
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

'MOTOR DX8 POR LADDER
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit
Public tX         As Byte
Public tY         As Byte
Public LastX      As Byte
Public LastY      As Byte
Public MouseX     As Long
Public MouseY     As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX     As Long
Private clicY     As Long

Private shlShell  As Shell32.Shell
Private shlFolder As Shell32.Folder

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 by GS
    '*************************************************
    
    On Error GoTo PonerAlAzar_Err
    
    Dim objindex As Long
    Dim NPCIndex As Long
    Dim X, y, i
    Dim Head    As Integer
    Dim Body    As Integer
    Dim Heading As Byte
    Dim Leer    As New clsIniReader
    i = n

    modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

    Do While i > 0
        X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
        y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
        Select Case T

            Case 0

                If MapData(X, y).OBJInfo.objindex = 0 Then
                    i = i - 1

                    If cInsertarBloqueo.Value = True Then
                        MapData(X, y).Blocked = 1
                    Else
                        MapData(X, y).Blocked = 0

                    End If

                    If cNumFunc(2).Text > 0 Then
                        objindex = cNumFunc(2).Text
                        InitGrh MapData(X, y).ObjGrh, ObjData(objindex).grhindex
                        MapData(X, y).OBJInfo.objindex = objindex
                        MapData(X, y).OBJInfo.Amount = Val(cCantFunc(2).Text)

                        Select Case ObjData(objindex).ObjType ' GS

                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh

                        End Select

                    End If

                End If

            Case 1

                If (MapData(X, y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If cNumFunc(T - 1).Text > 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y))
                        MapData(X, y).NPCIndex = NPCIndex

                    End If

                End If

            Case 2

                If (MapData(X, y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If cNumFunc(T - 1).Text >= 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y))
                        MapData(X, y).NPCIndex = NPCIndex

                    End If

                End If

        End Select

        DoEvents
    Loop

    
    Exit Sub

PonerAlAzar_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.PonerAlAzar", Erl)
    Resume Next
    
End Sub

Private Sub bloqqq_Click()
    
    On Error GoTo bloqqq_Click_Err
    
    Dim X As Byte
    Dim y As Byte
    Dim i As Long

    For X = 1 To 100
        For y = 1 To 100

            If MapData(X, y).Graphic(1).grhindex = 1 Then
                MapData(X, y).Blocked = 1

            End If

            ' If MapData(X, y).OBJInfo.objindex = 472 Then
            ' MapData(X, y).OBJInfo.objindex = 0
            ' MapData(X, y).Graphic(3).grhindex = 735
            '  MapData(x, y).Graphic(3).grhindex = 738
            
            ' End If
        Next y
    Next X

    
    Exit Sub

bloqqq_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.bloqqq_Click", Erl)
    Resume Next
    
End Sub

Private Sub abrirmapn_Click()
    
    On Error GoTo abrirmapn_Click_Err
    

    Dim Mapa As Integer
    Mapa = Val(InputBox("Ingrese el numero de mapa qe desea abrir."))

    If Mapa <> 0 Then
        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                modMapIO.GuardarMapa Dialog.FileName
            End If
        End If

        Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".csm"

        If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
        Call modMapIO.NuevoMapa
        DoEvents
        modMapIO.AbrirMapa Dialog.FileName
        EngineRun = True
        Exit Sub

    End If

    
    Exit Sub

abrirmapn_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.abrirmapn_Click", Erl)
    Resume Next
    
End Sub

Private Sub agua_Click()
    
    On Error GoTo agua_Click_Err
    
    cGrh.Text = DameGrhIndex(137)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

agua_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.agua_Click", Erl)
    Resume Next
    
End Sub

Private Sub Ambientacones_Click()
    
    On Error GoTo Ambientacones_Click_Err
    
    AmbientacionesForm.Show , FrmMain

    
    Exit Sub

Ambientacones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Ambientacones_Click", Erl)
    Resume Next
    
End Sub

Private Sub arbolazar_Click()
    
    On Error GoTo arbolazar_Click_Err
    
    FrmArboles.Show

    
    Exit Sub

arbolazar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.arbolazar_Click", Erl)
    Resume Next
    
End Sub

Private Sub arena_Click()
    
    On Error GoTo arena_Click_Err
    
    cGrh.Text = DameGrhIndex(245)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

arena_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.arena_Click", Erl)
    Resume Next
    
End Sub

Private Sub BloqAll_Click()
    
    On Error GoTo BloqAll_Click_Err
    
    Dim i As Integer
    
    If maskBloqueo = &HF Then
        For i = 0 To 3
            chkBloqueo(i).Value = vbUnchecked
        Next
        maskBloqueo = 0

    Else
        For i = 0 To 3
            chkBloqueo(i).Value = vbChecked
        Next
        maskBloqueo = &HF
    End If
    
    Exit Sub

BloqAll_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.BloqAll_Click", Erl)
    Resume Next
    
End Sub

Private Sub BloquesOpen_Click()
    
    On Error GoTo BloquesOpen_Click_Err
    
    Call CargarBloq

    
    Exit Sub

BloquesOpen_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.BloquesOpen_Click", Erl)
    Resume Next
    
End Sub

Private Sub blqspaciosvacios_Click()
    
    On Error GoTo blqspaciosvacios_Click_Err
    
    Dim X As Byte
    Dim y As Byte
    Dim i As Long

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).Graphic(1).grhindex = 0 Or MapData(X, y).Graphic(1).grhindex = 1 Then
                MapData(X, y).Blocked = 1

            End If

        Next X
    Next y

    Call DibujarMiniMapa

    
    Exit Sub

blqspaciosvacios_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.blqspaciosvacios_Click", Erl)
    Resume Next
    
End Sub

Private Sub borrarnegros_Click()
    
    On Error GoTo borrarnegros_Click_Err
    
    Dim X As Byte
    Dim y As Byte
    Dim i As Long

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).Graphic(2).grhindex = 7284 Or MapData(X, y).Graphic(2).grhindex = 7303 Or MapData(X, y).Graphic(2).grhindex = 7304 _
               Or MapData(X, y).Graphic(2).grhindex = 7308 Or MapData(X, y).Graphic(2).grhindex = 7310 Or MapData(X, y).Graphic(2).grhindex = 7315 Or MapData(X, y).Graphic(2).grhindex = 7316 _
               Or MapData(X, y).Graphic(2).grhindex = 7306 Or MapData(X, y).Graphic(2).grhindex = 7328 Or MapData(X, y).Graphic(2).grhindex = 7327 Or MapData(X, y).Graphic(2).grhindex = 7357 _
               Or MapData(X, y).Graphic(2).grhindex = 29382 Or MapData(X, y).Graphic(2).grhindex = 29384 Or MapData(X, y).Graphic(2).grhindex = 29383 Or MapData(X, y).Graphic(2).grhindex = 7290 Or MapData(X, y).Graphic(2).grhindex = 7291 Or MapData(X, y).Graphic(2).grhindex = 7358 Or MapData(X, y).Graphic(2).grhindex = 7376 _
               Or MapData(X, y).Graphic(2).grhindex = 7313 Or MapData(X, y).Graphic(2).grhindex = 7314 _
               Or MapData(X, y).Graphic(2).grhindex = 29379 Or MapData(X, y).Graphic(2).grhindex = 29649 Or MapData(X, y).Graphic(2).grhindex = 29393 Or MapData(X, y).Graphic(2).grhindex = 29401 Or MapData(X, y).Graphic(2).grhindex = 29403 Or MapData(X, y).Graphic(2).grhindex = 29366 Or MapData(X, y).Graphic(2).grhindex = 29388 Or MapData(X, y).Graphic(2).grhindex = 29390 Or MapData(X, y).Graphic(2).grhindex = 29392 Or MapData(X, y).Graphic(2).grhindex = 29395 Or MapData(X, y).Graphic(2).grhindex = 29396 Or MapData(X, y).Graphic(2).grhindex = 29399 Or MapData(X, y).Graphic(2).grhindex = 29398 Or MapData(X, y).Graphic(2).grhindex = 29397 Or MapData(X, y).Graphic(2).grhindex = 29407 Or MapData(X, y).Graphic(2).grhindex = 29408 Or MapData(X, y).Graphic(2).grhindex = 29409 Or MapData(X, y).Graphic(2).grhindex = 29410 Or MapData(X, y).Graphic(2).grhindex = 29373 Or MapData(X, y).Graphic(2).grhindex = 29372 _
               Or MapData(X, y).Graphic(2).grhindex = 7321 Or MapData(X, y).Graphic(2).grhindex = 7297 Or MapData(X, y).Graphic(2).grhindex = 7300 Or MapData(X, y).Graphic(2).grhindex = 7301 _
               Or MapData(X, y).Graphic(2).grhindex = 7302 Or MapData(X, y).Graphic(2).grhindex = 29619 Or MapData(X, y).Graphic(2).grhindex = 7311 _
               Or MapData(X, y).Graphic(2).grhindex = 29612 Or MapData(X, y).Graphic(2).grhindex = 29630 Or MapData(X, y).Graphic(2).grhindex = 29618 Or MapData(X, y).Graphic(2).grhindex = 29634 Or MapData(X, y).Graphic(2).grhindex = 29625 Or MapData(X, y).Graphic(2).grhindex = 29628 Or MapData(X, y).Graphic(2).grhindex = 29629 Or MapData(X, y).Graphic(2).grhindex = 29631 Or MapData(X, y).Graphic(2).grhindex = 29632 Or MapData(X, y).Graphic(2).grhindex = 29637 Or MapData(X, y).Graphic(2).grhindex = 29638 Or MapData(X, y).Graphic(2).grhindex = 29640 Or MapData(X, y).Graphic(2).grhindex = 29642 Or MapData(X, y).Graphic(2).grhindex = 29643 Or MapData(X, y).Graphic(2).grhindex = 29645 Or MapData(X, y).Graphic(2).grhindex = 29646 Or MapData(X, y).Graphic(2).grhindex = 29655 Or MapData(X, y).Graphic(2).grhindex = 29656 Or MapData(X, y).Graphic(2).grhindex = 29647 Or MapData(X, y).Graphic(2).grhindex = 29648 Or MapData(X, y).Graphic(2).grhindex = 29651 Or MapData(X, y).Graphic(2).grhindex = 29653 _
               Or MapData(X, y).Graphic(2).grhindex = 7325 Or MapData(X, y).Graphic(2).grhindex = 7326 Or MapData(X, y).Graphic(2).grhindex = 7354 _
               Or MapData(X, y).Graphic(2).grhindex = 7373 Or MapData(X, y).Graphic(2).grhindex = 7371 Or MapData(X, y).Graphic(2).grhindex = 7365 _
               Or MapData(X, y).Graphic(2).grhindex = 29597 Or MapData(X, y).Graphic(2).grhindex = 29595 Or MapData(X, y).Graphic(2).grhindex = 29596 _
               Or MapData(X, y).Graphic(2).grhindex = 29571 Or MapData(X, y).Graphic(2).grhindex = 29608 Or MapData(X, y).Graphic(2).grhindex = 29607 _
               Or MapData(X, y).Graphic(2).grhindex = 29588 Or MapData(X, y).Graphic(2).grhindex = 29590 Or MapData(X, y).Graphic(2).grhindex = 29583 _
               Or MapData(X, y).Graphic(2).grhindex = 29584 Or MapData(X, y).Graphic(2).grhindex = 29586 _
               Or MapData(X, y).Graphic(2).grhindex = 7369 Or MapData(X, y).Graphic(2).grhindex = 7367 Or MapData(X, y).Graphic(2).grhindex = 7352 _
               Or MapData(X, y).Graphic(2).grhindex = 7375 Or MapData(X, y).Graphic(2).grhindex = 7351 Or MapData(X, y).Graphic(2).grhindex = 7368 _
               Or MapData(X, y).Graphic(2).grhindex = 7332 Or MapData(X, y).Graphic(2).grhindex = 7339 Or MapData(X, y).Graphic(2).grhindex = 7366 _
               Or MapData(X, y).Graphic(2).grhindex = 7360 Or MapData(X, y).Graphic(2).grhindex = 7338 Or MapData(X, y).Graphic(2).grhindex = 7363 Or MapData(X, y).Graphic(2).grhindex = 29582 Or MapData(X, y).Graphic(2).grhindex = 29581 Or MapData(X, y).Graphic(2).grhindex = 29580 _
               Or MapData(X, y).Graphic(2).grhindex = 29593 Or MapData(X, y).Graphic(2).grhindex = 29594 Or MapData(X, y).Graphic(2).grhindex = 29570 _
               Or MapData(X, y).Graphic(2).grhindex = 29599 Or MapData(X, y).Graphic(2).grhindex = 29601 Or MapData(X, y).Graphic(2).grhindex = 29591 _
               Or MapData(X, y).Graphic(2).grhindex = 7349 Or MapData(X, y).Graphic(2).grhindex = 7348 Or MapData(X, y).Graphic(2).grhindex = 7345 _
               Or MapData(X, y).Graphic(2).grhindex = 29606 Or MapData(X, y).Graphic(2).grhindex = 29605 Or MapData(X, y).Graphic(2).grhindex = 29577 _
               Or MapData(X, y).Graphic(2).grhindex = 7350 Or MapData(X, y).Graphic(2).grhindex = 7362 Or MapData(X, y).Graphic(2).grhindex = 7338 _
               Or MapData(X, y).Graphic(2).grhindex = 7317 Or MapData(X, y).Graphic(2).grhindex = 7319 Or MapData(X, y).Graphic(2).grhindex = 8272 Or MapData(X, y).Graphic(2).grhindex = 8263 Then
                Rem 7357 Or 7358 Or 7375 Or 7376 Or 22590 Or 22588 Or 22594 Or 22595 Or 22582 Or 22583 Then
                MapData(X, y).Graphic(2).grhindex = 0

            End If
        
            If MapData(X, y).Graphic(1).grhindex = 0 Then
                MapData(X, y).Graphic(1).grhindex = 1

            End If

        Next X
    Next y

    Call DibujarMiniMapa
    Call mnuGuardarMapa_Click
       
    
    Exit Sub

borrarnegros_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.borrarnegros_Click", Erl)
    Resume Next
    
End Sub

Private Sub cAgregarFuncalAzar_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cAgregarFuncalAzar_Click_Err
    

    If IsNumeric(cCantFunc(Index).Text) = False Or cCantFunc(Index).Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub

    End If

    cAgregarFuncalAzar(Index).Enabled = False
    Call PonerAlAzar(CInt(cCantFunc(Index).Text), 1 + (IIf(Index = 2, -1, Index)))
    cAgregarFuncalAzar(Index).Enabled = True

    
    Exit Sub

cAgregarFuncalAzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cAgregarFuncalAzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cCantFunc_Change(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cCantFunc_Change_Err
    

    If Val(cCantFunc(Index)) < 1 Then
        cCantFunc(Index).Text = 1

    End If

    If Val(cCantFunc(Index)) > 10000 Then
        cCantFunc(Index).Text = 10000

    End If

    
    Exit Sub

cCantFunc_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCantFunc_Change", Erl)
    Resume Next
    
End Sub

Private Sub cCapas_Change()
    
    On Error GoTo cCapas_Change_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 31/05/06
    '*************************************************
    If Val(cCapas.Text) < 1 Then
        cCapas.Text = 1

    End If

    If Val(cCapas.Text) > 4 Then
        cCapas.Text = 4

    End If

    cCapas.Tag = vbNullString

    
    Exit Sub

cCapas_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCapas_Change", Erl)
    Resume Next
    
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
    
    On Error GoTo cCapas_KeyPress_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

    
    Exit Sub

cCapas_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cCapas_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_GotFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_GotFocus_Err
    
    HotKeysAllow = False

    
    Exit Sub

cFiltro_GotFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_GotFocus", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    
    On Error GoTo cFiltro_KeyPress_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If KeyAscii = 13 Then
        Call Filtrar(Index)

    End If

    
    Exit Sub

cFiltro_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_LostFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_LostFocus_Err
    
    HotKeysAllow = True

    
    Exit Sub

cFiltro_LostFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cFiltro_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)

        If FrmMain.cGrh.ListCount > 5 Then
            FrmMain.cGrh.RemoveItem 0

        End If

        FrmMain.cGrh.AddItem FrmMain.cGrh.Text

    End If

    Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub Check1_Click()
    
    On Error GoTo Check1_MouseUp_Err
    
    If LoadingMap Then Exit Sub

    If MapDat.lluvia = 0 Then

        MapDat.lluvia = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Lluvia en mapa activada.", 255, 255, 255, False, True, False)
    Else
        MapDat.lluvia = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Lluvia en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

Check1_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check1_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check10_Click()
    
    On Error GoTo Check10_Click_Err
    
    MiniMap_objetos = Not MiniMap_objetos

    
    Exit Sub

Check10_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check10_Click", Erl)
    Resume Next
    
End Sub

Private Sub Check11_Click()
    
    On Error GoTo Check11_Click_Err
    
    MiniMap_Npcs = Not MiniMap_Npcs

    
    Exit Sub

Check11_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check11_Click", Erl)
    Resume Next
    
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo Check2_MouseUp_Err
    

    If Nieba = 0 Then
        Nieba = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nieve en mapa activada.", 255, 255, 255, False, True, False)
    Else
        Nieba = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nieve en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

Check2_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check2_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check3_Click()
    
    On Error GoTo Check3_MouseUp_Err
    
    If LoadingMap Then Exit Sub

    If ColorAmb <> 0 Then
        Picture3.Enabled = False
        Picture3.BackColor = vbBlack
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "La luz del mapa sera segun la climatologia.", 255, 255, 255, False, True, False)
        ColorAmb = 0 'Luz Base por defecto5
        engine.Map_Base_Light_Set &HFFFFFFFF
        LuzMapa.Text = 0
        LightA.LightRenderAll

    Else
        Picture3.Enabled = True
        Picture3.BackColor = &HFFFFFF
        engine.Map_Base_Light_Set &HFFFFFFFF 'Luz de trabajo.
        ColorAmb = &HFFFFFFFF
        LuzMapa.Text = Hex(&HFFFFFF)
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Luz del mapa segun climatologia desactivada.", 255, 255, 255, False, True, False)
        LightA.LightRenderAll

    End If
    
    MapInfo.Changed = 1
    Exit Sub

Check3_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check3_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub DiaNoche()
    
    On Error GoTo DiaNoche_Err
    

    mnuVerParticulas_Click
    
    
        If ColorAmb <> 0 Then
        Picture3.Enabled = False
        Picture3.BackColor = vbBlack
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "La luz del mapa sera segun la climatologia.", 255, 255, 255, False, True, False)
        ColorAmb = 0 'Luz Base por defecto5
        engine.Map_Base_Light_Set &HFFFFFFFF
        LuzMapa.Text = 0
        LightA.LightRenderAll

    Else
        Picture3.Enabled = True
        Picture3.BackColor = &HFFFFFF
        engine.Map_Base_Light_Set &HFF8080AA 'Luz de trabajo.
        ColorAmb = &HFFFFFFFF
        LuzMapa.Text = Hex(&HFFFFFF)
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Luz del mapa segun climatologia desactivada.", 255, 255, 255, False, True, False)
        LightA.LightRenderAll

    End If
   
    
    Exit Sub

DiaNoche_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.DiaNoche", Erl)
    Resume Next
    
End Sub

Private Sub lvlMax_Change()
    If LoadingMap Then Exit Sub
    
    Dim Value As Double
    Value = Val(lvlMax.Text)
    
    If Value < 0 Then lvlMax.Text = "0"
    If Value > 255 Then lvlMax.Text = "255"

    MapDat.level = (MapDat.level And &HFFFF00FF) Or (Val(lvlMax.Text) * &H100)
    
    MapInfo.Changed = 1
End Sub

Private Sub lvlMin_Change()
    If LoadingMap Then Exit Sub
    
    Dim Value As Double
    Value = Val(lvlMin.Text)
    
    If Value < 0 Then lvlMin.Text = "0"
    If Value > 255 Then lvlMin.Text = "255"

    MapDat.level = (MapDat.level And &HFFFFFF00) Or Val(lvlMin.Text)
    
    MapInfo.Changed = 1
End Sub

Private Sub mnuVerMarco_Click()

    On Error GoTo mnuVerMarco_Click_Err
    
    VerMarco = (VerMarco = False)
    mnuVerMarco.Checked = VerMarco

    
    Exit Sub

mnuVerMarco_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerMarco_Click", Erl)
    Resume Next
End Sub

Private Sub OutMap_Change()
    If LoadingMap Then Exit Sub
    
    MapDat.salida = Val(OutMap.Text) & "-" & Val(OutX.Text) & "-" & Val(OutY.Text)
    
    MapInfo.Changed = 1
End Sub

Private Sub OutX_Change()
    If LoadingMap Then Exit Sub
    
    MapDat.salida = Val(OutMap.Text) & "-" & Val(OutX.Text) & "-" & Val(OutY.Text)
    
    MapInfo.Changed = 1
End Sub

Private Sub OutY_Change()
    If LoadingMap Then Exit Sub
    
    MapDat.salida = Val(OutMap.Text) & "-" & Val(OutX.Text) & "-" & Val(OutY.Text)
    
    MapInfo.Changed = 1
End Sub

Private Sub Seguro_Click()
    
    On Error GoTo Check4_MouseUp_Err
    
    If LoadingMap Then Exit Sub
    

    If MapDat.Seguro = 1 Then
        MapDat.Seguro = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Mapa inseguro", 255, 255, 255, False, True, False)
    Else
        MapDat.Seguro = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Mapa seguro.", 255, 255, 255, False, True, False)

    End If
    
    MapInfo.Changed = 1

    
    Exit Sub

Check4_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check4_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub BackUp_Click()
    
    On Error GoTo Check5_MouseUp_Err
    
    If LoadingMap Then Exit Sub
    

    If MapDat.backup_mode = 1 Then
        MapDat.backup_mode = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Backup de mapa desactivado.", 255, 255, 255, False, True, False)
    Else
        MapDat.backup_mode = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Backup de mapa activado.", 255, 255, 255, False, True, False)

    End If
    
    MapInfo.Changed = 1

    
    Exit Sub

Check5_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check5_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Check6_Click()
    
    On Error GoTo Check6_Click_Err
    
    AlphaTecho = Not AlphaTecho

    
    Exit Sub

Check6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Check7_Click()
    
    On Error GoTo Check7_Click_Err
    
    MiniMap_capa2 = Not MiniMap_capa2

    
    Exit Sub

Check7_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check7_Click", Erl)
    Resume Next
    
End Sub

Private Sub Check8_Click()
    
    On Error GoTo Check8_Click_Err
    
    MiniMap_capa3 = Not MiniMap_capa3

    
    Exit Sub

Check8_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check8_Click", Erl)
    Resume Next
    
End Sub

Private Sub Check9_Click()
    
    On Error GoTo Check9_Click_Err
    
    MiniMap_capa4 = Not MiniMap_capa4

    
    Exit Sub

Check9_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Check9_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command14_Click()
    
    On Error GoTo Command14_Click_Err
    
    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).particle_Index = 180 Then
                MapData(X, y).particle_Index = 0

            End If

        Next X
    Next y

    
    Exit Sub

Command14_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Command14_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command15_Click()
    
    On Error GoTo Command15_Click_Err
    
    Dim y As Long
    Dim X As Long

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            'If MapData(X, Y).NPCIndex = Text1 Then
            '     MapData(X, Y).NPCIndex = Text2
            'End If
        Next X
    Next y

    modMapIO.GuardarMapa Dialog.FileName

    
    Exit Sub

Command15_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Command15_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    SavePicture MiniMapas2.image, App.Path & "\recursos\minimapas\" & MapPest(4).Caption & ".png"
    Debug.Print Dialog.FileName
    modMapIO.GuardarMapa Dialog.FileName

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub chkBloqueo_Click(Index As Integer)
    
    On Error GoTo chkBloqueo_Click_Err
    
    maskBloqueo = maskBloqueo Xor 2 ^ Index

    
    Exit Sub

chkBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.chkBloqueo_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdCovertitMap_Click()
    
    On Error GoTo cmdCovertitMap_Click_Err
    
    FormatoIAO = True
    Dim i As Integer

    For i = 1 To 318
        FormatoIAO = False

        If FileExist(App.Path & "\Conversor\Mapa" & i & ".map", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call MapaV3_Cargar(App.Path & "\Conversor\Mapa" & i & ".map")
            FormatoIAO = True
            Call MapaV2_Guardar(App.Path & "\Conversor\Mapa" & i & ".csm")
            
            'Info.Caption = "Mapa" & i & " convertido correctamente!"
        End If

    Next i

    
    Exit Sub

cmdCovertitMap_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdCovertitMap_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdDM_Click(Index As Integer)
    
    On Error GoTo cmdDM_Click_Err
    
    frmConfigSup.DespMosaic.Value = vbChecked

    Select Case Index

        Case 0 'A
    
            frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) + 1

        Case 1 '<
            frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) + 1

        Case 2 '>
            frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) - 1

        Case 3 'V
            frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) - 1

        Case 4 '0
            frmConfigSup.DMAncho.Text = 0
            frmConfigSup.DMLargo.Text = 0

    End Select

    
    Exit Sub

cmdDM_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdDM_Click", Erl)
    Resume Next
    
End Sub

Private Sub Remplazograficos()
    
    On Error GoTo Remplazograficos_Err
    

    Dim y As Integer
    Dim X As Integer
    Dim c As Integer
    Dim D As Integer
    

    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            ' If MapData(X, y).OBJInfo.objindex > 0 Then
            '  If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 4 Then
            '   If MapData(X, y).Graphic(3).grhindex = MapData(X, y).ObjGrh.grhindex Then MapData(X, y).Graphic(3).grhindex = 0
            '   MapData(X, y).OBJInfo.objindex = 0
            '   MapData(X, y).OBJInfo.Amount = 0
            '   MapData(X, y).Blocked = 0
            ' End If
            '  End If
        
'            If MapData(X, y).Graphic(c).grhindex = txtGRH.Text Then
'                MapData(X, y).Graphic(D).grhindex = TxtGrh2.Text
            
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
        
        Next X
    Next y

    
    Exit Sub

Remplazograficos_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Remplazograficos", Erl)
    Resume Next
    
End Sub

Private Sub copyauto_Click()
    
    On Error GoTo copyauto_Click_Err
    
    Form3.Show , FrmMain

    
    Exit Sub

copyauto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.copyauto_Click", Erl)
    Resume Next
    
End Sub

Private Sub copyborder_Click()
    
    On Error GoTo copyborder_Click_Err
    
    Form2.Show , FrmMain

    
    Exit Sub

copyborder_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.copyborder_Click", Erl)
    Resume Next
    
End Sub

Private Sub desptranslados_Click()
    
    On Error GoTo desptranslados_Click_Err
    
    DesplazarTranslados.Show

    
    Exit Sub

desptranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.desptranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub Dibujarmini_Click()
    
    On Error GoTo Dibujarmini_Click_Err
    
    Call DibujarMiniMapa

    
    Exit Sub

Dibujarmini_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Dibujarmini_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    

'    If KeyCode = vbKeySpace Then
'        If FrmBloques.Visible = True Then
'            Call InsertarBloque
'
'        End If
'
'    End If

    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub hielo_Click()
    
    On Error GoTo hielo_Click_Err
    
    cGrh.Text = DameGrhIndex(621)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

hielo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.hielo_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label16_Click()
    
    On Error GoTo Label16_Click_Err
    
    Timer4.Enabled = True

    
    Exit Sub

Label16_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Label16_Click", Erl)
    Resume Next
    
End Sub

Private Sub LuzMapa_Change()
    'MapInfo.Changed = 1
End Sub

Private Sub LvBOpcion_Click(Index As Integer)
    
    On Error GoTo LvBOpcion_Click_Err
    

    Select Case Index

        Case 0
            cVerBloqueos.Value = (cVerBloqueos.Value = False)
            mnuVerBloqueos.Checked = cVerBloqueos.Value
            
                If mnuVerBloqueos.Checked = False Then
                    LvBOpcion(0).BackColor = &H80000000
                Else
                    LvBOpcion(0).BackColor = &H80FF80
                End If

        Case 1
            mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
            
                If mnuVerTranslados.Checked = False Then
                    LvBOpcion(1).BackColor = &H80000000
                Else
                    LvBOpcion(1).BackColor = &H80FF80
                End If

        Case 2
            mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
            
                If mnuVerObjetos.Checked = False Then
                    LvBOpcion(2).BackColor = &H80000000
                Else
                    LvBOpcion(2).BackColor = &H80FF80
                End If

        Case 3
            cVerTriggers.Value = (cVerTriggers.Value = False)
            mnuVerTriggers.Checked = cVerTriggers.Value
            
                If mnuVerTriggers.Checked = False Then
                    LvBOpcion(3).BackColor = &H80000000
                Else
                    LvBOpcion(3).BackColor = &H80FF80
                End If

        Case 4
            mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
            
                If mnuVerCapa1.Checked = False Then
                    LvBOpcion(4).BackColor = &H80000000
                Else
                    LvBOpcion(4).BackColor = &H80FF80
                End If

        Case 5
            mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
            
                If mnuVerCapa2.Checked = False Then
                    LvBOpcion(5).BackColor = &H80000000
                Else
                    LvBOpcion(5).BackColor = &H80FF80
                End If
        Case 6
            mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
            
                If mnuVerCapa3.Checked = False Then
                    LvBOpcion(6).BackColor = &H80000000
                Else
                    LvBOpcion(6).BackColor = &H80FF80
                End If
                
        Case 7
            mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
            
                If mnuVerCapa4.Checked = False Then
                    LvBOpcion(7).BackColor = &H80000000
                Else
                    LvBOpcion(7).BackColor = &H80FF80
                End If
        Case 8
            Call frmOptimizar.cOptimizar_Click
            'Norte
            Form2.Command1_Click
            Form2.Command5_Click
            ' copio el de arriba al oeste
            Form2.Command2_Click
            Form2.Command7_Click
            ' vuelvo
            Form2.Command3_Click
            Form2.Command8_Click
            'copio al sur
            Form2.Command4_Click
            Form2.Command6_Click
            'Oeste
            Form2.Command2_Click
            Form2.Command7_Click
            'copio sur y vuelvo
            Form2.Command4_Click
            Form2.Command6_Click
            Form2.Command1_Click
            Form2.Command5_Click
            
            Form2.Command3_Click
            Form2.Command8_Click
            'Este
            Form2.Command3_Click
            Form2.Command8_Click
            ' copio y vuelvo al norte
            Form2.Command1_Click
            Form2.Command5_Click
            Form2.Command4_Click
            Form2.Command6_Click
                       
            Form2.Command2_Click
            Form2.Command7_Click
            'Sur
            Form2.Command4_Click
            Form2.Command6_Click
            'copio este y vuelvo
            Form2.Command3_Click
            Form2.Command8_Click
            Form2.Command2_Click
            Form2.Command7_Click
            
            Form2.Command1_Click
            Form2.Command5_Click
            
        Case 9
            Call frmOptimizar.cOptimizar_Click
            frmUnionAdyacente.Show
    
        Case 10
            Form2.Show , FrmMain

        Case 11
            abrirmapn_Click

        Case 12
            AmbientacionesForm.Show , FrmMain
            Call SelectPanel_Click(0)
            modPaneles.VerFuncion 0, True
            cSeleccionarSuperficie.Enabled = True
            
        Case 13 'Norte
            Form2.Command1_Click
            Form2.Command5_Click
            Form2.Command4_Click
            Form2.Command6_Click

        Case 14 'Oeste
            Form2.Command2_Click
            Form2.Command7_Click
            Form2.Command3_Click
            Form2.Command8_Click

        Case 15 'Este
            Form2.Command3_Click
            Form2.Command8_Click
            Form2.Command2_Click
            Form2.Command7_Click

        Case 16 'Sur
            Form2.Command4_Click
            Form2.Command6_Click
            Form2.Command1_Click
            Form2.Command5_Click

        Case 17
            Call modEdicion.Bloquear_Bordes
            Call frmOptimizar.cOptimizar_Click

        Case 18
            mnuAutoCompletarSuperficies_Click

        Case 19
            Call DiaNoche
            cVerBloqueos.Value = False
            cVerTriggers.Value = False
            mnuVerParticulas.Checked = True
            
        Case 20
            Call InsertarBloque

        Case 21
            Call frmRemplazo.Show
        Case 22
            Call Todas_las_Particulas_Click
            Call Todas_las_luces_Click
            Call mnuQuitarTriggers_Click
            Call DiaNoche
    End Select

    
    Exit Sub

LvBOpcion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.LvBOpcion_Click", Erl)
    Resume Next
    
End Sub

Private Sub MapFlags_Click(Index As Integer)

    If Not LoadingMap Then

        Dim Flag As Byte
        Flag = 2 ^ Index
    
        MapDat.restrict_mode = Val(MapDat.restrict_mode) Xor Flag
    
        MapInfo.Changed = 1
        
    End If

End Sub

Private Sub MiniMap_Bloqueos_Click()
    
    On Error GoTo MiniMap_Bloqueos_Click_Err
    
    MiniMap_Bloqueos.Checked = (MiniMap_Bloqueos.Checked = False)
    MMiniMap_Bloqueos = Not MMiniMap_Bloqueos
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_Bloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Bloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa1_Click()
    
    On Error GoTo MiniMap_capa1_Click_Err
    
    MiniMap_capa1.Checked = (MiniMap_capa1.Checked = False)
    MMiniMap_capa1 = Not MMiniMap_capa1
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_capa1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa2_Click()
    
    On Error GoTo MiniMap_capa2_Click_Err
    
    MiniMap_capa2.Checked = (MiniMap_capa2.Checked = False)
    MMiniMap_capa2 = Not MMiniMap_capa2
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_capa2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa3_Click()
    
    On Error GoTo MiniMap_capa3_Click_Err
    
    MiniMap_capa3.Checked = (MiniMap_capa3.Checked = False)
    MMiniMap_capa3 = Not MMiniMap_capa3
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_capa3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa4_Click()
    
    On Error GoTo MiniMap_capa4_Click_Err
    
    MiniMap_capa4.Checked = (MiniMap_capa4.Checked = False)
    MMiniMap_capa4 = Not MMiniMap_capa4
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_capa4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_ndemapa_Click()
    
    On Error GoTo MiniMap_ndemapa_Click_Err
    
    MiniMap_ndemapa.Checked = (MiniMap_ndemapa.Checked = False)
    MMiniMap_Nombre = Not MMiniMap_Nombre
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_ndemapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_ndemapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_Npcs_Click()
    
    On Error GoTo MiniMap_Npcs_Click_Err
    
    MiniMap_Npcs.Checked = (MiniMap_Npcs.Checked = False)
    MMiniMap_Npcs = Not MMiniMap_Npcs
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_Npcs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Npcs_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_objetos_Click()
    
    On Error GoTo MiniMap_objetos_Click_Err
    
    MiniMap_objetos.Checked = (MiniMap_objetos.Checked = False)
    MMiniMap_objetos = Not MMiniMap_objetos
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_objetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_objetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_particulas_Click()
    
    On Error GoTo MiniMap_particulas_Click_Err
    
    MiniMap_particulas.Checked = (MiniMap_particulas.Checked = False)
    MMiniMap_particulas = Not MMiniMap_particulas
    Call DibujarMiniMapa

    
    Exit Sub

MiniMap_particulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_particulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAbrirMapaLong_Click()
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    FormatoIAO = False

    DeseaGuardarMapa Dialog.FileName

    ObtenerNombreArchivo False

    If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If
    
    Call modMapIO.NuevoMapa

    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuActualizarIndices_Click()
    
    On Error GoTo mnuActualizarIndices_Click_Err
    
    frmActualizarIndices.Show , Me

    
    Exit Sub

mnuActualizarIndices_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuActualizarIndices_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuEditarIndices_Click()
Shell "C:\WINDOWS\System32\notepad.exe " & App.Path & "\..\Recursos\init\indices.ini", vbNormalFocus


End Sub

Private Sub mnuPegarCasa_Click()
    Dim Mapa As Integer
    Dim MapaOld As String
    
    MapaOld = Label16.Caption
    
    Mapa = 169

    If Mapa <> 0 Then

        Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".csm"

        If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
        Call modMapIO.NuevoMapa
        DoEvents
        modMapIO.AbrirMapa Dialog.FileName
        EngineRun = True

    End If

Call CopiarSeleccionCasa

Dialog.FileName = PATH_Save & NameMap_Save & MapaOld & ".csm"
modMapIO.AbrirMapa Dialog.FileName

Call modEdicion.Deshacer_Add("Pegar Selección")
Call PegarSeleccionCasa
End Sub

Private Sub niebla_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo niebla_MouseUp_Err
    

    If nieblaV = 0 Then
        nieblaV = 1
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Niebla en mapa activada.", 255, 255, 255, False, True, False)
    Else
        nieblaV = 0
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Niebla en mapa desactivada.", 255, 255, 255, False, True, False)

    End If

    MapInfo.Changed = 1

    
    Exit Sub

niebla_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.niebla_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarFunc_Click(Index As Integer)
    
    On Error GoTo cInsertarFunc_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarFunc(Index).Value = True Then
        cQuitarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False

        If Index <> 2 Then cCantFunc(Index).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
    Else
        cQuitarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True

        If Index <> 2 Then cCantFunc(Index).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)

    End If

    
    Exit Sub

cInsertarFunc_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarFunc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarTrans_Click()
    
    On Error GoTo cInsertarTrans_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    If cInsertarTrans.Value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

    
    Exit Sub

cInsertarTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarTrigger_Click()
    
    On Error GoTo cInsertarTrigger_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarTrigger.Value = True Then
        cQuitarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        cQuitarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

    
    Exit Sub

cInsertarTrigger_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarTrigger_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdInformacionDelMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdInformacionDelMapa_Click_Err
    
    frmMapInfo.Show
    frmMapInfo.Visible = True

    
    Exit Sub

cmdInformacionDelMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdInformacionDelMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdQuitarFunciones_Click_Err
    
    Call mnuQuitarFunciones_Click

    
    Exit Sub

cmdQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cmdQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()
    'Timer1.Enabled = True

    On Error Resume Next

    Dim Folder As String

    If shlShell Is Nothing Then
        Set shlShell = New Shell32.Shell

    End If

    Set shlFolder = shlShell.BrowseForFolder(Me.hWnd, "Seleccione la carpeta de los mapas a convertir", 1)
    
    If shlFolder Is Nothing Then Exit Sub

    FormatoIAO = True
    
    Dim Mapa As Long

    For Mapa = 0 To shlFolder.Items.Count - 1
        Call modMapIO.NuevoMapa
        Call Load_Map_Data_CSM_Fast_ConBloqueosViejos(shlFolder.Self.Path & "\" & shlFolder.Items.Item(Mapa))
        Call Save_Map_Data(App.Path & "\Mapas Convertidos\" & shlFolder.Items.Item(Mapa))
    Next

    Set shlFolder = Nothing

End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    

    Label1.Caption = MapData(90, 50).TileExit.Map ' & " Derecha" 'Derecha
    Label2.Caption = MapData(11, 50).TileExit.Map ' & " Izquierda" 'Izquierda
    Label3.Caption = MapData(50, 10).TileExit.Map '& " arriba" 'arriba
    Label4.Caption = MapData(50, 91).TileExit.Map ' & " Abajo" 'Abajo

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub cUnionManual_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cUnionManual_Click_Err
    
    cInsertarTrans.Value = (cUnionManual.Value = True)
    Call cInsertarTrans_Click

    
    Exit Sub

cUnionManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cUnionManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub cverBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverBloqueos_Click_Err
    
    mnuVerBloqueos.Checked = cVerBloqueos.Value

    
    Exit Sub

cverBloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cverBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub cverTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverTriggers_Click_Err
    
    mnuVerTriggers.Checked = cVerTriggers.Value

    
    Exit Sub

cverTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cverTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_KeyPress(Index As Integer, KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_KeyPress_Err
    

    If KeyAscii = 13 Then
        Dim Cont As String
        Cont = FrmMain.cNumFunc(Index).Text
        Call cNumFunc_LostFocus(Index)

        If Cont <> FrmMain.cNumFunc(Index).Text Then Exit Sub
        If FrmMain.cNumFunc(Index).ListCount > 5 Then
            FrmMain.cNumFunc(Index).RemoveItem 0

        End If

        FrmMain.cNumFunc(Index).AddItem FrmMain.cNumFunc(Index).Text
        Exit Sub
    ElseIf KeyAscii = 8 Then
    
    ElseIf IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
        Exit Sub

    End If

    
    Exit Sub

cNumFunc_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_KeyUp_Err
    

    If cNumFunc(Index).Text = vbNullString Then
        FrmMain.cNumFunc(Index).Text = IIf(Index = 1, 500, 1)

    End If

    
    Exit Sub

cNumFunc_KeyUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub cNumFunc_LostFocus(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cNumFunc_LostFocus_Err
    

    If Index = 0 Then
        If FrmMain.cNumFunc(Index).Text > 499 Or FrmMain.cNumFunc(Index).Text < 1 Then
            FrmMain.cNumFunc(Index).Text = 1

        End If

    ElseIf Index = 1 Then

        If FrmMain.cNumFunc(Index).Text < 500 Or FrmMain.cNumFunc(Index).Text > 32000 Then
            FrmMain.cNumFunc(Index).Text = 500

        End If

    ElseIf Index = 2 Then

        If FrmMain.cNumFunc(Index).Text < 1 Or FrmMain.cNumFunc(Index).Text > 32000 Then
            FrmMain.cNumFunc(Index).Text = 1

        End If

    End If

    
    Exit Sub

cNumFunc_LostFocus_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cNumFunc_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo cInsertarBloqueo_Click_Err
    
    cInsertarBloqueo.Tag = vbNullString

    If cInsertarBloqueo.Value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

    
    Exit Sub

cInsertarBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cInsertarBloqueo_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cQuitarBloqueo_Click_Err
    
    cInsertarBloqueo.Tag = vbNullString

    If cQuitarBloqueo.Value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

    
    Exit Sub

cQuitarBloqueo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarBloqueo_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarEnEstaCapa_Click()
    
    On Error GoTo cQuitarEnEstaCapa_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnEstaCapa.Value = True Then
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cQuitarEnEstaCapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarEnEstaCapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
    
    On Error GoTo cQuitarEnTodasLasCapas_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnTodasLasCapas.Value = True Then
        cCapas.Enabled = False
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cCapas.Enabled = True
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cQuitarEnTodasLasCapas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarEnTodasLasCapas_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarFunc_Click(Index As Integer)
    
    On Error GoTo cQuitarFunc_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarFunc(Index).Value = True Then
        cInsertarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False
        cCantFunc(Index).Enabled = False
        cNumFunc(Index).Enabled = False
        cFiltro((Index) + 1).Enabled = False
        lListado((Index) + 1).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
    Else
        cInsertarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True
        cCantFunc(Index).Enabled = True
        cNumFunc(Index).Enabled = True
        cFiltro((Index) + 1).Enabled = True
        lListado((Index) + 1).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)

    End If

    
    Exit Sub

cQuitarFunc_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarFunc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarTrans_Click()
    
    On Error GoTo cQuitarTrans_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrans.Value = True Then
        cInsertarTransOBJ.Enabled = False
        cInsertarTrans.Enabled = False
        cUnionManual.Enabled = False
        cUnionAuto.Enabled = False
        tTMapa.Enabled = False
        tTX.Enabled = False
        tTY.Enabled = False
        mnuInsertarTransladosAdyasentes.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        tTMapa.Enabled = True
        tTX.Enabled = True
        tTY.Enabled = True
        cUnionAuto.Enabled = True
        cUnionManual.Enabled = True
        cInsertarTrans.Enabled = True
        cInsertarTransOBJ.Enabled = True
        mnuInsertarTransladosAdyasentes.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

    
    Exit Sub

cQuitarTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarTrigger_Click()
    
    On Error GoTo cQuitarTrigger_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrigger.Value = True Then
        lListado(4).Enabled = False
        cInsertarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        lListado(4).Enabled = True
        cInsertarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

    
    Exit Sub

cQuitarTrigger_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cQuitarTrigger_Click", Erl)
    Resume Next
    
End Sub

Public Sub cSeleccionarSuperficie_Click()
    
    On Error GoTo cSeleccionarSuperficie_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cSeleccionarSuperficie.Value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cSeleccionarSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cSeleccionarSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub cUnionAuto_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    'Call MapPest_Click(4)
    
    On Error GoTo cUnionAuto_Click_Err
    
    frmUnionAdyacente.Show

    
    Exit Sub

cUnionAuto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.cUnionAuto_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    'Me.SetFocus

End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Caption = "WorldEditor DX8 por Ladder"

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

    'If Seleccionando Then CopiarSeleccion
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, y As Single)

    Rem Estado Climatico
End Sub

Private Sub insertarLuz_Click()
    
    On Error GoTo insertarLuz_Click_Err
    

    If insertarLuz.Value = True Then
        QuitarLuz.Enabled = False
    Else
        QuitarLuz.Enabled = True

    End If

    
    Exit Sub

insertarLuz_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insertarLuz_Click", Erl)
    Resume Next
    
End Sub

Private Sub insertarParticula_Click()
    
    On Error GoTo insertarParticula_Click_Err
    

    'Call modEdicion.Deshacer_Add("Insertar Particula")
    
    modEdicion.Deshacer_Add "Insertar Particula" ' Hago deshacer
    
    If insertarParticula.Value = True Then
        quitarparticula.Enabled = False
    Else
        quitarparticula.Enabled = True

    End If

    
    Exit Sub

insertarParticula_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insertarParticula_Click", Erl)
    Resume Next
    
End Sub

Private Sub insnpcrandom_Click()
    
    On Error GoTo insnpcrandom_Click_Err
    
    Dim cantidad As Byte
    cantidad = InputBox("Ingrese la cantidad de npcs ingresamos")

    
    Exit Sub

insnpcrandom_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.insnpcrandom_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo lListado_Click_Err
    

    If HotKeysAllow = False Then
        lListado(Index).Tag = lListado(Index).ListIndex

        Select Case Index
    
            Case 0
                cGrh.Text = DameGrhIndex(ReadField(2, lListado(Index).Text, Asc("#")))

                If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa <> 0 Then
                    If LenB(ReadField(2, lListado(Index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                    cCapas.Text = SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa
                Else

                    If LenB(cCapas.Tag) <> 0 Then
                        cCapas.Text = cCapas.Tag
                        cCapas.Tag = vbNullString

                    End If

                End If

                'If SupData(ReadField(2, lListado(index).Text, Asc("#"))).Block = True Then
                '   If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                '    cInsertarBloqueo.value = True
                '   Call cInsertarBloqueo_Click
                ' Else
                '    If LenB(cInsertarBloqueo.Tag) <> 0 Then
                '        cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                '       cInsertarBloqueo.Tag = vbNullString
                '       Call cInsertarBloqueo_Click
                '   End If
                'End If
                Call fPreviewGrh(cGrh.Text)

            Case 1
                cNumFunc(0).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                Call Grh_Render_To_Hdc(picture1, BodyData(NpcData(cNumFunc(0).Text).Body).Walk(3).grhindex, 0, 0, False)

            Case 2
                cNumFunc(1).Text = ReadField(2, lListado(Index).Text, Asc("#"))

            Case 3
                cNumFunc(2).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                Call Grh_Render_To_Hdc(picture1, ObjData(cNumFunc(2).Text).grhindex, 0, 0, False)

            Case 4
                TriggerBox = FrmMain.lListado(4).ListIndex

        End Select

    Else

        Rem lListado(index).ListIndex = lListado(index).Tag
    End If

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

lListado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo lListado_MouseDown_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If Index = 3 And Button = 2 Then
        If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc

    End If

    
    Exit Sub

lListado_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub lListado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo lListado_MouseMove_Err
    

    HotKeysAllow = False

    
    Exit Sub

lListado_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.lListado_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub LuzColor_Click()
    
    On Error GoTo LuzColor_Click_Err
    
    ColorLuz.Text = Selected_Color()

    
    Exit Sub

LuzColor_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.LuzColor_Click", Erl)
    Resume Next
    
End Sub

Public Sub NextMap()
    Call MapPest_Click(5)
End Sub

Private Sub MapPest_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo MapPest_Click_Err
    

    If Index = 5 And Timer4.Enabled = True Then

        Dim arch As String
        arch = App.Path & "\..\Recursos\MiniMap\translados.ini"
        Call WriteVar(arch, MapPest(4).Caption, "Abajo", CLng(MapData(50, 91).TileExit.Map))
        Call WriteVar(arch, MapPest(4).Caption, "Arriba", CLng(MapData(50, 10).TileExit.Map))
        Call WriteVar(arch, MapPest(4).Caption, "Izquierda", CLng(MapData(13, 50).TileExit.Map))
        Call WriteVar(arch, MapPest(4).Caption, "Derecha", CLng(MapData(88, 50).TileExit.Map))
    
        SavePicture MiniMapas2.image, App.Path & "\recursos\minimapas\" & MapPest(4).Caption & ".png"

        If MapPest(5).Visible = False Then
            Timer4.Enabled = False
            Call AddtoRichTextBox(FrmMain.RichTextBox1, "Generacion de minimapas finalizada.", 255, 255, 255, False, True, False)

            Exit Sub

        End If

    End If

    Dim Mapa As Integer
    Mapa = Index + NumMap_Save - 4

    MapaActual = Mapa
    Form3.Label5.Caption = MapaActual
    Label16.Caption = MapaActual

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName

        End If

    End If

    Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".csm"

    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    Call modMapIO.NuevoMapa
    DoEvents
    modMapIO.AbrirMapa Dialog.FileName
    EngineRun = True
    Exit Sub

    Exit Sub

ErrHandler:
    MsgBox Err.Description

    
    Exit Sub

MapPest_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MapPest_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo MiniMap_MouseDown_Err
    
    UserPos.X = CByte(X)
    UserPos.y = CByte(y)
    bRefreshRadar = True

    
    Exit Sub

MiniMap_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub minimapSave_Click()
    
    On Error GoTo minimapSave_Click_Err
    

    SavePicture MiniMapas2.image, App.Path & "\recursos\minimapas\" & MapPest(4).Caption & ".png"

    
    Exit Sub

minimapSave_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.minimapSave_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAbrirMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    DeseaGuardarMapa Dialog.FileName

    ObtenerNombreArchivo False

    If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If
    
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuacercade_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuacercade_Click_Err
    
    frmAbout.Show

    
    Exit Sub

mnuacercade_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuacercade_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCapturarTranslados_Click_Err
    
    mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)

    
    Exit Sub

mnuAutoCapturarTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCapturarTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCapturarSuperficie_Click_Err
    
    mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

    
    Exit Sub

mnuAutoCapturarSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCapturarSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoCompletarSuperficies_Click_Err
    
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

    If mnuAutoCompletarSuperficies.Checked = False Then
        FrmMain.LvBOpcion(18).Caption = "Grh Normal"
    Else
        FrmMain.LvBOpcion(18).Caption = "AutoCompletar"

    End If

    
    Exit Sub

mnuAutoCompletarSuperficies_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoCompletarSuperficies_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoGuardarMapas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoGuardarMapas_Click_Err
    
    frmAutoGuardarMapa.Show

    
    Exit Sub

mnuAutoGuardarMapas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoGuardarMapas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoQuitarFunciones_Click_Err
    
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

    
    Exit Sub

mnuAutoQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuAutoQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuBloquear_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 2 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    modPaneles.VerFuncion 2, True

    
    Exit Sub

mnuBloquear_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquear_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBloquearBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuBloquearBordes_Click_Err
    
    Call modEdicion.Bloquear_Bordes

    
    Exit Sub

mnuBloquearBordes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquearBordes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBloquearMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuBloquearMapa_Click_Err
    
    Call modEdicion.Bloqueo_Todo(&HF)

    
    Exit Sub

mnuBloquearMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquearMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBloquearS_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuBloquearS_Click_Err
    
    Call modEdicion.Deshacer_Add("Bloquear Selección")
    Call BlockearSeleccion

    
    Exit Sub

mnuBloquearS_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuBloquearS_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConfigAvanzada_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuConfigAvanzada_Click_Err
    
    frmConfigSup.Show

    
    Exit Sub

mnuConfigAvanzada_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuConfigAvanzada_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConfigObjTrans_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo mnuConfigObjTrans_Click_Err
    
    Cfg_TrOBJ = cNumFunc(2).Text

    
    Exit Sub

mnuConfigObjTrans_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuConfigObjTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuCopiar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuCopiar_Click_Err
    
    Call CopiarSeleccion

    
    Exit Sub

mnuCopiar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuCopiar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuCortar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuCortar_Click_Err
    
    Call modEdicion.Deshacer_Add("Cortar Selección")
    Call CortarSeleccion

    
    Exit Sub

mnuCortar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuCortar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo mnuDeshacer_Click_Err
    
    Call modEdicion.Deshacer_Recover

    
    Exit Sub

mnuDeshacer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDeshacerPegado_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuDeshacerPegado_Click_Err
    
    Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
    Call DePegar

    
    Exit Sub

mnuDeshacerPegado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacerPegado_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGRHaBMP_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuGRHaBMP_Click_Err
    
    frmGRHaBMP.Show

    
    Exit Sub

mnuGRHaBMP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGRHaBMP_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarcomoBMP_Click()
    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    
    On Error GoTo mnuGuardarcomoBMP_Click_Err
    
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then

    End If

    
    Exit Sub

mnuGuardarcomoBMP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarcomoBMP_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarcomoJPG_Click()
    '*************************************************
    'Author: Salvito
    'Last modified: 01/05/2008 - ^[GS]^
    '*************************************************
    
    On Error GoTo mnuGuardarcomoJPG_Click_Err
    
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))

    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
  
    End If

    
    Exit Sub

mnuGuardarcomoJPG_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarcomoJPG_Click", Erl)
    Resume Next
    
End Sub

Public Sub mnuGuardarMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuGuardarMapa_Click_Err
    
    SavePicture MiniMapas2.image, App.Path & "\recursos\minimapas\" & MapPest(4).Caption & ".png"
    modMapIO.GuardarMapa Dialog.FileName

    
    Exit Sub

mnuGuardarMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarMapaComo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuGuardarMapaComo_Click_Err
    
    modMapIO.GuardarMapa

    
    Exit Sub

mnuGuardarMapaComo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapaComo_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarUltimaConfig_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 23/05/06
    '*************************************************
    Rem mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInfoMap_Click_Err
    
    frmMapInfo.Show
    frmMapInfo.Visible = True

    
    Exit Sub

mnuInfoMap_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInfoMap_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInformes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInformes_Click_Err
    
    frmInformes.Show

    
    Exit Sub

mnuInformes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInformes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInsertarSuperficieAlAzar_Click_Err
    
    Call modEdicion.Superficie_Azar

    
    Exit Sub

mnuInsertarSuperficieAlAzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInsertarSuperficieAlAzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInsertarSuperficieEnBordes_Click_Err
    
    Call modEdicion.Superficie_Bordes

    
    Exit Sub

mnuInsertarSuperficieEnBordes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInsertarSuperficieEnBordes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInsertarSuperficieEnTodo_Click_Err
    
    Call modEdicion.Superficie_Todo

    
    Exit Sub

mnuInsertarSuperficieEnTodo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInsertarSuperficieEnTodo_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuInsertarTransladosAdyasentes_Click_Err
    
    frmUnionAdyacente.Show

    
    Exit Sub

mnuInsertarTransladosAdyasentes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuInsertarTransladosAdyasentes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuManual_Click()
    
    On Error GoTo mnuManual_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) <> 0 Then
        Call Shell("explorer " & App.Path & "\manual\index.html")
        DoEvents

    End If

    
    Exit Sub

mnuManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuModoCaminata_Click_Err
    
    ToggleWalkMode

    
    Exit Sub

mnuModoCaminata_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuModoCaminata_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuNPCs_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 3 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 3, True

    
    Exit Sub

mnuNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuNuevoMapa_Click_Err
    

    Dim loopc As Integer

    DeseaGuardarMapa Dialog.FileName

    For loopc = 0 To FrmMain.MapPest.Count
        FrmMain.MapPest(loopc).Visible = False
    Next

    FrmMain.Dialog.FileName = Empty

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If

    Call modMapIO.NuevoMapa

    Call cmdInformacionDelMapa_Click

    
    Exit Sub

mnuNuevoMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuNuevoMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuObjetos_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 5 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 5, True

    
    Exit Sub

mnuObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo mnuOptimizar_Click_Err
    
    frmOptimizar.Show

    
    Exit Sub

mnuOptimizar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuOptimizar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuPegar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuPegar_Click_Err
    
    Call modEdicion.Deshacer_Add("Pegar Selección")
    Call PegarSeleccion

    
    Exit Sub

mnuPegar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuPegar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQBloquear_Click_Err
    
    modPaneles.VerFuncion 2, False

    
    Exit Sub

mnuQBloquear_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQBloquear_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQNPCs_Click_Err
    
    modPaneles.VerFuncion 3, False

    
    Exit Sub

mnuQNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQObjetos_Click_Err
    
    modPaneles.VerFuncion 5, False

    
    Exit Sub

mnuQObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQSuperficie_Click_Err
    
    modPaneles.VerFuncion 0, False

    
    Exit Sub

mnuQSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQTranslados_Click_Err
    
    modPaneles.VerFuncion 1, False

    
    Exit Sub

mnuQTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQTriggers_Click_Err
    
    modPaneles.VerFuncion 6, False

    
    Exit Sub

mnuQTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarBloqueos_Click_Err
    
    Call modEdicion.Bloqueo_Todo(0)

    
    Exit Sub

mnuQuitarBloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarFunciones_Click_Err
    

    ' Superficies
    cSeleccionarSuperficie.Value = False
    Call cSeleccionarSuperficie_Click
    cQuitarEnEstaCapa.Value = False
    Call cQuitarEnEstaCapa_Click
    cQuitarEnTodasLasCapas.Value = False
    Call cQuitarEnTodasLasCapas_Click

    ' Translados
    cQuitarTrans.Value = False
    Call cQuitarTrans_Click
    cInsertarTrans.Value = False
    Call cInsertarTrans_Click

    ' Bloqueos
    cQuitarBloqueo.Value = False
    Call cQuitarBloqueo_Click
    cInsertarBloqueo.Value = False
    Call cInsertarBloqueo_Click

    ' Otras funciones
    cInsertarFunc(0).Value = False
    Call cInsertarFunc_Click(0)
    cInsertarFunc(1).Value = False
    Call cInsertarFunc_Click(1)
    cInsertarFunc(2).Value = False
    Call cInsertarFunc_Click(2)
    cQuitarFunc(0).Value = False
    Call cQuitarFunc_Click(0)
    cQuitarFunc(1).Value = False
    Call cQuitarFunc_Click(1)
    cQuitarFunc(2).Value = False
    Call cQuitarFunc_Click(2)

    ' Triggers
    cInsertarTrigger.Value = False
    Call cInsertarTrigger_Click
    cQuitarTrigger.Value = False
    Call cQuitarTrigger_Click

    ' particulas
    insertarParticula.Value = False
    Call insertarParticula_Click
    quitarparticula.Value = False
    Call quitarparticula_Click

    ' Luces
    insertarLuz.Value = False
    Call insertarLuz_Click
    QuitarLuz.Value = False
    Call QuitarLuz_Click

    
    Exit Sub

mnuQuitarFunciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarNPCs_Click_Err
    
    Call modEdicion.Quitar_NPCs(False)

    
    Exit Sub

mnuQuitarNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarNPCs_Click", Erl)
    Resume Next
    
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarObjetos_Click_Err
    
    Call modEdicion.Quitar_Objetos

    
    Exit Sub

mnuQuitarObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarObjetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarSuperficieBordes_Click_Err
    
    Call modEdicion.Quitar_Bordes

    
    Exit Sub

mnuQuitarSuperficieBordes_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarSuperficieBordes_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarSuperficieDeCapa_Click_Err
    
    Call modEdicion.Quitar_Capa(cCapas.Text)

    
    Exit Sub

mnuQuitarSuperficieDeCapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarSuperficieDeCapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTODO_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarTODO_Click_Err
    
    Call modEdicion.Borrar_Mapa

    
    Exit Sub

mnuQuitarTODO_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTODO_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo mnuQuitarTranslados_Click_Err
    
    Call modEdicion.Quitar_Translados

    
    Exit Sub

mnuQuitarTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuQuitarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuQuitarTriggers_Click_Err
    
    Call modEdicion.Quitar_Triggers

    
    Exit Sub

mnuQuitarTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuReAbrirMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ErrHandler

    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName

        End If

    End If

    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuRealizarOperacion_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuRealizarOperacion_Click_Err
    

    Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
    mnuAutoCompletarSuperficies.Checked = False

    Call AccionSeleccion

    
    Exit Sub

mnuRealizarOperacion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuRealizarOperacion_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuSalir_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuSalir_Click_Err
    
    Unload Me

    
    Exit Sub

mnuSalir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuSalir_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuSuperficie_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 0 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 0, True

    
    Exit Sub

mnuSuperficie_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuSuperficie_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuTranslados_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 1 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 1, True

    
    Exit Sub

mnuTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuTriggers_Click_Err
    
    Dim i As Byte

    For i = 0 To 6

        If i <> 6 Then
            FrmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    modPaneles.VerFuncion 6, True

    
    Exit Sub

mnuTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuUtilizarDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo mnuUtilizarDeshacer_Click_Err
    
    mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)

    
    Exit Sub

mnuUtilizarDeshacer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuUtilizarDeshacer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerAutomatico_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerAutomatico_Click_Err
    
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)

    
    Exit Sub

mnuVerAutomatico_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerAutomatico_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerBloqueos_Click_Err
    
    cVerBloqueos.Value = (cVerBloqueos.Value = False)
    mnuVerBloqueos.Checked = cVerBloqueos.Value

    
    Exit Sub

mnuVerBloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa1_Click()
    
    On Error GoTo mnuVerCapa1_Click_Err
    
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)

    
    Exit Sub

mnuVerCapa1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa2_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa2_Click_Err
    
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)

    
    Exit Sub

mnuVerCapa2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa3_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa3_Click_Err
    
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)

    
    Exit Sub

mnuVerCapa3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa4_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa4_Click_Err
    
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)

    
    Exit Sub

mnuVerCapa4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerCapa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerGrilla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo mnuVerGrilla_Click_Err
    
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla

    
    Exit Sub

mnuVerGrilla_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerGrilla_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerLuces_Click()
    
    On Error GoTo mnuVerLuces_Click_Err
    
    mnuVerLuces.Checked = (mnuVerLuces.Checked = False)

    
    Exit Sub

mnuVerLuces_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerLuces_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerNPCs_Click_Err
    
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

    
    Exit Sub

mnuVerNPCs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerNPCs_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerObjetos_Click_Err
    
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

    
    Exit Sub

mnuVerObjetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerObjetos_Click", Erl)
    Resume Next
    
End Sub

Public Sub mnuVerParticulas_Click()
    
    On Error GoTo mnuVerParticulas_Click_Err
    

    mnuVerParticulas.Checked = (mnuVerParticulas.Checked = False)

    
    Exit Sub

mnuVerParticulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerParticulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerTranslados_Click_Err
    
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

    
    Exit Sub

mnuVerTranslados_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerTriggers_Click_Err
    
    cVerTriggers.Value = (cVerTriggers.Value = False)
    mnuVerTriggers.Checked = cVerTriggers.Value

    
    Exit Sub

mnuVerTriggers_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuVerTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo picRadar_MouseDown_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If X < 11 Then X = 11
    If X > 89 Then X = 89
    If y < 10 Then y = 10
    If y > 92 Then y = 92
    UserPos.X = X
    UserPos.y = y
    bRefreshRadar = True

    
    Exit Sub

picRadar_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.picRadar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo picRadar_MouseMove_Err
    
    MiRadarX = X
    MiRadarY = y

    
    Exit Sub

picRadar_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.picRadar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo Form_QueryUnload_Err
    

    ' Guardar configuración
    Rem WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(FrmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")

    WriteVar IniPath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(FrmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(FrmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(FrmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(FrmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(FrmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(FrmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(FrmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(FrmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Marco", IIf(FrmMain.mnuVerMarco.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(FrmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(FrmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.y
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(FrmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(FrmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(FrmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)

    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

    
    Exit Sub

Form_QueryUnload_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub Npcalazarpormapa_Click()
    
    On Error GoTo Npcalazarpormapa_Click_Err
    
    Dim NPCIndex As Long
    Dim X        As Byte
    Dim tmp      As String
    Dim y        As Byte
    Dim i        As Byte

    tmp = InputBox("¿Cuantos npcs?", "Ingresar npcs al azar por todo el mapa.")

    If tmp = "" Then Exit Sub

    For i = 1 To CLng(tmp)
        X = RandomNumber(15, 87)
        y = RandomNumber(15, 87)
            
        If (MapData(X, y).Blocked And &HF) <> &HF Then

            NPCIndex = FrmMain.cNumFunc(0).Text
                
            If NPCIndex <> MapData(X, y).NPCIndex Then
                modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
             
                Call MakeChar(NextOpenChar(), NpcData(NPCIndex).Body, NpcData(NPCIndex).Head, NpcData(NPCIndex).Heading, X, y)
                MapData(X, y).NPCIndex = NPCIndex

            End If

        End If

    Next i

    
    Exit Sub

Npcalazarpormapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Npcalazarpormapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub objalazar_Click()
    
    On Error GoTo objalazar_Click_Err
    

    Dim cantidad As Long
    Dim bloquear As Byte
    Dim objeto   As Long
    Dim X        As Byte
    Dim y        As Byte
    Dim i        As Long

    cantidad = InputBox("Ingrese la cantidad de objetos a mapear")

    If cantidad <= 0 Then Exit Sub
    bloquear = InputBox("¿Desea bloquear los obejtos? (1= SI | 0 = NO")

    If bloquear > 1 Then Exit Sub
    objeto = FrmMain.cNumFunc(2).Text

    For i = 1 To cantidad
        X = RandomNumber(10, 91)
        y = RandomNumber(8, 93)

        If MapData(X, y).Graphic(1).grhindex < 1505 Or MapData(X, y).Graphic(1).grhindex > 1520 Then
            
            MapInfo.Changed = 1 'Set changed flag
                
            MapData(X, y).Blocked = bloquear * &HF
        
            InitGrh MapData(X, y).ObjGrh, ObjData(objeto).grhindex
            MapData(X, y).OBJInfo.objindex = objeto
            MapData(X, y).OBJInfo.Amount = 1

        End If
            
    Next i

    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Se agregaron " & cantidad & " " & ObjData(objeto).Name & " al mapa.", 255, 255, 255, False, True, False)

    
    Exit Sub

objalazar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.objalazar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Objeto_Click()
    
    On Error GoTo Objeto_Click_Err
    
    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            'If MapData(X, Y).OBJInfo.objindex = Text1 Then
            '         InitGrh MapData(X, Y).ObjGrh, 1
            '        MapData(X, Y).OBJInfo.objindex = Text2
            '         MapData(X, Y).OBJInfo.Amount = 1
            ' End If
        Next X
    Next y

    
    Exit Sub

Objeto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Objeto_Click", Erl)
    Resume Next
    
End Sub

Private Sub openminimap_Click()
    
    On Error GoTo openminimap_Click_Err
    

    Dim ret As Long
    ret = ShellExecute(Me.hWnd, "Open", App.Path & "\recursos\index.htm", "", "", 1)

    
    Exit Sub

openminimap_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.openminimap_Click", Erl)
    Resume Next
    
End Sub

Private Sub pasto_Click()
    
    On Error GoTo pasto_Click_Err
    
    cGrh.Text = DameGrhIndex(0)

    Call modPaneles.VistaPreviaDeSup

    
    Exit Sub

pasto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.pasto_Click", Erl)
    Resume Next
    
End Sub

Private Sub Picture3_Click()
    
    On Error GoTo Picture3_Click_Err
    
    If ColorAmb = 0 Then Exit Sub
    
    ColorAmb = Selected_Color()
    LuzMapa = ColorAmb
    
    Dim BackC As Long
    
    Dim r, g, b As Byte
    r = (LuzMapa And 16711680) / 65536
    g = (LuzMapa And 65280) / 256
    b = LuzMapa And 255
    
    BackC = RGB(r, g, b)
    
    Picture3.BackColor = BackC
    LuzMapa.Text = Hex(LuzMapa)

    engine.Map_Base_Light_Set ColorAmb
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Luz de mapa aceptada. Luz: " & ColorAmb & ".", 255, 255, 255, False, True, False)
    MapInfo.Changed = 1

    
    Exit Sub

Picture3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Picture3_Click", Erl)
    Resume Next
    
End Sub

Private Sub ProbarAmbiental_Click()
    
    On Error GoTo ProbarAmbiental_Click_Err
    
    WavAmbiental.Show

    
    Exit Sub

ProbarAmbiental_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.ProbarAmbiental_Click", Erl)
    Resume Next
    
End Sub

Private Sub QuitarLuz_Click()
    
    On Error GoTo QuitarLuz_Click_Err
    

    If QuitarLuz.Value = True Then
        insertarLuz.Enabled = False
    Else
        insertarLuz.Enabled = True

    End If

    
    Exit Sub

QuitarLuz_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.QuitarLuz_Click", Erl)
    Resume Next
    
End Sub

Private Sub quitarparticula_Click()
    
    On Error GoTo quitarparticula_Click_Err
    

    If quitarparticula.Value = True Then
        insertarParticula.Enabled = False
    Else
        insertarParticula.Enabled = True

    End If

    
    Exit Sub

quitarparticula_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.quitarparticula_Click", Erl)
    Resume Next
    
End Sub

Private Sub render_mapa_Click()
    
    On Error GoTo render_mapa_Click_Err
    
    'Radio = Val(InputBox("Escriba la escala de 1 a 5 en la que generemos su mapa", "la escala se multiplica x 32")) 'ver ReyarB
    Radio = 1
    
    If Radio = 0 Then Radio = 1
    If Radio >= 5 Then Radio = 5

    FrmRender.Show

    
    Exit Sub

render_mapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.render_mapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub renderer_Click()
    
    On Error GoTo renderer_Click_Err
    

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    UltimoClickX = tX
    UltimoClickY = tY

    If DesdeBloq = True Then
        RepetirSup = False
        modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
        DesdeBloq = False
        Call PonerGrh
        Call DibujarMiniMapa

        If RepetirSup Then
            Call InsertarBloque

        End If

    End If

    
    Exit Sub

renderer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_Click", Erl)
    Resume Next
    
End Sub

Private Sub renderer_DblClick()
    
    On Error GoTo renderer_DblClick_Err
    
    Dim tX As Integer
    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)

    End If

    
    Exit Sub

renderer_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo renderer_MouseDown_Err
    

    If Not MapaCargado Then Exit Sub

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
    Else
        ClickEdit Button, tX, tY

    End If

    
    Exit Sub

renderer_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo renderer_MouseMove_Err
    

    MouseX = X
    MouseY = y

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    POSX.Caption = "X: " & tX & " - Y: " & tY

    If tX < 14 Or tY < 11 Or tX > 87 Or tY > 90 Then
        POSX.ForeColor = vbRed
    Else
        POSX.ForeColor = vbWhite

    End If

    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
    Else

        If tX = 0 Then Exit Sub
        If tY = 0 Then Exit Sub
        If tX = LastX And tY = LastY Then Exit Sub
        
        ClickEdit Button, tX, tY
        
        LastX = tX
        LastY = tY

    End If

    
    Exit Sub

renderer_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.renderer_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub SaveAllMiniMap_Click()
    
    On Error GoTo SaveAllMiniMap_Click_Err
    

    If MsgBox("Esta funcion generara todos los minimap de nuevo. ¿Esta seguro que desea continuar?", vbExclamation + vbYesNo) = vbYes Then
        Timer4.Enabled = True

    End If

    
    Exit Sub

SaveAllMiniMap_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.SaveAllMiniMap_Click", Erl)
    Resume Next
    
End Sub

Public Sub SelectPanel_Click(Index As Integer)
    
    On Error GoTo SelectPanel_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 8

        If i <> Index Then
            SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click

    Call VerFuncion(Index, SelectPanel(Index).Value)

    
    Exit Sub

SelectPanel_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.SelectPanel_Click", Erl)
    Resume Next
    
End Sub



Private Sub Stopminimap_Click()
    
    On Error GoTo Stopminimap_Click_Err
    
    Timer4.Enabled = False

    
    Exit Sub

Stopminimap_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Stopminimap_Click", Erl)
    Resume Next
    
End Sub

Private Sub Text3_Change()

End Sub

Private Sub TiggerEspecial_Click()

    On Error Resume Next

    TriggerBox = InputBox("Ingrese el numero de trigger a usar.")

End Sub

Private Sub TimAutoGuardarMapa_Timer()
    
    On Error GoTo TimAutoGuardarMapa_Timer_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If mnuAutoGuardarMapas.Checked = True Then
        bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1

        If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
            If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
                modMapIO.GuardarMapa Dialog.FileName

            End If

            bAutoGuardarMapaCount = 0

        End If

    End If

    
    Exit Sub

TimAutoGuardarMapa_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.TimAutoGuardarMapa_Timer", Erl)
    Resume Next
    
End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    With Dialog

        If FormatoIAO Then
            .Filter = "Mapas de RevolucionAO (*.csm)|*.csm"
        Else
            .Filter = "Mapas de ArgentumOnline (*.map)|*.map"

        End If

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .FLAGS = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .FileName = vbNullString
            .FLAGS = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub

Private Sub Timer1_Timer()
    
    On Error GoTo Timer1_Timer_Err
    
    Call MapPest_Click(5)
    modMapIO.GuardarMapa Dialog.FileName

    
    Exit Sub

Timer1_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Timer1_Timer", Erl)
    Resume Next
    
End Sub

Private Sub Timer2_Timer()
    
    On Error GoTo Timer2_Timer_Err
    

    If engine.bRunning Then engine.Engine_ActFPS

    
    Exit Sub

Timer2_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Timer2_Timer", Erl)
    Resume Next
    
End Sub

Private Sub Timer4_Timer()
    
    On Error GoTo Timer4_Timer_Err
    

    'Call Command2_Click
    'Call modEdicion.Quitar_NPCs(False)
    'Call Form3.Command1_Click
    'Call DesplazarTranslados.Command1_Click
    'Call frmOptimizar.cOptimizar_Click
    'Call borrarnegros_Click

    'modMapIO.GuardarMapa Dialog.FileName
    'Call Form3.HacerTranslados
    'modMapIO.GuardarMapa Dialog.FileName
    SavePicture MiniMapas2.image, App.Path & "\recursos\minimapas\" & MapPest(4).Caption & ".png"
    'Call mnuGuardarMapa_Click
    'Call Command2_Click
    'modMapIO.GuardarMapa Dialog.FileName
    Timer4.Interval = 1
    Call MapPest_Click(5)

    
    Exit Sub

Timer4_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Timer4_Timer", Erl)
    Resume Next
    
End Sub

Private Sub Todas_las_luces_Click()
    
    On Error GoTo Todas_las_luces_Click_Err
    
    Dim X As Byte
    Dim y As Byte
    Dim i As Long

    For X = 1 To 100
        For y = 1 To 100

            MapData(X, y).luz.Rango = 0
        Next y
    Next X

    engine.Light_Remove_All
    MapInfo.Changed = 1

    
    Exit Sub

Todas_las_luces_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Todas_las_luces_Click", Erl)
    Resume Next
    
End Sub

Private Sub Todas_las_Particulas_Click()
    
    On Error GoTo Todas_las_Particulas_Click_Err
    
    Dim X As Byte
    Dim y As Byte
    Dim i As Long

    For X = 1 To 100
        For y = 1 To 100
            MapData(X, y).particle_Index = 0
        Next y
    Next X

    engine.Particle_Group_Remove_All
    MapInfo.Changed = 1

    
    Exit Sub

Todas_las_Particulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.Todas_las_Particulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub txtMapTerreno_Click()
    
    On Error GoTo txtMapTerreno_Click_Err
    
    MapDat.terrain = txtMapTerreno
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Terreno de mapa cambiada a: " & MapDat.terrain, 255, 255, 255, False, True, False)
    MapInfo.Changed = 1

    
    Exit Sub

txtMapTerreno_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.txtMapTerreno_Click", Erl)
    Resume Next
    
End Sub

Private Sub txtMapZona_Click()
    
    On Error GoTo txtMapZona_Click_Err
    
    MapDat.zone = txtMapZona
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Zona de mapa cambiada a: " & MapDat.zone, 255, 255, 255, False, True, False)
    MapInfo.Changed = 1

    
    Exit Sub

txtMapZona_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.txtMapZona_Click", Erl)
    Resume Next
    
End Sub

Private Sub TxtMidi_Change()
    
    On Error GoTo TxtMidi_Change_Err
    

    If Not IsNumeric(TxtMidi) Then Exit Sub
    MidiMusic = CInt(TxtMidi)
    MapDat.music_numberLow = MidiMusic
    MapInfo.Changed = 1

    
    Exit Sub

TxtMidi_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.TxtMidi_Change", Erl)
    Resume Next
    
End Sub

Private Sub TxtMp3_Change()
    
    On Error GoTo TxtMp3_Change_Err
    

    If Not IsNumeric(TxtMp3) Then Exit Sub
    Mp3Music = CInt(TxtMp3)
    MapInfo.Changed = 1

    
    Exit Sub

TxtMp3_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.TxtMp3_Change", Erl)
    Resume Next
    
End Sub

Private Sub txtnamemapa_Change()
    
    On Error GoTo txtnamemapa_Change_Err
    
    MapDat.map_name = txtnamemapa
    Call AddtoRichTextBox(FrmMain.RichTextBox1, "Nombre de mapa cambiado a:  " & MapDat.map_name, 255, 255, 255, False, True, False)
    MapInfo.Changed = 1

    
    Exit Sub

txtnamemapa_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.txtnamemapa_Change", Erl)
    Resume Next
    
End Sub

Private Sub TxtWav_Change()
    
    On Error GoTo TxtWav_Change_Err
    
    Ambiente = TxtWav
    MapInfo.Changed = 1

    
    Exit Sub

TxtWav_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.TxtWav_Change", Erl)
    Resume Next
    
End Sub

Private Sub vergraficoslistado_Click()
    
    On Error GoTo vergraficoslistado_Click_Err
    
    Form1.Show , FrmMain

    
    Exit Sub

vergraficoslistado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.vergraficoslistado_Click", Erl)
    Resume Next
    
End Sub
