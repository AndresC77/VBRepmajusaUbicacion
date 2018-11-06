VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmV_FacPed 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturar Pedidos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
   Icon            =   "frmV_FacPed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13875
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   9600
      TabIndex        =   89
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   11280
      TabIndex        =   87
      Top             =   720
      Width           =   495
   End
   Begin VB.Frame frmFiltro 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Facturacion Programada"
         Height          =   2055
         Left            =   120
         TabIndex        =   75
         Top             =   4800
         Width           =   6735
         Begin VB.CheckBox chkContado 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Solo Contado"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   360
            TabIndex        =   88
            Top             =   960
            Width           =   1815
         End
         Begin VB.Timer TmrAct 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   4440
            Top             =   600
         End
         Begin VB.CommandButton cmdProgramar 
            Caption         =   "Programar"
            Height          =   375
            Left            =   4920
            TabIndex        =   82
            Top             =   600
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFechaIni 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   76
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   66387971
            Format          =   66453507
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   77
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   66453507
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpEjecutar 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   4440
            TabIndex        =   80
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   66453507
            CurrentDate     =   37463
         End
         Begin MSComctlLib.ProgressBar pbDetalle 
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   1680
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar pbGeneral 
            Height          =   375
            Left            =   240
            TabIndex        =   83
            Top             =   1320
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Max             =   11
            Scrolling       =   1
         End
         Begin VB.CommandButton cmdEjecutar 
            Caption         =   "Ejecutar"
            Height          =   375
            Left            =   3360
            TabIndex        =   86
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2850
            TabIndex        =   85
            Top             =   1080
            Width           =   3645
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fac."
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3600
            TabIndex        =   81
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   79
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   78
            Top             =   660
            Width           =   705
         End
      End
      Begin VB.OptionButton optCalendarioAyer 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Segun Calendario -1"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2040
         TabIndex        =   72
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optAbierto 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Calendario Abierto"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5040
         TabIndex        =   31
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optFiltro 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Segun Filtro"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optCalendario 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Segun Calendario"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   255
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   1800
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   2160
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   2520
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   315
         Left            =   840
         TabIndex        =   12
         Top             =   2880
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   315
         Left            =   840
         TabIndex        =   13
         Top             =   3240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   3600
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   3960
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPais 
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Top             =   4440
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provinc:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   165
         TabIndex        =   28
         Top             =   4500
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   25
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   24
         Top             =   4020
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   23
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   22
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   21
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   465
         TabIndex        =   20
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N1:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   450
         TabIndex        =   19
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N2:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   450
         TabIndex        =   18
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N3:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N4:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   450
         TabIndex        =   16
         Top             =   2220
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "ACTUALIZAR"
      Height          =   495
      Left            =   7200
      TabIndex        =   71
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame frmPrincipal 
      BackColor       =   &H00DDDDDD&
      Height          =   6135
      Left            =   2880
      TabIndex        =   32
      Top             =   840
      Width           =   10935
      Begin MSComCtl2.DTPicker dtpFechaFinAplicar 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   73
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   66453507
         CurrentDate     =   37463
      End
      Begin VB.Frame frmPed 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Listado de Pedidos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5655
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   10695
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Salir"
            Height          =   375
            Left            =   9240
            TabIndex        =   70
            Top             =   4005
            Width           =   1335
         End
         Begin VB.CommandButton CmdConfirmar 
            Caption         =   "Facturar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9120
            TabIndex        =   69
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton CmdConfirmarTodo 
            Caption         =   "Facturar Todo"
            Height          =   375
            Left            =   7920
            TabIndex        =   55
            Top             =   4005
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtDescuento 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   2175
            Width           =   615
         End
         Begin VB.Frame frmRec 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Recargos:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1695
            Left            =   120
            TabIndex        =   40
            Top             =   3885
            Width           =   7695
            Begin VB.TextBox TxtRecargo 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   12298
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtIva 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   12298
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox TxtDesc 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   12298
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   12298
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtSubTotal 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   12298
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtObserv 
               Height          =   285
               Left            =   240
               MaxLength       =   250
               TabIndex        =   41
               Top             =   1320
               Width           =   4215
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
               Height          =   855
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   4305
               _cx             =   1980243370
               _cy             =   1980237284
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmV_FacPed.frx":030A
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   1
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Image imgBtnUp 
               Height          =   210
               Left            =   4440
               Picture         =   "frmV_FacPed.frx":038A
               ToolTipText     =   "Elimina una Fila"
               Top             =   240
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Image imgBtnDn 
               Height          =   210
               Left            =   4680
               Picture         =   "frmV_FacPed.frx":04C0
               Top             =   240
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal pedido:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   5100
               TabIndex        =   53
               Top             =   270
               Width           =   1155
            End
            Begin VB.Label LblIva 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IVA X%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   5160
               TabIndex        =   52
               Top             =   750
               Width           =   570
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descuento:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   5160
               TabIndex        =   51
               Top             =   510
               Width           =   825
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recargos:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   5160
               TabIndex        =   50
               Top             =   990
               Width           =   750
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total pedido:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   5160
               TabIndex        =   49
               Top             =   1350
               Width           =   1065
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   210
               Left            =   300
               TabIndex        =   48
               Top             =   1080
               Width           =   1185
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
            Height          =   1815
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   10515
            _cx             =   1980254323
            _cy             =   1980238977
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   21
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmV_FacPed.frx":05EC
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   315
            Left            =   8760
            TabIndex        =   56
            Top             =   2175
            Width           =   1815
            _extentx        =   2355
            _extenty        =   556
            value           =   41816.5214351852
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFG 
            Height          =   1335
            Left            =   120
            TabIndex        =   57
            Top             =   2565
            Width           =   8595
            _cx             =   1980250937
            _cy             =   1980238131
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmV_FacPed.frx":0843
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   0
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   1
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSDataListLib.DataCombo CmbFpago 
            Height          =   315
            Left            =   8760
            TabIndex        =   58
            Top             =   2805
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbTC 
            Height          =   315
            Left            =   6240
            TabIndex        =   59
            Top             =   2175
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo CmbTipoFac 
            Height          =   315
            Left            =   8760
            TabIndex        =   60
            Top             =   3405
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFGTot 
            Height          =   1095
            Left            =   8280
            TabIndex        =   61
            Top             =   4485
            Width           =   2115
            _cx             =   1980239507
            _cy             =   1980237707
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmV_FacPed.frx":098C
            ScrollTrack     =   0   'False
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lista:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5760
            TabIndex        =   68
            Top             =   2205
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dcto (%):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4200
            TabIndex        =   67
            Top             =   2205
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   8160
            TabIndex        =   66
            Top             =   2205
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   8760
            TabIndex        =   65
            Top             =   2565
            Width           =   1125
         End
         Begin VB.Label LblPedido 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   2040
            TabIndex        =   64
            Top             =   2205
            Width           =   60
         End
         Begin VB.Label LblDetalle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle del Pedido Nº"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   330
            TabIndex        =   63
            Top             =   2205
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Factura:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   8760
            TabIndex        =   62
            Top             =   3165
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.CheckBox chkCIRUC 
         BackColor       =   &H00DDDDDD&
         Caption         =   "CI/RUC"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   135
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPedido 
         Height          =   285
         Left            =   8280
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optListaPedido 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Por Listado de Pedidos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   135
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optNoPedido 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Por Pedido"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   135
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6240
         TabIndex        =   74
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido a Buscar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6960
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdImprimirBloque 
      Caption         =   "ReImprimir en Bloque"
      Height          =   495
      Left            =   9000
      TabIndex        =   26
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCambiarBloc 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   12720
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFacturaHasta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   270
      Width           =   1215
   End
   Begin VB.TextBox txtFacturaDesde 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fac.Hasta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   11295
      TabIndex        =   3
      Top             =   300
      Width           =   795
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fac.Desde:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   11235
      TabIndex        =   1
      Top             =   30
      Width           =   855
   End
End
Attribute VB_Name = "frmV_FacPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################
'#  Forma para ver un pedido ya confirmado de bodega que está listo para ser
'#  facturado.
'#  frmV_VerPedConfirm V1.0
'#  Copyright (C) 2002
'#
'#  Opciones que permite:
'#  *   En una lista se despliegan los pedidos confirmados con sus detalles de
'#      cabecera como el cliente y el vendedor que lo atiende y el estado del
'#      mismo.
'#  *   De igual manera es necesario seleccionar el tipo de facturación que se
'#      va a aplicar al pedido.
'#  *   Es necesario también seleccionar la forma de pago.
'#  *   El usuario puede seleccionar los posibles recargos que puede generar
'#      la facturación de un pedido.
'#
'#  Procesos internos que maneja:
'#  *   La lista que muestra los distintos pedidos, se refresca automáticamente
'#      cada 20 segundos para buscar un nuevo pedido confirmado.
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los
'#      detalles del mismo en un segundo grid.
'#  *   Una vez que el pedido ha sido facturado su estado pasa a vendido.
'#  *   Se pueden ver solo los pedidos que están confirmados y los que ya
'#      se han vendido el día de hoy.
'#  *   Una vez que se va a facturar el pedido se generan automáticamente las
'#      respectivas retenciones que puede tener un cliente.
'#
'#  Tablas que maneja:
'#
'#  persona:
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el
'#      pedido que se está confirmando.
'#  *   También se extrae el nombre del vendedor asignado al pedido.
'#  pedido:
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.
'#  det_pedido:
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.
'#  persona_ret:
'#  *   De esta tala se extraen las diferentes retenciones que puede tener un
'#      cliente determinado para luego aplicarlas a esta factura.
'#  retencion:
'#  *   De aquí se extraen los valores y descripciones de las retenciones, que
'#      se aplicarán posteriormente.
'#  existencia:
'#  *   En esta tabla se actualizan las cantidades existentes de los productos
'#      vendidos.
'#  det_egreso:
'#  *   En esta tabla se guardan los detalles del nuevo documento de egreso de
'#      productos.
'#  ocargo:
'#  *   De esta tabla se extraen los diferentes recargos que se puede manejar
'#      al realizar un nuevo egreso de productos de bodega, como pueden ser:
'#      transporte, fletes, etc.
'#  det_egreso_c:
'#  *   En esta tabla se guardan los diferentes recargos que puede tener esta
'#      nueva compra o egreso de productos.
'#  det_egreso_ret:
'#  *   En esta tabla se guardan los valores de las retenciones aplicadas a este
'#      ingreso de productos a bodega.
'#
'################################################################################

Private clsPedidos As New clsConsulta
Private clsPed As New clsConsulta
Private clsSql As New clsConsulta
Private clsTFac As New clsConsulta
Private clsTC As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsFPago As New clsConsulta
Private clsFormaPago As New clsConsulta
Private clsRet As New clsConsulta
Private clsLstPrds As New clsConsulta
Private clsExis As New clsConsulta
Private strSql As String
Private IVA As Double
Private CodPer As String
Private codCot As String
Private codVen As String
Private codTC As String

Private GuiaAutomatica As Boolean

Private FechaUltFac As String
Private strClaveMAESTRA As String
Private MINCredito As Double
Private SecPublico As Boolean
Private SinIVA As Boolean
Private maxItem As Long
Private tipoPagoDetalle As String
Private strListaFactura As String
Private strListaFacturaPed As String
Private strListaPedido As String
Private lngNFacNPed As Long
Private strCantidadTipoPedido As String

Private Sub chkCIRUC_Click()
    cmbNegocio_Change
End Sub

Private Sub cmdLimpiar_Click()
    'Limpia el contenido del grid de detalles
    VSFG.Clear 1
    VSFG.Rows = 2
    VSFGReca.Clear 1
    VSFGReca.Rows = 2
    VSFGReca.Enabled = False
    CmdConfirmar.Enabled = False
    'cmdPreFactura.Enabled = False
    'CmdGuiaRemi.Enabled = False
    'cmdReserva.Enabled = False
    'CmdDeBaja.Enabled = False
    TxtSubTotal = ""
    TxtTotal = ""
    TxtRecargo = ""
    TxtIva = ""
    TxtDesc = ""
    txtDescuento = ""
    cmbTC = ""
    LblPedido = "-"
    CmbTipoFac = ""
    CmbFpago = ""
    TxtObserv = ""
End Sub


Private Sub cmbNegocio_Change()
    Dim strCli As String
    Dim strFiltro As String
    cmdLimpiar_Click
    strCantidadTipoPedido = "det_ped_cant_confirmada"
    If Me.cmbNegocio.BoundText = "JON" Or Me.cmbNegocio.BoundText = "LEM" Then
        strCantidadTipoPedido = "det_ped_cant_entregada"
    End If
    strSql = " SELECT tip_ped_ptofac,tip_ped_fac_desde,tip_ped_fac_hasta,tip_ped_guia_automatica " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
        If GeneraDocElec = 1 Then
            lngFacturaDesde = clsSql.adorec_Def("tip_ped_fac_desde")
            lngFacturaHasta = clsSql.adorec_Def("tip_ped_fac_hasta")
        End If
        GuiaAutomatica = IIf(clsSql.adorec_Def("tip_ped_guia_automatica") = 0, False, True)
    End If
    If optFiltro.Value = True Then
        strFiltro = " AND persona.per_codigo_ref like '" & IIf(FormatoD0(cmbGerente.BoundText) = -1, "%", cmbGerente.BoundText) & "'" & _
                     " AND persona.per_codigo_ref2 like '" & IIf(FormatoD0(cmbDirector.BoundText) = -1, "%", cmbDirector.BoundText) & "'" & _
                     " AND persona.per_codigo_ref3 like '" & IIf(FormatoD0(cmbEmprendedor.BoundText) = -1, "%", cmbEmprendedor.BoundText) & "'" & _
                     " AND persona.per_codigo_ref4 like '" & IIf(FormatoD0(cmbEjecutivo.BoundText) = -1, "%", cmbEjecutivo.BoundText) & "'" & _
                     " AND persona.per_codigo_ref5 like '" & IIf(FormatoD0(cmbN5.BoundText) = -1, "%", cmbN5.BoundText) & "'" & _
                     " AND persona.per_codigo_ref6 like '" & IIf(FormatoD0(cmbN6.BoundText) = -1, "%", cmbN6.BoundText) & "'" & _
                     " AND persona.per_codigo_ref7 like '" & IIf(FormatoD0(cmbN7.BoundText) = -1, "%", cmbN7.BoundText) & "'" & _
                     " AND persona.per_codigo_ref8 like '" & IIf(FormatoD0(cmbN8.BoundText) = -1, "%", cmbN8.BoundText) & "'" & _
                     " AND persona.per_codigo_ref9 like '" & IIf(FormatoD0(cmbN9.BoundText) = -1, "%", cmbN9.BoundText) & "'" & _
                     " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo AND pai_codigo like '" & IIf(FormatoD0(cmbPais.BoundText) = -1, "%", cmbPais.BoundText) & "' "
    ElseIf optCalendario.Value = True Then
'        strFiltro = " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
'                    " INNER JOIN dia_facturacion df1 ON persona.emp_codigo=df1.emp_codigo " & _
'                    " AND persona.tip_ped_codigo=df1.tip_ped_codigo " & _
'                    " AND ciudad.pai_codigo=df1.pai_codigo " & _
'                    " AND df1.ciu_codigo='' " & _
'                    " AND df1.dia_fac_dia=DATEPART(dw,CURRENT_TIMESTAMP) "
                    
        
        strFiltro = " INNER JOIN (" & _
                    " SELECT DISTINCT emp_codigo,tip_ped_codigo,ciu_codigo " & _
                    " FROM ( " & _
                    " SELECT df1.emp_codigo,df1.tip_ped_codigo,df1.ciu_codigo " & _
                    " FROM dia_facturacion df1 " & _
                    " WHERE df1.pai_codigo='' " & _
                    " AND df1.dia_fac_dia=DATEPART(dw,CURRENT_TIMESTAMP) " & _
                    " AND df1.emp_codigo='" & strEmpresa & "' " & _
                    " AND df1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                    " UNION " & _
                    " SELECT df2.emp_codigo,df2.tip_ped_codigo,ciudad.ciu_codigo " & _
                    " FROM dia_facturacion df2 INNER JOIN ciudad ON df2.pai_codigo=ciudad.pai_codigo " & _
                    " WHERE df2.ciu_codigo='' " & _
                    " AND df2.dia_fac_dia=DATEPART(dw,CURRENT_TIMESTAMP) " & _
                    " AND df2.emp_codigo='" & strEmpresa & "' " & _
                    " AND df2.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                    " AND ciudad.ciu_codigo NOT IN " & _
                    " (SELECT df3.ciu_codigo " & _
                    " FROM dia_facturacion df3 " & _
                    " WHERE df3.pai_codigo='' " & _
                    " AND df3.dia_fac_dia!=DATEPART(dw,CURRENT_TIMESTAMP) " & _
                    " AND df3.emp_codigo='" & strEmpresa & "' " & _
                    " AND df3.tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        strFiltro = strFiltro & " )) c ) cc ON persona.emp_codigo=cc.emp_codigo " & _
                    " AND persona.tip_ped_codigo=cc.tip_ped_codigo " & _
                    " AND persona.ciu_codigo=cc.ciu_codigo "
    ElseIf optCalendarioAyer.Value = True Then
        
        strFiltro = " INNER JOIN (" & _
                    " SELECT DISTINCT emp_codigo,tip_ped_codigo,ciu_codigo " & _
                    " FROM ( " & _
                    " SELECT df1.emp_codigo,df1.tip_ped_codigo,df1.ciu_codigo " & _
                    " FROM dia_facturacion df1 " & _
                    " WHERE df1.pai_codigo='' " & _
                    " AND df1.dia_fac_dia=DATEPART(dw,CURRENT_TIMESTAMP)-1 " & _
                    " AND df1.emp_codigo='" & strEmpresa & "' " & _
                    " AND df1.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                    " UNION " & _
                    " SELECT df2.emp_codigo,df2.tip_ped_codigo,ciudad.ciu_codigo " & _
                    " FROM dia_facturacion df2 INNER JOIN ciudad ON df2.pai_codigo=ciudad.pai_codigo " & _
                    " WHERE df2.ciu_codigo='' " & _
                    " AND df2.dia_fac_dia=DATEPART(dw,CURRENT_TIMESTAMP)-1 " & _
                    " AND df2.emp_codigo='" & strEmpresa & "' " & _
                    " AND df2.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                    " AND ciudad.ciu_codigo NOT IN " & _
                    " (SELECT df3.ciu_codigo " & _
                    " FROM dia_facturacion df3 " & _
                    " WHERE df3.pai_codigo='' " & _
                    " AND df3.dia_fac_dia!=DATEPART(dw,CURRENT_TIMESTAMP)-1 " & _
                    " AND df3.emp_codigo='" & strEmpresa & "' " & _
                    " AND df3.tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        strFiltro = strFiltro & " )) c ) cc ON persona.emp_codigo=cc.emp_codigo " & _
                    " AND persona.tip_ped_codigo=cc.tip_ped_codigo " & _
                    " AND persona.ciu_codigo=cc.ciu_codigo "
    ElseIf optAbierto.Value = True Then
        strFiltro = ""
    End If
    If cmbNegocio.BoundText <> "" And optListaPedido.Value = True Then
            'Consulta todos los pedidos que pasan a bodega para ser revisados
            If chkCIRUC.Value = 1 Then
                strCli = "CONCAT(per_ruc,' ',per_apellido,' ',per_nombre)"
            Else
                strCli = "CONCAT(per_apellido,' ',per_nombre)"
            End If
            strSql = " SELECT DISTINCT RIGHT(ped_codigo,7)+0 as c, ped_fecha, " & strCli & " as nombC, " & _
                     " ped_observacion, ped_estado, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
                     " IIF(pedido.ven_codigo='' OR pedido.ven_codigo is null,persona.ven_codigo,pedido.ven_codigo) as ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
                     " per_sec_publico,per_siniva,per_fac_flete,IIF(pedido.ped_egr_bodega=0,per_dcto,pedido.ped_egr_bodega),pedido.tar_cre_codigo,ped_codigo,IIF(per_bloqueado_g+per_bloqueado>=1,1,0) as per_bloqueado,lis_pre_codigo,ped_dctoadicional " & _
                     " FROM (pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado) " & _
                     " INNER JOIN persona ON (pedido.per_codigo = persona.per_codigo) " & _
                     " AND (pedido.emp_codigo = persona.emp_codigo) " & _
                     " INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                     " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                     " AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                     strFiltro & _
                     " INNER JOIN tipo_factura ON (pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo) " & _
                     " LEFT JOIN tarjeta_credito ON (pedido.emp_codigo = tarjeta_credito.emp_codigo) AND (pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo) " & _
                     " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                     " AND " & IIf(chkContado.Value = 1, "persona.for_pag_codigo IN ('CONT','EFEC') AND ped_estado=1 ", "((IIF(persona.for_pag_codigo NOT IN ('CONT','EFEC'),0,99)=ped_estado OR ped_estado=1)) ") & _
                     " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' and ped_fechamod>='2015-07-10'" & _
                     " AND ped_fechamod between '" & dtpFechaIni.Value & "' and '" & dtpFechaFin.Value & "' " & _
                     " "
            clsPedidos.Ejecutar (strSql)
    Else
        Exit Sub
    End If
'    clsPedidos.Actualizar
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
    VSFGTot.TextMatrix(3, 1) = "0/" & VSFGPeds.Rows - 1
    
    strSql = " SELECT est_codigo,est_descripcion " & _
             " FROM est_pedido " & _
             " ORDER BY est_codigo"
    clsSql.Ejecutar strSql
    
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
    
End Sub

Private Sub ActualizarPrecios()
    Dim i As Long
    Dim clsCambioPrecio As New clsConsulta
    Dim dctoMax As Double
    'Verifica cuado se presionó un enter para devolver un tab
        If VSFGPeds.Rows > 1 And VSFG.Rows > 1 Then
            clsCambioPrecio.Inicializar AdoConn, AdoConnMaster
            For i = 1 To VSFG.Rows - 1
                If Left(VSFG.TextMatrix(i, 1), 3) <> "PR-" Then
                    strSql = " SELECT producto.prd_codigo, " & _
                             " (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio " & _
                             " FROM producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                             " AND producto.emp_codigo=lista_precio_p.emp_codigo" & _
                             " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                             " AND lista_precio_p.lis_pre_codigo=" & VSFGPeds.TextMatrix(VSFGPeds.Row, 20) & " "
                    clsCambioPrecio.Ejecutar strSql
                    If clsCambioPrecio.adorec_Def.RecordCount > 0 Then
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_precio='" & clsCambioPrecio.adorec_Def("lis_pre_p_precio") & "' " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                                 " AND ped_codigo=" & LblPedido.Caption & " "
                        clsCambioPrecio.Ejecutar strSql, "M"
                    End If
                    
                    dctoMax = 0
                    dctoMax = FormatoD2(txtDescuento.Text)
                    strSql = " SELECT prd_pro_porcentaje " & _
                             " FROM producto_promo " & _
                             " WHERE emp_codigo = '" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                             " AND LEFT('" & VSFGPeds.TextMatrix(VSFGPeds.Row, 1) & "',10) BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                    clsCambioPrecio.Ejecutar strSql
                    If clsCambioPrecio.adorec_Def.RecordCount > 0 Then
                        If FormatoD2(clsCambioPrecio.adorec_Def(0)) > FormatoD2(TxtDesc.Text) Then
                            dctoMax = FormatoD2(clsCambioPrecio.adorec_Def(0))
                            
                            strSql = " UPDATE det_pedido " & _
                                     " SET det_ped_dcto=ROUND(det_ped_cant_entregada*det_ped_precio*" & dctoMax & "/100.00,4) " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                                     " AND ped_codigo=" & LblPedido.Caption & " "
                            clsCambioPrecio.Ejecutar strSql, "M"
                        End If
                    End If
                End If
            Next i
            VSFGPeds_DblClick
        End If

End Sub

Private Sub CmdConfirmarTodo_Click()
    Dim i As Long
    Dim clsBloq As New clsConsulta
    clsBloq.Inicializar AdoConn, AdoConnMaster
    strListaFactura = ""
    strListaPedido = ""
    lngNFacNPed = 0
    VSFGTot.TextMatrix(1, 1) = 0
    VSFGTot.TextMatrix(2, 1) = 0
    VSFGTot.TextMatrix(3, 1) = "0/0"
    pbDetalle.Max = VSFGPeds.Rows - 1
    For i = 1 To VSFGPeds.Rows - 1
        VSFGPeds.Row = i
        VSFGPeds.Col = 1
        VSFGPeds.Select i, 1
        VSFGPeds.ShowCell i, 1
        VSFGPeds_DblClick
        ActualizarPrecios
        If VSFGPeds.TextMatrix(VSFGPeds.Row, 19) = 0 Then
            If FormatoD2(TxtSubTotal.Text) >= 80 Then
                
                pp = 0
                For tt = 1 To VSFG.Rows - 1
                    If Left(VSFG.TextMatrix(tt, 1), 1) = "N" Then
                        pp = 1
                    End If
                    
                Next tt
                If pp = 0 Then
                    'MsgBox "AA"
                End If
            End If
            CmdConfirmar_Click
            VSFGTot.TextMatrix(1, 1) = VSFGTot.TextMatrix(1, 1) + 1
        Else
            strSql = " UPDATE pedido SET ped_estado=4, " & _
                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
                     " ped_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & VSFGPeds.TextMatrix(VSFGPeds.Row, 18)
            clsBloq.Ejecutar (strSql), "M"
            VSFGTot.TextMatrix(2, 1) = VSFGTot.TextMatrix(2, 1) + 1
        End If
        pbDetalle.Value = i
        pbDetalle.Refresh
        VSFGTot.TextMatrix(3, 1) = i & "/" & VSFGPeds.Rows - 1
        VSFGTot.Refresh
    Next i
    
''    Dim RepFactura As New frmReporte
''    Dim RepStk As New frmReporte
''    strListaFactura = Left(strListaFactura, Len(strListaFactura) - 1)
''    strListaPedido = Left(strListaPedido, Len(strListaPedido) - 1)
''
''
''    If ImpresoraEtiqueta = "" Then
''        RepStk.VSPrint.PrintDialog pdPrint
''        ImpresoraEtiqueta = RepStk.VSPrint.Device
''        GuardarImpresoras
''    End If
''    RepStk.VSPrint.Device = ImpresoraEtiqueta
''
''    RepStk.VSPrint.PaperWidth = 7669.292
''    RepStk.VSPrint.PaperHeight = 3885.039
''    RepStk.strNumero = strListaPedido
''    RepStk.strReporte = "rptSTKDespacho"
''    RepStk.strTipo = 5
''    RepStk.Show
''    RepStk.Form_Activate
''    RepStk.VSPrint.Copies = 1
''    RepStk.VSPrint.PrintDoc
''    MsgBox "Fin etiqueta"
''    'Unload RepStk
''
''    RepFactura.strNumero = strListaFactura
''    RepFactura.strReporte = "rptFactura"
''    RepFactura.Show
''    RepFactura.Form_Activate
''    RepFactura.VSPrint.Copies = 2
''    RepFactura.VSPrint.Collate = colFalse
''    'RepFactura.VSPrint.PrintDoc
''    MsgBox "Fin Factura"
''    'Unload RepFactura
''
    
End Sub

Private Sub cmdEjecutar_Click()
    pbGeneral.Value = 1
    pbGeneral.Refresh
    lblEstado.Caption = "Promo x Combo"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optPromoCombo.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    pbGeneral.Value = 2
    pbGeneral.Refresh
    lblEstado.Caption = "Promo Combo Pedido"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optPromoComboPedido.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    pbGeneral.Value = 3
    pbGeneral.Refresh
    lblEstado.Caption = "Dcto x Combo"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optDctoCombo.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    pbGeneral.Value = 4
    pbGeneral.Refresh
    lblEstado.Caption = "n Prendas a $x.xx"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optNPrendasAY.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    
    dtpFechaFinAplicar.Value = dtpEjecutar.Value
    pbGeneral.Value = 5
    pbGeneral.Refresh
    lblEstado.Caption = "Premio por monto Marca"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optPromoPremioPorMontoMarca.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    
    dtpFechaFinAplicar.Value = dtpEjecutar.Value
    pbGeneral.Value = 6
    pbGeneral.Refresh
    lblEstado.Caption = "Dcto por Fecha"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optDctoFecha.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    
    pbGeneral.Value = 7
    pbGeneral.Refresh
    lblEstado.Caption = "Premio x Monto"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optPremio.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    pbGeneral.Value = 8
    pbGeneral.Refresh
    lblEstado.Caption = "Cargando Incentivos"
    lblEstado.Refresh
    frmIncentivos.Show
    frmIncentivos.cmbNegocioAplicar.BoundText = cmbNegocio.BoundText
    frmIncentivos.dtpFechaInicioAplicar.Value = dtpFechaIni.Value
    frmIncentivos.dtpFechaFinAplicar.Value = dtpFechaFin.Value
    frmIncentivos.optIncentivo.Value = True
    frmIncentivos.cmdActualizar_Click
    frmIncentivos.cmdAplicarAplicar_Click
    

    pbGeneral.Value = 9
    pbGeneral.Refresh
    Unload frmIncentivos
    lblEstado.Caption = "Facturando..."
    lblEstado.Refresh
    HoyDia = Hoy
    dtpFecha.Value = HoyDia
    
    cmdActualizar_Click
    CmdConfirmarTodo_Click
    pbGeneral.Value = 10
    pbGeneral.Refresh
    lblEstado.Caption = "Enviando Correos..."
    lblEstado.Refresh
    frmV_ReImprimirFacPed.Show
    frmV_ReImprimirFacPed.cmbNegocio.BoundText = cmbNegocio.BoundText
    frmV_ReImprimirFacPed.chkFechas = 1
    frmV_ReImprimirFacPed.Fecha1.Value = dtpFechaFinAplicar.Value
    frmV_ReImprimirFacPed.Fecha2.Value = Now
    frmV_ReImprimirFacPed.dtpFechaFactura.Value = dtpFecha.Value
    frmV_ReImprimirFacPed.cmdActualizar_Click
    frmV_ReImprimirFacPed.cmdEnvioCorreo_Click
    pbGeneral.Value = 11
    pbGeneral.Refresh
    lblEstado.Caption = "Termino con Exito"
    lblEstado.Refresh
    MsgBox "Facturación finalizada"
End Sub

Private Sub cmdImprimirBloque_Click()
    frmV_ReImprimirFacPed.Show
End Sub

Private Sub cmdProgramar_Click()
    TmrAct.Enabled = True
    lblEstado.Caption = "Tiempo restante a ejecutar " & DateDiff("n", Now, dtpEjecutar.Value) & " minutos"
    lblEstado.Refresh
    'cmdEjecutar_Click
End Sub

Private Sub Command1_Click()
'    ActualizaPrecio
    Dim pednopago As String
    Dim facnopago As String
    pednopago = InputBox("Pedido")
    facnopago = InputBox("Factura") & ","
    PagarFacturaDePedidoPagado pednopago, facnopago, ""
    
End Sub

Private Sub Command2_Click()
    Dim clsPed As New clsPedidos
    Dim clsCons As New clsConsulta
    Dim clsCons2 As New clsConsulta
    Dim clsCambioPrecio As New clsConsulta
    clsPed.Inicializar AdoConn, AdoConnMaster
    clsCons.Inicializar AdoConn, AdoConnMaster
    clsCons2.Inicializar AdoConn, AdoConnMaster
    clsCambioPrecio.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ped_codigo,ped_fechamod " & _
             " FROM pedido " & _
             " where pedido.emp_codigo='RYB' AND ped_estado in (0,1)" & _
             " "
    clsCons.Ejecutar strSql
    MsgBox clsCons.adorec_Def.RecordCount & " Pedidos", vbInformation
    While Not clsCons.adorec_Def.EOF
        clsPed.RecalculoTotal clsCons.adorec_Def(0)
    clsCons.adorec_Def.MoveNext
    Wend
    MsgBox "FIN"
End Sub

Private Sub dtpFecha_LostFocus()
    If DateDiff("d", HoyDia, dtpFecha.Value) > 1 Or DateDiff("d", HoyDia, dtpFecha.Value) < 0 Then
        frmClave.strClaveMAESTRA = strClaveMAESTRA
        frmClave.dblPrecio = "Fecha"
        frmClave.Show vbModal
        If frmClave.Ret = False Then
            dtpFecha.Value = HoyDia
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidos = Nothing
    Set clsPed = Nothing
    Set clsSql = Nothing
    Set clsTFac = Nothing
    Set clsRecargos = Nothing
    Set clsFPago = Nothing
    Set clsFormaPago = Nothing
    Set clsTC = Nothing
    Set clsRet = Nothing
    Set clsExis = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    With VSFGReca
        For i = 1 To (.Rows - 1)
            .TextMatrix(i, 0) = i
            If conBot = True Then
                'Coloca los botones de elimniar fila en el grid
                .Cell(flexcpPicture, i, 0) = imgBtnUp
                .Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub

Private Sub CalcuReca()
    'Calcula el total del pedido
    Dim Suma As Double
    For i = 1 To VSFGReca.Rows - 1
        Suma = Suma + FormatoD2(VSFGReca.TextMatrix(i, 3))
    Next i
    TxtRecargo = FormatoD2(Suma)
End Sub

Private Sub CalcuTotal()
    'Calcula es total del pedido
    Dim Suma As Double
    Dim SumaIVA As Double
    Dim SumaDcto As Double
    Dim sumaSDcto As Double
    Dim SumaIVASDcto As Double
    
    CalcuReca
    Suma = 0
    SumaIVA = 0
    sumaSDcto = 0
    SumaIVASDcto = 0
    SumaDcto = 0
    For i = 1 To VSFG.Rows - 1
        If Abs(FormatoD0(VSFG.TextMatrix(i, 9))) = 0 Then
            If Val(FormatoD2(VSFG.TextMatrix(i, 4))) <> 0 Then
                If Abs(FormatoD0(VSFG.TextMatrix(i, 8))) = 1 Then
                    Suma = Suma + FormatoD2(FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5))) - FormatoD2(VSFG.TextMatrix(i, 6)))
                    sumaSDcto = sumaSDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
                Else
                    SumaIVA = SumaIVA + FormatoD2(FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5))) - FormatoD2(VSFG.TextMatrix(i, 6)))
                    SumaIVASDcto = SumaIVASDcto + FormatoD2(FormatoD4(VSFG.TextMatrix(i, 4)) * FormatoD4(VSFG.TextMatrix(i, 5)))
                End If
            SumaDcto = SumaDcto + FormatoD2(VSFG.TextMatrix(i, 6))
            End If
        Else
            'MsgBox "DCT"
        End If
    Next i
    TxtRecargo.Tag = FormatoD2(TxtRecargo.Text)
    TxtRecargo.Text = FormatoD2(TxtRecargo.Text)
    'Coloca los totales parciales de la factura
    TxtDesc.Text = FormatoD2(SumaDcto)
    TxtSubTotal = FormatoD2(sumaSDcto) + FormatoD2(SumaIVASDcto)
    If SinIVA = False Then
        TxtIva = FormatoD2((Suma) * IVA / 100#)
    Else
        TxtIva = 0
    End If
    TxtTotal = FormatoD2(Suma + SumaIVA + TxtIva + Val(TxtRecargo.Tag))
End Sub

Private Sub CalcuDesc()
    Dim strSql As String
    TxtDesc = 0
    strSql = " SELECT COALESCE(SUM(ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)),0) as suman," & _
             " SUM(ROUND(ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2)) as Descu " & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C' " & _
             " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
             " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
             " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
             " AND CAST(pedido.ped_fechamod as date) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
             " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
             " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' AND det_pedido.det_ped_incentivo=0 AND pedido.ped_codigo='" & FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 18)) & "' GROUP BY pedido.ped_codigo"
    clsPed.Ejecutar (strSql)
    If clsPed.adorec_Def.RecordCount > 0 Then
        TxtDesc.Text = FormatoD2(clsPed.adorec_Def("Descu"))
    Else
        TxtDesc.Text = "0.00"
    End If
End Sub


'Función que verifica si es necesario realizar un backOrder del pedido
Private Function verifBackOr() As Integer
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, 3)) <> Val(VSFG.TextMatrix(i, 4)) Then
            verifBackOr = 1
            Exit For
        End If
    Next i
End Function

'Función que genera un backOrder de un pedido
Private Sub backOrder(codPed As Double, CodigoPer As String, codEmp As String)
    Dim clsBack As New clsConsulta
    clsBack.Inicializar AdoConn, AdoConnMaster
'    'Recupera el código con el cual se debe generar un nuevo backOrder
'    strSql = " Select COALESCE(max(bac_codigo),0) as num " & _
'             " From backorder " & _
'             " Where emp_codigo='" & codEmp & "'" & _
'             " GROUP BY emp_codigo"
'    clsBack.Ejecutar (strSql)
'    Dim codBac As Double
'    codBac = clsBack.adorec_Def("num") + 1
'    'Inserta la cabecera del backOrder
'    strSql = " INSERT INTO backorder " & _
'             " SELECT " & codBac & " AS bac_codigo, emp_codigo, ped_codigo, CURRENT_TIMESTAMP AS bac_fecha, " & _
'             " 0 AS bac_baja, CURRENT_TIMESTAMP AS bac_fechamod, '" & strUsuario & "' AS bac_usumod " & _
'             " From pedido " & _
'             " WHERE ped_codigo=" & codPed & " AND emp_codigo='" & codEmp & "' "
'    clsBack.Ejecutar (strSql), "M"
'    'Inserta los detalles del backOrder
'    strSql = " INSERT INTO det_backorder " & _
'             " SELECT emp_codigo, prd_codigo, " & codBac & " AS bac_codigo, " & _
'             " det_ped_cant_pedida-det_ped_cant_entregada AS det_bac_cantidad, " & _
'             " det_ped_precio, CURRENT_TIMESTAMP AS det_bac_fechamod, " & _
'             " '" & strUsuario & "' AS det_bac_usumod " & _
'             " From det_pedido " & _
'             " WHERE emp_codigo='" & codEmp & "' " & _
'             " AND det_ped_cant_pedida > det_ped_cant_entregada " & _
'             " AND ped_codigo= " & codPed & _
'             " Order by prd_codigo "


'****** ACTUALIZAR TIPO PEDIDO
    
    Dim Fact As String
    Dim PedNum As String
    strSql = " SELECT tip_ped_ptofac,tip_ped_factura_directo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    clsBack.Ejecutar strSql
    If clsBack.adorec_Def.RecordCount > 0 Then
        Fact = clsBack.adorec_Def(0)
    End If
    strSql = " BEGIN TRAN"
    clsBack.Ejecutar (strSql), "M"
    strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(strSucursal & Fact & "0000001") & "') as num " & _
             " From pedido WITH (TABLOCKX) " & _
             " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(strSucursal & Fact) & "%'" & _
             " GROUP BY emp_codigo"
    clsBack.Ejecutar (strSql), "M"
    PedNum = clsBack.adorec_Def("num")
    strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
             " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega, ped_fechamod, ped_usumod) " & _
             " VALUES ('" & strEmpresa & "'," & PedNum & ",'" & CodigoPer & "','', " & _
             " 'SINTC','0'," & _
             " CURRENT_TIMESTAMP,'-2',0,'', " & _
             " '" & codPed & "',1,'" & codPed & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
    clsBack.Ejecutar (strSql), "M"
    
    strSql = " COMMIT TRAN"
    clsBack.Ejecutar (strSql), "M"
                
'                strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
'                             " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
'                             " VALUES ('" & strEmpresa & "'," & num & ",'" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 1) & "'," & .TextMatrix(i, 4) & ",0 " & _
'                             "," & .TextMatrix(i, 5) & "," & .TextMatrix(i, 6) & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
'                    clsSql.Ejecutar (strSql), "M"


    strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
             " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
             " SELECT 'RYB','" & PedNum & "','PR-CARGOO100330TU','PRI', " & _
             " 1,1, " & _
             " 1.34,0, CURRENT_TIMESTAMP, " & _
             " '" & strUsuario & "'" & _
             " UNION " & _
             " SELECT emp_codigo, '" & PedNum & "',prd_codigo,'PRI', " & _
             " det_ped_cant_pedida-det_ped_cant_entregada,0, " & _
             " det_ped_precio,det_ped_dcto, CURRENT_TIMESTAMP, " & _
             " '" & strUsuario & "' " & _
             " From det_pedido " & _
             " WHERE emp_codigo='" & codEmp & "' " & _
             " AND det_ped_cant_pedida > det_ped_cant_entregada " & _
             " AND ped_codigo= " & codPed & _
             " "

    clsBack.Ejecutar (strSql), "M"
    Set clsBack = Nothing
End Sub

Private Sub cmdActualizar_Click()
    cmbNegocio_Change
    
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
    Dim Resp As Integer
    Dim codEgr As Double
    Dim Hay As Integer, FechaSalida, CantDisponible, CantSacada
    Dim cantAFac As Double
    Dim cta_cobrar As String
    Dim tip_egr_ctaconta As String
    Dim cue_p_c_codigo As String
    Dim strSql As String
    Dim strNumero As String
    Dim booPasar As Boolean
    Dim booGuardar As Boolean
    Dim clsAuxxx As New clsConsulta
    Dim clsAsiento As New clsContable
    Dim clsEgreso As New clsInventario
    Dim clsPedRep As New clsPedidos
    Dim PedReprogramado As String
    Dim clsCta As New clsCtaXx
    Dim Nfac As Integer
    Dim INICIO As Integer
    Dim strObs As String
    Dim egr() As String
    Dim egrTot() As Double
    Dim strFormaCobro As String
    'clsSql.
    clsPedRep.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    frmFormaCobro.strNegocio = Me.cmbNegocio.BoundText
    frmFormaCobro.strCliente = CodPer
    If frmFormaCobro.strFormaCobro = "" Then
        frmFormaCobro.Show vbModal
    End If
    If frmFormaCobro.strFormaCobro <> "" Then
        strFormaCobro = frmFormaCobro.strFormaCobro
        clsEgreso.Inicializar AdoConn, AdoConnMaster
        clsAuxxx.Inicializar AdoConn, AdoConnMaster
        strObs = TxtObserv.Text
        'Verifica que se haya seleccionado un tipo de factura
        If CmbTipoFac = "" Or CmbTipoFac.MatchedWithList = False Then
            MsgBox "Seleccione un tipo de factura por favor.", vbInformation, "Factura"
            CmbTipoFac.SetFocus
            Exit Sub
        End If
        'Verifica que se haya seleccionado un tipo de forma de pago
        If CmbFpago.Text = "" Or CmbFpago.MatchedWithList = False Then
            MsgBox "Seleccione un tipo de forma de pago por favor.", vbInformation, "Forma de Pago"
            CmbFpago.SetFocus
            Exit Sub
        Else
            ''''MsgBox CmbFpago.BoundText
        End If
        'Detiene la actualización automática de pedidos por parte del control timer
        'TmrAct.Enabled = False
    '****** BACKORDER
        'Verifica si no se completa el pedido
        'Verifica si se quiere realizar un backorder de los productos no completados en el pedido
        If verifBackOr() = 1 Then
            Resp = vbYes
        End If
        'Verifica si la respuesta fue afirmativa para hacer un backOrder
        If Resp = vbYes And (Me.cmbNegocio.BoundText <> "JON" And Me.cmbNegocio.BoundText <> "LEM") Then
            'Llama al procedimiento que genera un backorder
            backOrder LblPedido, CodPer, strEmpresa
        End If
    '****** EGRESO
       Dim Fact As String
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            Fact = clsSql.adorec_Def(0)
        End If
    
        strSql = " SELECT tip_ped_item " & _
                 " FROM tipo_pedido " & _
                 " Where emp_codigo='" & strEmpresa & "' AND tip_ped_ptofac='" & Fact & "' "
        clsTC.Ejecutar (strSql)
        If clsTC.adorec_Def.RecordCount > 0 Then
            maxItem = FormatoD0(clsTC.adorec_Def(0))
        Else
            maxItem = PrdEntregar
        End If
        'Genera un egreso de mercadería
        'Obtiene el código con el que se debe insertar el nuevo egreso de productos
        MousePointer = 11
        
        Nfac = Round(PrdEntregar / maxItem + 0.4999)
        INICIO = 0
        If Nfac = 0 And Val(TxtRecargo) <> 0 Then
            Nfac = 1
        End If
        
        
        strSql = " SELECT * FROM pedido WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo='" & LblPedido & "' "
        clsAuxxx.Ejecutar strSql
        If FormatoD0(clsAuxxx.adorec_Def("ped_estado")) = 8 Or FormatoD0(clsAuxxx.adorec_Def("ped_estado")) = 2 Then
            'MsgBox "El pedido " & LblPedido & " ya fue facturado hace unos instantes en la factura " & clsAuxxx.adorec_Def("ped_egr_codigo"), vbInformation, "Pedido Facturado"
            Exit Sub
        ElseIf FormatoD0(clsAuxxx.adorec_Def("ped_estado")) <> 0 And FormatoD0(clsAuxxx.adorec_Def("ped_estado")) <> 1 Then
            'MsgBox "El pedido " & LblPedido & " no puede ser facturado esta bloqueado o anulado , vbInformation, "Pedido Facturado"
            Exit Sub
        End If
        
        'MsgBox "Usted necesitará" & vbNewLine & Nfac & vbNewLine & "Factura(s)", vbInformation, "NUMERO DE FACTURAS"'
        Dim Recs As Double
        ReDim egr(Nfac)
        ReDim egrTot(Nfac)
        For INICIO = 1 To Nfac
            If INICIO = 1 Then
                Recs = FormatoD4(TxtRecargo.Text)
            Else
                Recs = 0
            End If
            If LblPedido = "10021016506" Or LblPedido = "10021016512" Or LblPedido = "10021016525" Or LblPedido = "10021016533" Or LblPedido = "10021016537" Or LblPedido = "10021016539" Or LblPedido = "10021016550" Or LblPedido = "10021016551" Or LblPedido = "10021016565" Or LblPedido = "40020198482" Or LblPedido = "40020198485" Then
                'MsgBox LblPedido
            End If
            strSql = " EXEC Sp_Drop_Table_if_Exist '#pedauxs' "
            clsAuxxx.Ejecutar strSql
            strSql = " CREATE TABLE #pedauxs (row integer NOT NULL,codigo varchar(40) NOT NULL, subtotal decimal(14,4), dcto decimal(14,4), subtotal2 decimal(14,4), PRIMARY KEY(codigo)) "
            clsAuxxx.Ejecutar strSql
            strSql = " INSERT INTO #pedauxs " & _
                     " SELECT * FROM (SELECT ROW_NUMBER() OVER (ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo) as row,det_pedido.prd_codigo as codigo, " & _
                     " (ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)) as subtotal," & _
                     " ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)>ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00),2) " & _
                     " + (ROUND(((" & strCantidadTipoPedido & "*det_ped_precio)-ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)>ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00),2))*(pedido.ped_dctoadicional/100.00),2)) as dcto, " & _
                     " (ROUND(((" & strCantidadTipoPedido & "*det_ped_precio)-ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)>ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & "*det_ped_precio,2)*COALESCE(IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                     FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & "),0)/100.00),2))*(1-pedido.ped_dctoadicional/100.00),2)) as subtotal2 " & _
                     " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                     " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo " & _
                     " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                     " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                     " AND producto_promo.prd_pro_fechaini<=CAST(pedido.ped_fechamod as date) AND producto_promo.prd_pro_fechafin>=CAST(pedido.ped_fechamod as date) AND producto_promo.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                     " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                     " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                     " WHERE pedido.emp_codigo='" & strEmpresa & "' AND " & strCantidadTipoPedido & ">0 " & _
                     " AND det_pedido.det_ped_incentivo=0 AND pedido.ped_codigo=" & LblPedido & _
                     " ) a WHERE row > " & ((INICIO - 1) * maxItem) & " and row <= " & (INICIO * maxItem) & "  "
            clsAuxxx.Ejecutar strSql
                
            If INICIO = 1 Then
                If VSFG.TextMatrix(VSFG.Rows - 1, 1) = tipoPagoDetalle Then
                    strSql = " INSERT INTO #pedauxs VALUES('" & VSFG.TextMatrix(VSFG.Rows - 1, 1) & "','" & VSFG.TextMatrix(VSFG.Rows - 1, 5) & "',0,'" & VSFG.TextMatrix(VSFG.Rows - 1, 5) & "')"
                    clsAuxxx.Ejecutar strSql
                End If
            End If
            
            If SinIVA = False Then
                strSql = " select COALESCE(SUM(subtotal),0),COALESCE(SUM(dcto),0),ROUND((COALESCE(SUM(IIF(prd_iva=1,subtotal,0)),0)-COALESCE(SUM(IIF(prd_iva=1,dcto,0)),0))*" & IVA & "/100.00,2),COALESCE(SUM(subtotal),0)-COALESCE(SUM(dcto),0)+ROUND((COALESCE(SUM(IIF(prd_iva=1,subtotal,0)),0)-COALESCE(SUM(IIF(prd_iva=1,dcto,0)),0))*" & IVA & "/100.00,2)+ROUND(" & Recs & ",2),ROUND(" & Recs & ",2),COUNT(*) as numItem " & _
                         " FROM #pedauxs inner join producto on #pedauxs.codigo=producto.prd_codigo AND producto.emp_codigo='" & strEmpresa & "'"
            Else
                strSql = " select COALESCE(SUM(subtotal),0),COALESCE(SUM(dcto),0),ROUND((COALESCE(SUM(subtotal),0)-COALESCE(SUM(dcto),0))*0/100.00,2),COALESCE(SUM(subtotal),0)-COALESCE(SUM(dcto),0)+ROUND((COALESCE(SUM(subtotal),0)-COALESCE(SUM(dcto),0))*0/100.00,2)+ROUND(" & Recs & ",2),ROUND(" & Recs & ",2),COUNT(*) as numItem " & _
                         " FROM #pedauxs "
            End If
            
            clsAuxxx.Ejecutar strSql
            booGuardar = False
            TxtObserv = strObs & " - Ped.:" & LblPedido & " - Fact:" & INICIO & "/" & Nfac
            If FormatoD0(clsAuxxx.adorec_Def("numItem")) > 0 Then
                While booGuardar = False
                    booGuardar = clsEgreso.NuevoEgr(True, "FAC", True, , Fact, , CmbFpago.BoundText, CodPer, Format(dtpFecha.Value, "yyyy-MM-dd"), , codVen, TxtObserv, , strAutorFactura, strCaducaFactura, FormatoD2(clsAuxxx.adorec_Def(0)), FormatoD2(clsAuxxx.adorec_Def(4)), FormatoD2(clsAuxxx.adorec_Def(1)), FormatoD2(clsAuxxx.adorec_Def(2)), FormatoD2(clsAuxxx.adorec_Def(3)), 0, SecPublico, SinIVA, CodigoIVA, strFormaCobro)
                    egrTot(INICIO) = FormatoD2(clsAuxxx.adorec_Def(3))
                    If INICIO = 1 And booGuardar = False Then
                        strSql = " EXEC Sp_Drop_Table_if_Exist '#pedauxs' "
                        clsAuxxx.Ejecutar strSql
                        strSql = " COMMIT TRAN "
                        clsAuxxx.Ejecutar strSql, "M"
                        MousePointer = 0
                        Exit Sub
                    ElseIf booGuardar = False And INICIO <> 1 Then
                        MsgBox "Inicio la facturacion de un pedido de" & vbNewLine & Nfac & " factura(s). " & vbNewLine & "Debe concluir con todas caso contrario el pedido quedará incompleto de facturar.", vbInformation, "FACTURACION DE PEDIDOS"
                    End If
                Wend
                
                strSql = " EXEC Sp_Drop_Table_if_Exist '#pedauxs' "
                clsAuxxx.Ejecutar strSql
                codEgr = clsEgreso.strDoc
                
                If booGuardar = True Then
                    codEgr = clsEgreso.strDoc
                    clsAsiento.Inicializar AdoConn, AdoConnMaster
                    egr(INICIO) = codEgr
                    clsAsiento.NuevoAsiento "F", dtpFecha.Value, 0, 0, TxtTotal.Text, "FACTURA " & codEgr
                    'Inserta la cabecera del egreso
                    clsEgreso.ModificaEgr , , , , , , clsAsiento.NumAsiento
                    'Inserta los detalles de egreso
                    strSql = " SELECT * FROM (SELECT ROW_NUMBER() OVER (ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo) as row,det_pedido.prd_codigo, dep_codigo, " & strCantidadTipoPedido & ", det_ped_precio,prd_costo, " & _
                             " ROUND(ROUND(IIF((ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))>" & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & _
                             FormatoD4(txtDescuento.Text) & ") / 100.00,(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))," & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00),2) " & _
                             " + ROUND(" & strCantidadTipoPedido & "*det_ped_precio - ROUND(IIF((ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))>" & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                             FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00,(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))," & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00),2),2)*(pedido.ped_dctoadicional/100.00),2) as det_ped_dcto," & _
                             " prd_iva,ROUND((ROUND(ROUND(IIF((ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))>" & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                             FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00,(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))," & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00),2) " & _
                             " + ROUND(" & strCantidadTipoPedido & "*det_ped_precio - ROUND(IIF((ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))>" & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & _
                             FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00,(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2))," & strCantidadTipoPedido & "*det_ped_precio * IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(txtDescuento.Text) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(txtDescuento.Text) & ") / 100.00),2),2)*(pedido.ped_dctoadicional/100.00),2))" & _
                             " /(" & strCantidadTipoPedido & "*det_ped_precio)*100,2) as pdcto " & _
                             " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                             " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                             " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                             " AND producto_promo.prd_pro_fechaini<=CAST(pedido.ped_fechamod as date) AND producto_promo.prd_pro_fechafin>=CAST(pedido.ped_fechamod as date) AND producto_promo.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                             " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                             " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                             " WHERE pedido.emp_codigo='" & strEmpresa & "' AND " & strCantidadTipoPedido & ">0 " & _
                             " AND det_pedido.det_ped_incentivo=0 AND det_pedido.det_ped_incentivo=0 AND " & strCantidadTipoPedido & ">0 AND pedido.ped_codigo=" & LblPedido & _
                             " ) a WHERE row > " & ((INICIO - 1) * maxItem) & " and row <= " & (INICIO * maxItem) & "  "
                    clsSql.Ejecutar (strSql)
                    
                    While Not clsSql.adorec_Def.EOF
                        'If Left(clsSql.adorec_Def("prd_codigo"), 1) = "N" Then
                        clsEgreso.NuevoDetEgr clsSql.adorec_Def("prd_codigo"), clsSql.adorec_Def("dep_codigo"), clsSql.adorec_Def("" & strCantidadTipoPedido & ""), clsSql.adorec_Def("det_ped_precio"), clsSql.adorec_Def("prd_costo"), clsSql.adorec_Def("det_ped_dcto"), Abs(FormatoD0(clsSql.adorec_Def("prd_iva"))), clsSql.adorec_Def("pdcto")
                        'End If
                        clsSql.adorec_Def.MoveNext
                    Wend
                    
                    'Actualiza el estado del pedido a vendido
                    strSql = " UPDATE pedido SET ped_estado=8, " & _
                             " tipo_fac_codigo=" & CmbTipoFac.BoundText & ", " & _
                             " ped_tip_egr_codigo='FAC', " & _
                             " ped_egr_codigo=" & codEgr & ", " & _
                             " ped_subtotal='" & FormatoD2(TxtTotal.Text) & "', " & _
                             " " & _
                             " ped_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido
                    clsAuxxx.Ejecutar (strSql), "M"
                'TC
                    If INICIO = 1 Then
                        If VSFG.TextMatrix(VSFG.Rows - 1, 1) = tipoPagoDetalle Then
                            clsEgreso.NuevoDetEgr VSFG.TextMatrix(VSFG.Rows - 1, 1), VSFG.TextMatrix(VSFG.Rows - 1, 0), 1, VSFG.TextMatrix(VSFG.Rows - 1, 5), VSFG.TextMatrix(VSFG.Rows - 1, 5), 0, 1
                        End If
                    End If
                
                '****** RECARGOS
                    'Genera los posibles recargos que podujo esta factura
                    If INICIO = 1 Then
                        For i = 1 To VSFGReca.Rows - 1
                            If VSFGReca.TextMatrix(i, 1) <> "" Then
                                clsEgreso.NuevoDetEgrRecargo VSFGReca.TextMatrix(i, 1), FormatoD2(VSFGReca.TextMatrix(i, 3))
                            End If
                        Next i
                    End If
                    clsFPago.adorec_Def.MoveFirst
                    strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
                    'Inserta un nuevo registro de la cuenta por cobrar*/
                    clsCta.Inicializar AdoConn, AdoConnMaster
                    clsFPago.adorec_Def.Find strComparar
                    clsCta.NuevaCta "C", 1, "00", Format(dtpFecha.Value, "yyyy-MM-dd"), Format(DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha.Value), "yyyy-MM-dd"), CodPer, "Factura # " & codEgr & " - " & TxtObserv, strSucursal & Fact, Right(codEgr, 7), strAutorFactura, strCaducaFactura, clsEgreso.dblTotalProd, clsEgreso.dblTotalServ, clsEgreso.dblTotalProdIVA, clsEgreso.dblTotalServIVA, 2, clsEgreso.dblIVA, clsEgreso.dblSubTotal0, 0, 0, 0, clsEgreso.dblTotal, clsAsiento.NumAsiento
                    clsCta.IngAsientoEgr clsAsiento, clsEgreso
                    Set clsCta = Nothing
                    Set clsAsiento = Nothing
                    'MsgBox "Pedido No. " & LblPedido & " facturado." & vbNewLine & "Factuna No. " & codEgr, vbInformation, "Pedido"
                    DocElectronico "01", (codEgr)
                    CmdConfirmar.Enabled = False
                    'cmdPreFactura.Enabled = False
                    'CmdGuiaRemi.Enabled = False
                    'cmdReserva.Enabled = False
                End If
DatosGuia:
                If GuiaAutomatica = True Then
                    Unload frmDatosGuia
                    frmDatosGuia.booGuiaCreada = False
                    frmDatosGuia.strCliente = ""
                    frmDatosGuia.strTipoDocumento = ""
                    frmDatosGuia.strCliente = CodPer
                    frmDatosGuia.strTipoDocumento = "FAC"
                    frmDatosGuia.strNumeroDocumento = codEgr
                    
                    If frmDatosGuia.booGuiaCreada = False Then
                        frmDatosGuia.Show vbModal
                    End If
                    If frmDatosGuia.strCourier <> "" And frmDatosGuia.strPlaca <> "" Then
                        clsEgreso.CrearGuia frmDatosGuia.strCourier, frmDatosGuia.strPlaca, strSucursal, Fact
                    Else
                        If MsgBox("Este egreso sale sin Guia de Remision? ", vbQuestion + vbYesNo, "Guia Remision") = vbNo Then
                            GoTo DatosGuia
                        End If
                    End If
                End If
            End If
            
        Next INICIO
        strListaFacturaPed = ""
        For i = Nfac To 1 Step -1
            strListaFacturaPed = strListaFacturaPed & egr(i) & ","
        Next i
        
        PedReprogramado = clsPedRep.GenerarReprogramacion(LblPedido)
        
        PagarFacturaDePedidoPagado LblPedido, strListaFacturaPed, PedReprogramado
    
        'Actualiza el estado del pedido a vendido
        strSql = " UPDATE pedido SET ped_estado=8, " & _
                 " tipo_fac_codigo=" & CmbTipoFac.BoundText & ", " & _
                 " ped_tip_egr_codigo='FAC', " & _
                 " ped_egr_codigo=" & codEgr & ", " & _
                 " ped_subtotal='" & FormatoD2(TxtTotal.Text) & "', " & _
                 " ped_fechamod=CURRENT_TIMESTAMP, " & _
                 " ped_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido
        clsSql.Ejecutar (strSql), "M"
    '****** GRID
        'Actualiza el grid que muestra los pedidos actuales
        'clsPedidos.Actualizar
        Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
    '****** RETENCIONES
        If clsEgreso.strDoc <> "" Then
            clsEgreso.DetRetenciones
        End If
    '****** COTIZACION
        'Actualiza el estado de la cotización relacionada a vendido en caso de que esta exista
        If codCot <> "" Then
            strSql = " UPDATE cotizacion SET cot_estado=2, " & _
                     " cot_fechamod=CURRENT_TIMESTAMP, " & _
                     " cot_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND cot_codigo='" & codCot & "' "
            clsSql.Ejecutar (strSql), "M"
        End If
        MousePointer = 0
        If booGuardar = True Then
            For i = Nfac To 1 Step -1
                strListaFactura = strListaFactura & egr(i) & ","
                strListaPedido = strListaPedido & Trim(LblPedido) & ","
                lngNFacNPed = lngNFacNPed + 1
            Next i
            CmdLimpiar = True
            dtpFecha.Value = HoyDia
        End If
    End If
End Sub

Private Function PrdEntregar() As Long
    Dim i As Long
    Dim num As Long
    num = 0
    For i = 1 To VSFG.Rows - 1
        If FormatoD4(VSFG.TextMatrix(i, 4)) > 0 Then
            num = num + 1
        End If
    Next i
    PrdEntregar = num
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtPedido" Then
        KeyCode = 0
        SendKeys vbKeyTab
    ElseIf KeyCode = vbKeyF4 Then
        optAbierto.Enabled = True
        optCalendario.Enabled = True
        optCalendarioAyer.Enabled = True
        optFiltro.Enabled = True
    End If
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa los objetos de conexión con la base de datos
    clsPedidos.Inicializar AdoConn, AdoConnMaster
    clsPed.Inicializar AdoConn, AdoConnMaster
    clsTC.Inicializar AdoConn, AdoConnMaster
    clsTFac.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsLstPrds.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    clsRet.Inicializar AdoConn, AdoConnMaster
    clsExis.Inicializar AdoConn, AdoConnMaster
    clsFormaPago.Inicializar AdoConn, AdoConnMaster
    
    dtpFechaFinAplicar.Value = HoyDia & " 00:00:00"
    dtpEjecutar.Value = HoyDia & " 00:00:00"
    dtpFechaFin.Value = HoyDia & " 00:00:00"
    dtpFechaFinAplicar.Value = HoyDia & " 00:00:00"
    dtpFechaIni.Value = HoyDia & " 00:00:00"
    
    '****** CREDITO
    VSFG.Tag = "N"
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT par_numero " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'MIC' "
    clsSql.Ejecutar (strSql)
    MINCredito = clsSql.adorec_Def("par_numero")
    
    strSql = " SELECT est_codigo,est_descripcion " & _
             " FROM est_pedido " & _
             " ORDER BY est_codigo"
    clsSql.Ejecutar strSql
    
    'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
    VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "*est_descripcion", "est_codigo")
    
    
    '****** TARJETAS
    strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo " & _
             " FROM tarjeta_credito " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY tar_cre_nombre "
    clsTC.Ejecutar (strSql)
    Set cmbTC.RowSource = clsTC.adorec_Def.DataSource
    cmbTC.ListField = "tar_cre_nombre"
    cmbTC.BoundColumn = "tar_cre_codigo"
    'cmbTC.BoundText = "EFE"
    
    
    '****** CLAVE
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSql)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")
    
    
    'Consulta todos los pedidos que pasan a bodega para ser revisados
    strSql = " SELECT RIGHT(ped_codigo,7)+0 as c, ped_fechamod, CONCAT(per_apellido,' ',per_nombre) as nombC, " & _
             " ped_observacion, ped_estado, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
             " IIF(pedido.ven_codigo='' OR pedido.ven_codigo is null,persona.ven_codigo,pedido.ven_codigo) as ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
             " per_sec_publico,per_siniva,per_fac_flete,per_dcto,pedido.tar_cre_codigo,ped_codigo " & _
             " FROM ((pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado) " & _
             " INNER JOIN persona ON (pedido.per_codigo = persona.per_codigo) " & _
             " AND (pedido.emp_codigo = persona.emp_codigo)) " & _
             " INNER JOIN tipo_factura ON (pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo) " & _
             " LEFT JOIN tarjeta_credito ON (pedido.emp_codigo = tarjeta_credito.emp_codigo) AND (pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo) " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado<>0 " & _
             " AND (ped_fecha='" & Format(HoyDia, "yyyy-MM-dd") & "' OR ped_estado=1) AND ped_codigo LIKE CONCAT('" & strSucursal & strPtoFactura & "'+0,'%') " & _
             " ORDER BY ped_estado,ped_codigo "
    'clsPedidos.Ejecutar (strSql)
    'Muestra los datos de los distintos proyectos de trabajo en un listado
    'Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
    
    cargarTipoPedido
    cargarGZDir
    'Consulta los recargos que puede manejar una empresa
    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    'Obtiene el IVA vigente para realizar la factura
    IVA = PorIVA
    LblIva = "IVA " & IVA & " %:"
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
'''    'strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
'''    '         " FROM forma_pago " & _
'''    '         " WHERE emp_codigo='" & strEmpresa & "' " & _
'''    '         " ORDER BY for_pag_nombre "
'''    'clsFPago.Ejecutar (strSql)
'''    'Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
'''    'CmbFpago.ListField = "for_pag_nombre"
'''    'CmbFpago.BoundColumn = "for_pag_codigo"
    'Consulta todos los tipos de factura y los muestra en un combo
    strSql = " SELECT * FROM tipo_factura ORDER BY tipo_fac_descripcion"
    clsTFac.Ejecutar (strSql)
    Set CmbTipoFac.RowSource = clsTFac.adorec_Def.DataSource
    CmbTipoFac.ListField = "tipo_fac_descripcion"
    CmbTipoFac.BoundColumn = "tipo_fac_codigo"
    'Coloca los botones de eliminar fila en el grid de recargos
    PonerBotones
    'Coloca la fecha actual
    dtpFecha.Value = HoyDia
    cmbNegocio_Change
    If lngFacturaDesde = 0 Or lngFacturaHasta = 0 Then
        IngresoBlocFactura
    End If
    txtFacturaDesde.Text = lngFacturaDesde
    txtFacturaHasta.Text = lngFacturaHasta
        
End Sub
Private Sub frmFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    frmFiltro.ZOrder 0
End Sub

Private Sub frmPrincipal_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    frmPrincipal.ZOrder 0
End Sub

Private Sub optListaPedido_Click()
    cmdLimpiar_Click
    txtPedido.Enabled = False
'    VSFGPeds.Height = 1455
    VSFGPeds.Rows = 1
'    cmdActualizar.Height = 1455
'    frmPed.Height = 1815
    
'    frmDet.Top = 3000
'    frmDet.Height = 5055
'    VSFG.Height = 2295
'    frmRec.Top = 2880
    cmbNegocio_Change

End Sub

Private Sub optNoPedido_Click()
    txtPedido.Enabled = True
'    VSFGPeds.Height = 855
    VSFGPeds.Rows = 1
'    cmdActualizar.Height = 855
'    frmPed.Height = 1215
    
'    frmDet.Top = 2400
'    frmDet.Height = 5655
'    VSFG.Height = 2895
'    frmRec.Top = 3480
    cmbNegocio_Change
End Sub

'Verifica cada 10 segundos si existe un nuevo pedido a revisar
Private Sub TmrAct_Timer()
    lblEstado.Caption = "Tiempo restante a ejecutar " & DateDiff("n", Now, dtpEjecutar.Value) & " minutos"
    lblEstado.Refresh
    If DateDiff("n", Now, dtpEjecutar.Value) <= 0 Then
        TmrAct.Enabled = False
        cmdEjecutar_Click
    End If
End Sub
Private Sub TxtDesc_LostFocus()
    'Calcula el total de la factura
    CalcuTotal
End Sub

Private Sub TxtDesc_Validate(Cancel As Boolean)
    TxtDesc.Text = Format(FormatoD2(TxtDesc.Text), "####0.00")
End Sub

Public Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLimpiar_Click
        If cmbNegocio.BoundText <> "" Then
            strSql = " SELECT tip_ped_ptofac " & _
                     " FROM tipo_pedido " & _
                     " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
                'Consulta todos los pedidos que pasan a bodega para ser revisados
                If chkCIRUC.Value = 1 Then
                    strCli = "CONCAT(per_ruc,' ',per_apellido,' ',per_nombre)"
                Else
                    strCli = "CONCAT(per_apellido,' ',per_nombre)"
                End If
                strSql = " SELECT RIGHT(ped_codigo,7)+0 as c, ped_fechamod, " & strCli & " as nombC, " & _
                         " ped_observacion, ped_estado, tipo_fac_descripcion, persona.per_codigo, cot_codigo, " & _
                         " IIF(pedido.ven_codigo='' OR pedido.ven_codigo is null,persona.ven_codigo,pedido.ven_codigo) as ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo," & _
                         " per_sec_publico,per_siniva,per_fac_flete,IIF(pedido.ped_egr_bodega=0,per_dcto,pedido.ped_egr_bodega),pedido.tar_cre_codigo,ped_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,lis_pre_codigo,ped_dctoadicional " & _
                         " FROM ((pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado) " & _
                         " INNER JOIN persona ON (pedido.per_codigo = persona.per_codigo) " & _
                         " AND (pedido.emp_codigo = persona.emp_codigo) AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "') " & _
                         " INNER JOIN tipo_factura ON (pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo) " & _
                         " LEFT JOIN tarjeta_credito ON (pedido.emp_codigo = tarjeta_credito.emp_codigo) AND (pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo) " & _
                         " Where pedido.emp_codigo='" & strEmpresa & "' AND ped_estado<>0 " & _
                         " AND ped_codigo ='" & txtPedido.Text & "' " & _
                         " ORDER BY ped_estado,ped_codigo "
                clsPedidos.Ejecutar (strSql)
                Set VSFGPeds.DataSource = clsPedidos.adorec_Def.DataSource
                
                strSql = " SELECT est_codigo,est_descripcion " & _
                         " FROM est_pedido " & _
                         " ORDER BY est_codigo"
                clsSql.Ejecutar strSql
                
                'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
                VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
                
                    If VSFGPeds.Rows > 1 Then
                        VSFGPeds.Row = 1
                        VSFGPeds.Col = 4
                        VSFGPeds.Select 1, 4
                        VSFGPeds_DblClick
                    End If
            End If
        Else
            Exit Sub
        End If
        txtPedido.Text = ""
    '    clsPedidos.Actualizar
        'Muestra los datos de los distintos pedidos en un listado
    
    End If
End Sub

Private Sub VSFGPeds_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Marca toda la fila con otra tonalidad si el pedido puede ser vendido
    If Col = 4 And Row <> 0 Then
        If VSFGPeds.TextMatrix(Row, Col) = 1 Then
            VSFGPeds.Select Row, 0, Row, VSFGPeds.Cols - 1
            VSFGPeds.FillStyle = flexFillRepeat
            VSFGPeds.CellBackColor = &HC0C0FF
        End If
    End If
End Sub


Private Sub cmbTC_Change()
    Dim i As Long, comis As Double
    Dim ProdAux As String
    Dim CantAux As Long
    Dim Tipo As Integer
    Dim cade As String
    Dim tipoPagoDetalleANT As String
    
    strCantidadTipoPedido = "det_ped_cant_confirmada"
    If Me.cmbNegocio.BoundText = "JON" Or Me.cmbNegocio.BoundText = "LEM" Then
        strCantidadTipoPedido = "det_ped_cant_entregada"
    End If
    If cmbTC.MatchedWithList = True Then
        If cmbTC <> "" Then
        '****** TARJETAS
        strSql = " SELECT tar_cre_codigo, tar_cre_nombre,tar_cre_porcentaje,tip_com_codigo,prd_codigo " & _
                 " FROM tarjeta_credito " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " ORDER BY tar_cre_nombre "
        clsTC.Ejecutar strSql
    
        clsTC.Filtrar "tar_cre_codigo='" & cmbTC.BoundText & "' "
        Tipo = FormatoD0(clsTC.adorec_Def("tip_com_codigo"))
        tipoPagoDetalleANT = tipoPagoDetalle
        tipoPagoDetalle = clsTC.adorec_Def("prd_codigo")
        If Tipo = 1 Then 'No comision
            ''dblComision = clsTC.adorec_Def("tar_cre_porcentaje")
            cade = ""
            dblComision = "0"
        ElseIf Tipo = 2 Then
            cade = ""
            dblComision = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
        ElseIf Tipo = 3 Then
            dblComision = "0"
            comis = FormatoD2(clsTC.adorec_Def("tar_cre_porcentaje"))
        End If
    
            strSqlPrd = " SELECT dep_codigo, det_pedido.prd_codigo, det_ped_precio,prd_nombre,ROUND((" & strCantidadTipoPedido & " * det_ped_precio)-IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & "),2) as total," & strCantidadTipoPedido & ",det_ped_cant_pedida, " & _
                    " (producto.prd_costo/(1 - 0.1)) as prd_costo, det_ped_precio*'" & 1 + dblComision / 100# & "' as lis_precio,prd_cambia_precio, " & _
                    " (IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ")) as dcto, " & _
                    " ROUND((" & strCantidadTipoPedido & " * det_ped_precio)-IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ")*(" & 1 + dblComision / 100# & "),2) as totales, " & _
                    " det_ped_precio,(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ") as det_ped_dcto " & _
                    " FROM pedido INNER JOIN det_pedido ON pedido.ped_codigo = det_pedido.ped_codigo " & _
                    " AND pedido.emp_codigo = det_pedido.emp_codigo " & _
                    " INNER JOIN persona ON persona.per_codigo=pedido.per_codigo AND persona.emp_codigo=pedido.emp_codigo " & _
                    " INNER JOIN categoria_p ON persona.cat_p_codigo=categoria_p.cat_p_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.emp_codigo=categoria_p.emp_codigo " & _
                    " INNER JOIN producto " & _
                    " ON det_pedido.emp_codigo = producto.emp_codigo AND det_pedido.prd_codigo = producto.prd_codigo " & _
                    " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                    " INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                    " AND producto.emp_codigo=lista_precio_p.emp_codigo AND lista_precio_p.lis_pre_codigo=categoria_p.lis_pre_codigo " & _
                    " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                    " AND producto_promo.prd_pro_fechaini<=CAST(pedido.ped_fechamod as date) AND producto_promo.prd_pro_fechafin>=CAST(pedido.ped_fechamod as date) AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                    " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                    " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                    " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                    " det_pedido.ped_codigo='" & VSFGPeds.TextMatrix(VSFGPeds.Row, 18) & "' " & _
                    " ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo "
            clsLstPrds.Ejecutar strSqlPrd
            
            VSFG.Tag = "P"
        For i = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(i, 1) <> "" Then
                ProdAux = VSFG.TextMatrix(i, 1)
                'CantAux = VSFG.TextMatrix(i, 4)
                VSFG.TextMatrix(i, 1) = ""
                VSFG.TextMatrix(i, 1) = ProdAux
                'VSFG.TextMatrix(i, 4) = CantAux
            End If
        Next i
        VSFG.Tag = "N"
            If Tipo = 3 Then
                strSql = " SELECT prd_codigo,prd_nombre " & _
                         " FROM producto " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & tipoPagoDetalle & "' "
                clsSql.Ejecutar strSql
                If clsSql.adorec_Def.RecordCount > 0 Then
                    VSFG.AddItem VSFG.TextMatrix(1, 0) & vbTab & clsSql.adorec_Def(0) & vbTab & clsSql.adorec_Def(1) & vbTab & "1" & vbTab & "1" & vbTab & ((FormatoD4(TxtSubTotal.Text) - FormatoD4(TxtDesc.Text)) * (FormatoD4(comis) / 100#)) & vbTab & "0" & vbTab & ((FormatoD4(TxtSubTotal.Text) - FormatoD4(TxtDesc.Text)) * (FormatoD4(comis) / 100#)) & vbTab & "1"
                End If
            Else
                For i = 1 To VSFG.Rows - 1
                    If VSFG.TextMatrix(i, 1) = tipoPagoDetalleANT Then
                        VSFG.RemoveItem i
                        Exit For
                    End If
                Next i
            End If
            
            
        End If
    End If

    CalcuTotal
    'txtComi.Text = FormatoD2(FormatoD4(TxtTotal.Text) * FormatoD4(comis / 100.00))
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Coloca la descripción del producto en caso que se haga un pedido manual y el usuario haya seleccionado un código de producto
    If Row > 0 And VSFG.Tag = "P" Then
        If Col = 2 Or Col = 1 Then
            
            If VSFG.TextMatrix(Row, 1) = "" Then
                VSFG.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            
            If Col = 1 Then
                VSFG.TextMatrix(Row, 2) = VSFG.TextMatrix(Row, 1)
            End If
            'Verifica que no se seleccione más de una vez el mismo producto en la misma bodega
            
            'Coloca los datos de un producto seleccionado
            If VSFG.TextMatrix(Row, 1) <> "" Then
                'Busca el producto seleccionado y coloca sus datos respectivos
                clsLstPrds.adorec_Def.MoveFirst
                clsLstPrds.Filtrar "dep_codigo='" & VSFG.TextMatrix(Row, 0) & "' AND prd_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
                If Not clsLstPrds.adorec_Def.EOF Then
                    
                    VSFG.TextMatrix(Row, 2) = clsLstPrds.adorec_Def("prd_nombre")
                    VSFG.TextMatrix(Row, 3) = clsLstPrds.adorec_Def("det_ped_cant_pedida")
                    VSFG.TextMatrix(Row, 4) = clsLstPrds.adorec_Def("" & strCantidadTipoPedido & "")
                    'Coloca el costo del producto en una columna oculta
                    ''''VSFG.TextMatrix(Row, 9) = clsLstPrds.adorec_Def("prd_costo")
                    ''''VSFG.TextMatrix(Row, 10) = Abs(FormatoD0(clsLstPrds.adorec_Def("prd_cambia_precio")))
                    ''''VSFG.TextMatrix(Row, 6) = 0#
                    'Verifica que el precio de la lista no sea menor al costo del producto y tampoco sea una cotización
                   ''' If clsLstPrds.adorec_Def("lis_precio") <>  Then ''''And tipoPed <> 1 Then
                    VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("lis_precio"))
'''''                    If clsLstPrds.adorec_Def("prd_costo") > clsLstPrds.adorec_Def("lis_pre_p_precio") Then ''''And tipoPed <> 1 Then
'''''                        If FormatoD4(clsLstPrds.adorec_Def("prd_costo")) <> FormatoD4(VSFG.TextMatrix(Row, 5)) Then
'''''                            VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("prd_costo"))
'''''                        End If
'''''                    Else
'''''                        If FormatoD4(clsLstPrds.adorec_Def("lis_pre_p_precio")) <> FormatoD4(VSFG.TextMatrix(Row, 5)) Then
'''''                            VSFG.TextMatrix(Row, 5) = FormatoD4(clsLstPrds.adorec_Def("lis_pre_p_precio"))
'''''                        End If
'''''                    End If
                    'Verifica que la existencia del producto sea mayor que cero
'                    If clsLstPrds.adorec_Def("exi_cantidad") > 0 Then
'                        VSFG.TextMatrix(Row, 4) = 1
'                    Else
'                        VSFG.TextMatrix(Row, 4) = 0
'                    End If
'                    VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD4(VSFG.TextMatrix(Row, 6)))
                    
                    dctoMax = 0
                    dctoMax = FormatoD2(VSFGPeds.TextMatrix(VSFGPeds.Row, 16))
                    strSql = " SELECT prd_pro_porcentaje " & _
                             " FROM producto_promo " & _
                             " WHERE emp_codigo = '" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(Row, 1) & "' " & _
                             " AND LEFT('" & VSFGPeds.TextMatrix(VSFGPeds.Row, 1) & "',10) BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                    clsSql.Ejecutar strSql
                    If clsSql.adorec_Def.RecordCount > 0 Then
                        If FormatoD2(clsSql.adorec_Def(0)) > FormatoD2(txtDescuento.Text) Then
                            dctoMax = FormatoD2(clsSql.adorec_Def(0))
                        End If
                    End If
                    VSFG.TextMatrix(Row, 6) = IIf(FormatoD2(clsLstPrds.adorec_Def("det_ped_dcto")) > FormatoD2(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD2(dctoMax) / 100#), FormatoD2(clsLstPrds.adorec_Def("det_ped_dcto")), FormatoD2(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) * FormatoD2(dctoMax) / 100#))
                    'dcto pedido adicional
                    VSFG.TextMatrix(Row, 6) = VSFG.TextMatrix(Row, 6) + FormatoD4((FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD2(VSFG.TextMatrix(Row, 6)))) * (VSFGPeds.TextMatrix(VSFGPeds.Row, 21) / 100))
                    
                    VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 5)) * FormatoD4(VSFG.TextMatrix(Row, 4)) - FormatoD2(VSFG.TextMatrix(Row, 6)))
                    'VSFG.TextMatrix(Row, 7) = FormatoD4(FormatoD4(VSFG.TextMatrix(Row, 7)) * FormatoD4(dctoMax) / 100.00)
                    ''''VSFG.TextMatrix(Row, 8) = clsLstPrds.adorec_Def("exi_cantidad")
                End If
                clsLstPrds.QuitarFiltro
                CalcuTotal
            End If
        End If
    End If
End Sub

Private Sub FacturaConFlete(strPedido As String)
''''''''    strSql = " REPLACE INTO det_pedido(emp_codigo, ped_codigo, prd_codigo, dep_codigo," & _
''''''''             " det_ped_cant_pedida, det_ped_cant_entregada, det_ped_cant_entregada," & _
''''''''             " det_ped_precio, det_ped_dcto, det_ped_descripcion, det_ped_fechamod, det_ped_usumod)" & _
''''''''             " VALUES('" & strEmpresa & "','" & strPedido & "','PR-','PRI'," & _
''''''''             " 1,1,1," & _
''''''''             " '2',0,'',CURRENT_TIMESTAMP,'" & strUsuario & "'"
'''''''    strSql = " REPLACE INTO det_pedido(emp_codigo, ped_codigo, prd_codigo, dep_codigo," & _
'''''''             " det_ped_cant_pedida, det_ped_cant_entregada, det_ped_cant_entregada," & _
'''''''             " det_ped_precio, det_ped_dcto, det_ped_descripcion, det_ped_fechamod, det_ped_usumod)" & _
'''''''             " SELECT pedido.emp_codigo,pedido.ped_codigo,producto_ciudad.prd_codigo,dep_codigo," & _
'''''''             " producto_ciudad.prd_ciu_cantidad,producto_ciudad.prd_ciu_cantidad,producto_ciudad.prd_ciu_cantidad," & _
'''''''             " producto_ciudad.prd_ciu_precio,0,'',CURRENT_TIMESTAMP,'" & strUsuario & "'" & _
'''''''             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
'''''''             " AND pedido.per_codigo=persona.per_codigo" & _
'''''''             " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo" & _
'''''''             " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
'''''''             " INNER JOIN producto_ciudad ON persona.emp_codigo=producto_ciudad.emp_codigo" & _
'''''''             " AND persona.ciu_codigo=producto_ciudad.ciu_codigo" & _
'''''''             " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
'''''''             " AND pedido.ped_codigo='" & strPedido & "'"
'''''''    clsPed.Ejecutar strSql, "M"
End Sub

Private Sub VSFGPeds_DblClick()
    FechaUltFac = ""
    Dim strForma As String, codDC As String
    Dim Fp As String
    'Verifica cuando se da un doble click sobre una fila del grid de pedidos
    dtpFecha.Value = HoyDia
    Me.CmdConfirmarTodo.Enabled = True
    'Limpia el grid de recargos
    VSFGReca.Rows = 2
    VSFGReca.Clear 1
    VSFG.Tag = "N"
    TxtObserv.Text = ""
    If VSFGPeds.Row > 0 Then
        '*** FACTURA FLETE
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 15))) = 1 Then
            FacturaConFlete VSFGPeds.TextMatrix(VSFGPeds.Row, 18)
        End If
        'Consulta el detalle de un pedido específico
        strSql = " SELECT dep_codigo, det_pedido.prd_codigo, prd_nombre, det_ped_cant_pedida, " & strCantidadTipoPedido & ", det_ped_precio,  " & _
                 " ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)!=0,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & " * det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ")/100.00),2)" & _
                 " + ROUND(((" & strCantidadTipoPedido & " * det_ped_precio)-ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)!=0,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & " * det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ")/100.00),2))*(pedido.ped_dctoadicional/100.00),2) as det_ped_dcto," & _
                 " ROUND(((" & strCantidadTipoPedido & " * det_ped_precio)-ROUND(IIF(ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2)!=0,ROUND(det_ped_dcto/det_ped_cant_pedida*" & strCantidadTipoPedido & ",2),ROUND(" & strCantidadTipoPedido & " * det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>" & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ",IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))," & FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16)) & ")/100.00),2))*(1-pedido.ped_dctoadicional/100.00),2) as total," & _
                 " prd_iva,det_ped_incentivo " & _
                 " FROM ((pedido INNER JOIN det_pedido ON (pedido.ped_codigo = det_pedido.ped_codigo) " & _
                 " AND (pedido.emp_codigo = det_pedido.emp_codigo)) INNER JOIN producto " & _
                 " ON (det_pedido.emp_codigo = producto.emp_codigo) AND (det_pedido.prd_codigo = producto.prd_codigo)) " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND producto_promo.prd_pro_fechaini<=CAST(pedido.ped_fechamod as date) AND producto_promo.prd_pro_fechafin>=CAST(pedido.ped_fechamod as date) AND producto_promo.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' AND " & _
                 " pedido.ped_codigo= " & VSFGPeds.TextMatrix(VSFGPeds.Row, 18) & _
                 " ORDER BY producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,det_pedido.prd_codigo "
        clsPed.Ejecutar (strSql)
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsPed.adorec_Def.DataSource
        'Muestra el número del pedido a modificar
        LblPedido.Caption = VSFGPeds.TextMatrix(VSFGPeds.Row, 18)
        'Muestra el tipo de factura sugerido
        CmbTipoFac = VSFGPeds.TextMatrix(VSFGPeds.Row, 5)
        'Obtiene el código del cliente
        CodPer = VSFGPeds.TextMatrix(VSFGPeds.Row, 6)
        codVen = VSFGPeds.TextMatrix(VSFGPeds.Row, 8)
        codTC = VSFGPeds.TextMatrix(VSFGPeds.Row, 10)
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 13))) = 0 Then
            SecPublico = False
        Else
            SecPublico = True
        End If
        If Abs(FormatoD0(VSFGPeds.TextMatrix(VSFGPeds.Row, 14))) = 0 Then
            SinIVA = False
        Else
            SinIVA = True
        End If
        
        txtDescuento.Text = FormatoD4(VSFGPeds.TextMatrix(VSFGPeds.Row, 16))
        cmbTC.BoundText = ""
        cmbTC.BoundText = VSFGPeds.TextMatrix(VSFGPeds.Row, 17)
        
        'Obtiene el código de la cotización relacionada
        codCot = VSFGPeds.TextMatrix(VSFGPeds.Row, 7)
        'Calcula los totales de la factura
        CalcuDesc
        CalcuTotal
        Dim especial As Boolean
        '********* forma de pago ******************
        'CmbFpago.BoundText = VSFGPeds.TextMatrix(VSFGPeds.Row, 12)
        'Nueva Implementacion
        strSql = " SELECT persona.per_especial,persona.per_email,CONCAT(COALESCE(persona.per_apellido,''),' ',COALESCE(persona.per_nombre,'')) as cli, " & _
                 " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,N9.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,N8.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,N7.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,N6.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,N5.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,N4.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,N3.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,N2.per_email," & _
                 " IIF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,N1.per_email,''))))))))) as emailpapa"
        strSql = strSql & " FROM persona " & _
                 " LEFT JOIN persona as N1 ON N1.emp_codigo=persona.emp_codigo " & _
                 " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
                 " LEFT JOIN persona as N2 ON N2.emp_codigo=persona.emp_codigo " & _
                 " AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1 " & _
                 " LEFT JOIN persona as N3 ON persona.emp_codigo = N3.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1 " & _
                 " LEFT JOIN persona as N4 ON persona.emp_codigo = N4.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & CodPer & "' "
        clsFPago.Ejecutar strSql
        especial = False
        If clsFPago.adorec_Def.RecordCount > 0 Then
            especial = CBool(clsFPago.adorec_Def(0))
        End If
        
        If Not especial Then
        codDC = ""
        strForma = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo  " & _
                  " FROM forma_pago " & _
                  " WHERE emp_codigo='" & strEmpresa & "' " & _
                  " AND for_pag_codigo IN ('CONT','TAR','TAD'"
        Fp = VSFGPeds.TextMatrix(VSFGPeds.Row, 12)
        If Val(TxtTotal.Text) < Val(MINCredito) Then
            strSql = " SELECT COALESCE(egreso.egr_fecha,'" & HoyDia & "') as fecha, COALESCE(egreso.for_pag_codigo,'CONT') as codigo, COALESCE(forma_pago.for_pag_tiempo,0) as tiempo  " & _
                " FROM egreso " & _
                " INNER JOIN forma_pago ON forma_pago.emp_codigo=egreso.emp_codigo AND egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                " WHERE egreso.emp_codigo = '" & strEmpresa & "' AND egr_anulado=0 AND egreso.tip_egr_codigo='FAC' " & _
                " AND egreso.egr_total>='" & Val(MINCredito) & "' AND egreso.per_codigo = '" & CodPer & "' ORDER BY egreso.egr_fecha DESC LIMIT 1 "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                FechaUltFac = Format(DateAdd("d", CDbl(clsSql.adorec_Def("tiempo")), clsSql.adorec_Def("fecha")), "yyyy-MM-dd")
                
                If CStr(dtpFecha.Value) < FechaUltFac Then
                    'coloca el tiempo
                    codDC = Right(clsSql.adorec_Def("codigo"), 1)
                    Dim Diferencia As Long
                    Diferencia = CLng(DateDiff("d", dtpFecha.Value, FechaUltFac))
                    strForma = strForma & ",'" & clsSql.adorec_Def("codigo") & "')"
                    Fp = clsSql.adorec_Def("codigo")
                Else
                    strForma = strForma & ") "
                    Fp = "CONT"
                End If
            Else
                strForma = strForma & ") "
                Fp = "CONT"
            End If
        Else
            strForma = strForma & ",'" & VSFGPeds.TextMatrix(VSFGPeds.Row, 12) & "') "
        End If
        strForma = strForma & " ORDER BY for_pag_nombre "
        clsFPago.Ejecutar strForma
        
        Else
            'SI ES ESPECIAL LE MANDA TODAS LAS FORMAS DE PAGO
            strForma = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo  " & _
                  " FROM forma_pago " & _
                  " WHERE emp_codigo='" & strEmpresa & "' " & _
                  " AND for_pag_codigo IN ('CONT','TAR','TAD','" & VSFGPeds.TextMatrix(VSFGPeds.Row, 12) & "') ORDER BY for_pag_nombre"
            clsFPago.Ejecutar strForma
            Fp = VSFGPeds.TextMatrix(VSFGPeds.Row, 12)
        End If
        Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
        CmbFpago.ListField = "for_pag_nombre"
        CmbFpago.BoundColumn = "for_pag_codigo"
        CmbFpago.BoundText = Fp
        
        
                
        '******************************************
        
        'Verifica que no se haya facturado o que haya algún producto en el pedido para no poder volver a facturarlo
        If VSFGPeds.TextMatrix(VSFGPeds.Row, 4) = 2 Or VSFGPeds.TextMatrix(VSFGPeds.Row, 4) = 3 Then
            CmdConfirmar.Enabled = False
            VSFGReca.Enabled = False
        Else
            CmdConfirmar.Enabled = True
            VSFGReca.Enabled = True
        End If
        If VSFGPeds.TextMatrix(VSFGPeds.Row, 9) <> "" Then
            'MsgBox VSFGPeds.TextMatrix(VSFGPeds.Row, 9), vbInformation, "Observaciones"
        End If
    End If
    If VSFGPeds.TextMatrix(VSFGPeds.Row, 19) <> 0 Then
        'MsgBox "Cliente BLOQUEADO en Cartera", vbInformation, "BLOQUEO"
        CmdConfirmar.Enabled = False
    End If
End Sub

Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
    'And VSFGReca.TextMatrix(OldRow, 1) <> ""
    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 Then
        VSFGReca.AddItem ""
        PonerBotones
    End If
End Sub

Private Sub VSFGReca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 0 del recargo
    If Col = 2 Then
        Cancel = True
    ElseIf (Col = 3 And VSFGReca.TextMatrix(Row, 1) = "") Then
        Cancel = True
        VSFGReca.TextMatrix(Row, Col) = ""
    End If
    
End Sub

Private Sub VSFGReca_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    With VSFGReca
        ' only interesetd in left button
        If Button <> 1 Then Exit Sub
        
        ' get cell that was clicked
        Dim r&, c&
        r = .MouseRow
        c = .MouseCol
        
        ' make sure the click was on the sheet
        If r < 0 Or c < 0 Then Exit Sub
        
        If (c <> 0 Or r = (.Rows - 1)) Then Exit Sub
         
        ' make sure the click was on a cell with a button
        If .Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = .Cell(flexcpLeft, r, c) + .Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        
        ' click was on a button: do the work
         .Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
            
        'Recorro el FlexGrid para poner números a las filas
            
        If respuesta = vbYes Then
             Dim i As Integer
              .RemoveItem (r)
             PonerBotones
             CalcuDesc
             CalcuTotal
        Else
             .Cell(flexcpPicture, r, c) = imgBtnUp
        End If
        
        ' cancel default processing
        ' note: this is not strictly necessary in this case, because
        '       the dialog box already stole the focus etc, but let's be safe.
        Cancel = True
    End With
End Sub

Private Sub VSFGReca_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Busca y coloca el valor del recargo seleccionado
    If Row > 0 And VSFGReca.TextMatrix(Row, 1) <> "" And Col <> 3 Then
        clsRecargos.Filtrar "oca_codigo='" & VSFGReca.TextMatrix(Row, 1) & "'"
        VSFGReca.TextMatrix(Row, 2) = clsRecargos.adorec_Def("oca_nombre")
        VSFGReca.TextMatrix(Row, 3) = clsRecargos.adorec_Def("oca_precio")
        clsRecargos.QuitarFiltro
        'Verifica que no se haya escogido antes el mismo recargo, en ese caso suma sus valores
        For i = 1 To VSFGReca.Rows - 1
            If VSFGReca.TextMatrix(Row, 1) = VSFGReca.TextMatrix(i, 1) And Row <> i Then
                VSFGReca.TextMatrix(i, 3) = Val(VSFGReca.TextMatrix(i, 3)) + (VSFGReca.TextMatrix(Row, 3))
                VSFGReca.RemoveItem Row
                PonerBotones
                Exit For
            End If
        Next i
    End If
    CalcuDesc
    CalcuTotal
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cargarGZDir()
    strSql = " SELECT '-1' as codigo,' Todas las Provincias' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT pai_codigo as codigo,pai_nombre as nombre " & _
             " FROM pais " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbPais.RowSource = clsSql.adorec_Def.DataSource
    cmbPais.ListField = "nombre"
    cmbPais.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los Gerentes de Zona' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_gz=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbGerente.RowSource = clsSql.adorec_Def.DataSource
    cmbGerente.ListField = "nombre"
    cmbGerente.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los Directores' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_di=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbDirector.RowSource = clsSql.adorec_Def.DataSource
    cmbDirector.ListField = "nombre"
    cmbDirector.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los Emprendedores' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_em=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbEmprendedor.RowSource = clsSql.adorec_Def.DataSource
    cmbEmprendedor.ListField = "nombre"
    cmbEmprendedor.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los Ejecutivos Especial' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_ee=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbEjecutivo.RowSource = clsSql.adorec_Def.DataSource
    cmbEjecutivo.ListField = "nombre"
    cmbEjecutivo.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los N5' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_n5=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbN5.RowSource = clsSql.adorec_Def.DataSource
    cmbN5.ListField = "nombre"
    cmbN5.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los N6' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_n6=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbN6.RowSource = clsSql.adorec_Def.DataSource
    cmbN6.ListField = "nombre"
    cmbN6.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los N7' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_n7=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbN7.RowSource = clsSql.adorec_Def.DataSource
    cmbN7.ListField = "nombre"
    cmbN7.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los N8' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_n8=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbN8.RowSource = clsSql.adorec_Def.DataSource
    cmbN8.ListField = "nombre"
    cmbN8.BoundColumn = "codigo"
    
    strSql = " SELECT '-1' as codigo,' Todos los N9' as nombre " & _
             " UNION " & _
             " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
             " FROM persona p1 " & _
             " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
             " AND p1.per_es_n9=1 AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbN9.RowSource = clsSql.adorec_Def.DataSource
    cmbN9.ListField = "nombre"
    cmbN9.BoundColumn = "codigo"
    
    
    cmbPais.BoundText = "-1"
    cmbGerente.BoundText = "-1"
    cmbDirector.BoundText = "-1"
    cmbEmprendedor.BoundText = "-1"
    cmbEjecutivo.BoundText = "-1"
    cmbN5.BoundText = "-1"
    cmbN6.BoundText = "-1"
    cmbN7.BoundText = "-1"
    cmbN8.BoundText = "-1"
    cmbN9.BoundText = "-1"
    
End Sub

Private Sub ActualizaPrecio()
    Dim i As Long
    Dim clsCambioPrecio As New clsConsulta
    clsCambioPrecio.Inicializar AdoConn, AdoConnMaster
    For j = 1 To VSFGPeds.Rows
        VSFGPeds.Row = j
        VSFGPeds.ShowCell j, 1
        VSFGPeds_DblClick
        If VSFG.Rows > 1 Then
'            clsCambioPrecio.Inicializar AdoConn, AdoConnMaster
            For i = 1 To VSFG.Rows - 1
                If Left(VSFG.TextMatrix(i, 1), 3) <> "PR-" Then
                    strSql = " SELECT producto.prd_codigo, " & _
                             " (producto.prd_costo/(1 - 0.1)) as prd_costo, lis_pre_p_precio " & _
                             " FROM producto INNER JOIN lista_precio_p ON producto.prd_codigo=lista_precio_p.prd_codigo " & _
                             " AND producto.emp_codigo=lista_precio_p.emp_codigo" & _
                             " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                             " AND lista_precio_p.lis_pre_codigo=" & VSFGPeds.TextMatrix(VSFGPeds.Row, 20) & " "
                    clsCambioPrecio.Ejecutar strSql
                    If clsCambioPrecio.adorec_Def.RecordCount > 0 Then
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_precio='" & clsCambioPrecio.adorec_Def("lis_pre_p_precio") & "' " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                                 " AND ped_codigo=" & LblPedido.Caption & " "
                        clsCambioPrecio.Ejecutar strSql, "M"
                    End If
                    
                    dctoMax = 0
                    dctoMax = FormatoD2(txtDescuento.Text)
                    strSql = " SELECT prd_pro_porcentaje " & _
                             " FROM producto_promo " & _
                             " WHERE emp_codigo = '" & strEmpresa & "' " & _
                             " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                             " AND LEFT('" & VSFGPeds.TextMatrix(VSFGPeds.Row, 1) & "',10) BETWEEN prd_pro_fechaini AND prd_pro_fechafin "
                    clsCambioPrecio.Ejecutar strSql
                    If clsCambioPrecio.adorec_Def.RecordCount > 0 Then
                        If FormatoD2(clsCambioPrecio.adorec_Def(0)) > FormatoD2(TxtDesc.Text) Then
                            dctoMax = FormatoD2(clsCambioPrecio.adorec_Def(0))
                            
                            strSql = " UPDATE det_pedido " & _
                                     " SET det_ped_dcto=ROUND(det_ped_cant_entregada*det_ped_precio*" & dctoMax & "/100.00,4) " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "'" & _
                                     " AND ped_codigo=" & LblPedido.Caption & " "
                            clsCambioPrecio.Ejecutar strSql, "M"
                        End If
                    End If
                End If
            Next i
        End If
    Next j

End Sub
