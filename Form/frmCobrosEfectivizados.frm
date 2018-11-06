VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCobrosEfectivizados 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros a Efectivizar"
   ClientHeight    =   11715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmCobrosEfectivizados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11715
   ScaleWidth      =   11175
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
      Height          =   1095
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   10935
      Begin VB.CommandButton cmdMenos 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         TabIndex        =   89
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         TabIndex        =   88
         Top             =   360
         Width           =   375
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1200
         TabIndex        =   48
         Top             =   255
         Width           =   4215
         _ExtentX        =   7435
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
         Left            =   1170
         TabIndex        =   70
         Top             =   720
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   315
         Left            =   1170
         TabIndex        =   71
         Top             =   1080
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   315
         Left            =   1170
         TabIndex        =   72
         Top             =   1440
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   315
         Left            =   1170
         TabIndex        =   73
         Top             =   1800
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   315
         Left            =   1185
         TabIndex        =   74
         Top             =   2160
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   315
         Left            =   1185
         TabIndex        =   75
         Top             =   2520
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   315
         Left            =   1185
         TabIndex        =   76
         Top             =   2880
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   315
         Left            =   1185
         TabIndex        =   77
         Top             =   3240
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   315
         Left            =   1200
         TabIndex        =   78
         Top             =   3600
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eje.E. N4:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   390
         TabIndex        =   87
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empren N3:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   86
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dir N2:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   585
         TabIndex        =   85
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.Zona N1:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   84
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   83
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   82
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   81
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   80
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   79
         Top             =   3660
         Width           =   240
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
         Left            =   435
         TabIndex        =   49
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   10935
      Begin VB.Frame fraEfec 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Efectivización de Cobros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   52
         Top             =   6000
         Width           =   10455
         Begin VB.TextBox txtDescripciont 
            Enabled         =   0   'False
            Height          =   765
            Left            =   7920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   840
            Width           =   2415
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Contabilización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   2535
            Begin VB.OptionButton optBanco 
               BackColor       =   &H00DDDDDD&
               Caption         =   "En Banco o Cajas"
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Value           =   -1  'True
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker dtpFechaConta 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   720
               TabIndex        =   62
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
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
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   70254595
               CurrentDate     =   37463
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00C3DBD1&
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
               Left            =   120
               TabIndex        =   63
               Top             =   780
               Width           =   495
            End
         End
         Begin VB.Frame frmBanco 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Para depositar en :"
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
            Height          =   1455
            Left            =   2880
            TabIndex        =   53
            Top             =   240
            Width           =   3855
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1530
               TabIndex        =   54
               Top             =   1080
               Width           =   2055
            End
            Begin MSDataListLib.DataCombo dcmbCuenta 
               Height          =   315
               Left            =   1530
               TabIndex        =   55
               Top             =   720
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcmbBancoE 
               Height          =   315
               Left            =   1530
               TabIndex        =   56
               Top             =   360
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Banco:"
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
               Left            =   120
               TabIndex        =   59
               Top             =   405
               Width           =   510
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº. Documento:"
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
               Left            =   120
               TabIndex        =   58
               Top             =   1110
               Width           =   1125
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Bancaria:"
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
               Left            =   120
               TabIndex        =   57
               Top             =   765
               Width           =   1245
            End
         End
         Begin MSDataListLib.DataCombo dcmbTipo 
            Height          =   315
            Left            =   7920
            TabIndex        =   64
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbldescripcion1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
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
            Left            =   6840
            TabIndex        =   67
            Top             =   885
            Width           =   900
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de nota:"
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
            Left            =   6840
            TabIndex        =   65
            Top             =   525
            Width           =   930
         End
      End
      Begin VB.TextBox txtFpag 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Frame fraRet 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos Retención"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7327
         TabIndex        =   38
         Top             =   2520
         Width           =   3255
         Begin NEED2.dtpFecha dtpFechaR 
            Height          =   285
            Left            =   1680
            TabIndex        =   69
            Top             =   263
            Width           =   1335
            _extentx        =   2355
            _extenty        =   503
            value           =   42892.721712963
         End
         Begin VB.TextBox txtDocumentoR 
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtSerieR 
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtAutorizacionR 
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   210
            TabIndex        =   42
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Autorizacion"
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
            Left            =   210
            TabIndex        =   41
            Top             =   1350
            Width           =   1425
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Serie"
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
            Left            =   210
            TabIndex        =   40
            Top             =   630
            Width           =   375
         End
         Begin VB.Label Label15 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
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
            Left            =   210
            TabIndex        =   39
            Top             =   300
            Width           =   585
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   9480
         Width           =   1215
      End
      Begin VB.CheckBox chkAnticipo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Anticipo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   525
         Left            =   1485
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   5400
         Width           =   7695
      End
      Begin VB.TextBox txtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   9480
         Width           =   1935
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   9480
         Width           =   1815
      End
      Begin VB.OptionButton optproveedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Proveedor"
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
         Left            =   1200
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton optcliente 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente"
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
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Forma de Pago"
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
         Left            =   735
         TabIndex        =   32
         Top             =   2520
         Width           =   2040
         Begin VB.OptionButton optNCredito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Nota de Crédito"
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
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optNDebito 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Nota de Débito"
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
            Left            =   240
            TabIndex        =   9
            Top             =   1320
            Width           =   1455
         End
         Begin VB.OptionButton optefectivo 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Efectivo"
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
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optcheque 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Documento"
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
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Height          =   1095
         Left            =   2895
         TabIndex        =   29
         Top             =   2520
         Width           =   4320
         Begin NEED2.dtpFecha dtpFechaCh 
            Height          =   285
            Left            =   2040
            TabIndex        =   68
            Top             =   600
            Width           =   2055
            _extentx        =   3625
            _extenty        =   503
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            _Version        =   393216
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   70254595
            CurrentDate     =   37463
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Cobro:"
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
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de documento:"
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
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1560
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Height          =   1455
         Left            =   2895
         TabIndex        =   25
         Top             =   3480
         Width           =   4320
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   2025
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   945
            TabIndex        =   13
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbDocumento 
            Height          =   315
            Left            =   945
            TabIndex        =   12
            Top             =   600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc:"
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
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   675
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
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
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   510
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de documento:"
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
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   1350
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG1 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   10695
         _cx             =   88164785
         _cy             =   88148698
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCobrosEfectivizados.frx":030A
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
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   555
         TabIndex        =   19
         Top             =   7920
         Width           =   9600
         _cx             =   88162853
         _cy             =   88148698
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCobrosEfectivizados.frx":0501
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin MSDataListLib.DataCombo dcmbBeneficiario 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNota 
         Height          =   315
         Left            =   720
         TabIndex        =   43
         Top             =   4320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbDeudorCh 
         Height          =   315
         Left            =   1485
         TabIndex        =   17
         Top             =   5040
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deudor:"
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
         Left            =   810
         TabIndex        =   90
         Top             =   5085
         Width           =   570
      End
      Begin VB.Label lblfp 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago Cliente:"
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
         Left            =   240
         TabIndex        =   50
         Top             =   2235
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   165
         TabIndex        =   46
         Top             =   4680
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   4365
         Width           =   375
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   480
         TabIndex        =   36
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label lbltotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES:"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   9600
         Width           =   855
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   120
         Picture         =   "frmCobrosEfectivizados.frx":05F1
         ToolTipText     =   "Elimina una Fila"
         Top             =   8880
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   120
         Picture         =   "frmCobrosEfectivizados.frx":0727
         Top             =   9120
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deudor:"
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
         Left            =   2520
         TabIndex        =   34
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         Left            =   6600
         TabIndex        =   33
         Top             =   2197
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   11280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5595
      TabIndex        =   21
      Top             =   11280
      Width           =   1575
   End
End
Attribute VB_Name = "frmCobrosEfectivizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso del comprobante de egresos comunes                         #
'#  frmComprobanteEgresoComun V1.0                                              #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el comprobante de egresos comunes                     #
'#  Permite ingresar los datos de egresos comunes y sus detalles                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  COMP_EGRESO: Esta tabla almacena los datos del comprobante                  #
'#  PERSONA: donde se guardan los datos de los benficiarios de los comprobantes #
'#  DET_COMP_EGRESO: Guarda los detalles del comprobante de Egreso              #
'#  RET_COMP_EGRESO: Guarda las retenciones que puede tener el comprobante      #
'#  CTA_BANCO: consulta los datos del numero de cuenta y el último cheque       #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsPag As New clsConsulta
Private clsPer As New clsConsulta
Private clsSql As New clsConsulta
Private clsCod As New clsConsulta
Private clsPgd As New clsConsulta
Private clsAsi As New clsConsulta
Private clsdoc As New clsConsulta
Private clscdo As New clsConsulta
Private clsRet As New clsConsulta
Private clsNot As New clsConsulta

Private clsDet As New clsConsulta
Private clsTip As New clsConsulta

Private lonNFijas As Long
Private strSql As String
Private FilaCxC As Long
Private Descripcion As String
Private strPersona As String

Private Sub chkAnticipo_Click()
'    dcmbBeneficiario_Change
    dcmbBeneficiario_Validate False
End Sub

Private Sub cmbGerente_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbDirector_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbEmprendedor_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbEjecutivo_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbN5_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbN6_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbN7_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbN8_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbN9_Change()
    If optcliente.Value = True Then
        OptCliente_Click
    End If
End Sub

Private Sub cmbNegocio_Change()
    dcmbBeneficiario.Tag = "SI"
    dcmbBeneficiario.BoundText = ""
    cargarGZDir
    If optcliente.Value = True Then
        OptCliente_Click
    ElseIf optproveedor.Value = True Then
        optproveedor_Click
    End If
    dcmbBeneficiario.Tag = "NO"
End Sub


Private Sub cmdMas_Click()
    frmFiltro.Height = 4095
End Sub

Private Sub cmdMenos_Click()
    frmFiltro.Height = 1095
End Sub

Private Sub dcmbBanco_Change()
    LlenarVariableDescripcion
End Sub

Private Sub dcmbBancoE_Change()
    dcmbCuenta = ""
    clsSql.Inicializar AdoConn, AdoConnMaster
    If dcmbBancoE.Text = "" Then
        dcmbCuenta.Text = ""
        Exit Sub
    Else
        strSql = " SELECT cta_ban_numero, cta_ban_ctaconta, ban_codigo " & _
                 " FROM cta_banco " & _
                 " WHERE ban_codigo = '" & dcmbBancoE.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
        clsSql.Ejecutar strSql
        
        If clsSql.adorec_Def.EOF = False Then
            Set dcmbCuenta.RowSource = clsSql.adorec_Def.DataSource
            dcmbCuenta.ListField = "cta_ban_numero"
            dcmbCuenta.BoundColumn = "cta_ban_ctaconta"
            'dcmbCuenta.Tag = clsCta.adorec_Def("cta_ban_ctaconta")
            'dcmbCuenta.Text = clsCta.adorec_Def("cta_ban_numero")
            
        Else
            Set dcmbCuenta.RowSource = Nothing
        End If
        
    End If
End Sub


Private Sub dcmbCuenta_Change()
 Dim j As Long
    'cmdConsultar.Enabled = True
    If VSFG.Rows > 1 Then
        Llenar_Grid (1)
    End If
End Sub

Private Sub dcmbDeudorCh_Change()
    LlenarVariableDescripcion
End Sub

Private Sub dcmbDocumento_Change()
    LlenarVariableDescripcion
    LlenarDatosDeDeposito
End Sub
Private Sub LlenarDatosDeDeposito()
    Dim TipoDocPago As String
    Dim clsSqlA As New clsConsulta
    If optcheque.Value = True Then
        TipoDocPago = dcmbDocumento.BoundText
    ElseIf optefectivo.Value = True Then
        TipoDocPago = "CH"
    End If
    clsSqlA.Inicializar AdoConn, AdoConnMaster
    'If optcheque.value = True Then
        strSql = " SELECT ban_codigo,cta_ban_numero " & _
                 " FROM tipo_doc_pago " & _
                 " WHERE tip_doc_pag_codigo = '" & TipoDocPago & "' "
        clsSqlA.Ejecutar strSql
        
        If clsSqlA.adorec_Def.RecordCount > 0 Then
            dcmbBancoE.BoundText = clsSqlA.adorec_Def("ban_codigo")
            dcmbCuenta = clsSqlA.adorec_Def("cta_ban_numero")
        Else
            dcmbBancoE.BoundText = ""
            dcmbCuenta = ""
        End If
    'End If
End Sub
Private Sub dcmbNota_Change()
    If dcmbNota.MatchedWithList = True Then
        clsNot.Filtrar "ing_codigo='" & dcmbNota.BoundText & "'"
        txtSaldo.Text = FormatoD2(clsNot.adorec_Def("sal"))
        txtSaldo.Tag = FormatoD2(clsNot.adorec_Def("ing_saldo"))
        dcmbNota.Tag = clsNot.adorec_Def("ing_numasiento")
        'dtpFecha.value = Format(clsNot.adorec_Def("ing_fecha"), "yyyy-mm-dd")
        'dtpFechaCh.value = Format(clsNot.adorec_Def("ing_fecha"), "yyyy-mm-dd")
        txtDocumento.Text = dcmbNota.Text
    Else
        txtDocumento.Text = ""
        txtSaldo.Text = "0"
    End If
End Sub

Private Sub dcmbTipo_Change()
 If dcmbTipo = "" Then
    txtDescripciont = ""
    Exit Sub
 End If
strSql = " SELECT CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
         " FROM tipo_nota " & _
         " WHERE tip_not_d_c = 'C' AND  tip_not_codigo = '" & dcmbTipo.BoundText & "' "
clsTip.Ejecutar strSql

If clsTip.adorec_Def.EOF = False Then
    txtDescripciont = clsTip.adorec_Def("descripcion")
Else
    txtDescripciont.Text = ""
End If

End Sub

Private Sub dtpFechaCh_Change()
    If dtpFecha.Value < dtpFechaCh.Value Then
        If dcmbDeudorCh.Locked = True Then
            dcmbDeudorCh.BoundText = ""
        End If
        dcmbDeudorCh.Locked = False
    Else
        dcmbDeudorCh.Locked = True
        dcmbDeudorCh.BoundText = ""
    End If
    LlenarVariableDescripcion
End Sub

'Private Sub dtpFechaCh_Validate(Cancel As Boolean)
'    If dtpFecha.Value < dtpFechaCh.Value Then
'        If dcmbDeudorCh.Locked = True Then
'            dcmbDeudorCh.BoundText = ""
'        End If
'        dcmbDeudorCh.Locked = False
'    Else
'        dcmbDeudorCh.Locked = True
'        dcmbDeudorCh.BoundText = ""
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsPag = Nothing
    Set clsPer = Nothing
    Set clsSql = Nothing
    Set clsCod = Nothing
    Set clsPgd = Nothing
    Set clsAsi = Nothing
    Set clsdoc = Nothing
    Set clscdo = Nothing
    Set clsRet = Nothing
    Set clsNot = Nothing
    Set clsTip = Nothing
    Set clsDet = Nothing
End Sub
Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True And i >= lonNFijas + 1 Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
    
'    For i = 1 To (VSFG1.Rows - 1)
'        VSFG1.TextMatrix(i, 0) = i
'    Next i
End Sub

Private Sub CalcuTotal()

   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double

    'Calcula total debe
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    txtTotalHaber = Format(SumaHaber, "##0.00")
   TxtTotal = Format(txtTotalDebe - txtTotalHaber, "##0.00")
    
End Sub
Private Sub pagos()
    Dim aux As Long
    Dim i As Long
    Dim j As Long
    Dim TRET As Double
    aux = 0
    TRET = 0
    If VSFG.Rows > 2 Then
        For j = 2 To lonNFijas - 1
                VSFG.TextMatrix(j, 3) = 0
        Next j
    End If
    For i = 1 To VSFG1.Rows - 1
        If Abs(FormatoD0(VSFG1.TextMatrix(i, 1))) = 1 Then
            If aux <> VSFG1.TextMatrix(i, 0) Then
                Suma = Suma + Val(VSFG1.TextMatrix(i, 11))
                aux = VSFG1.TextMatrix(i, 0)
            End If
            If Trim(VSFG1.TextMatrix(i, 14)) <> "" Then
                For j = 1 To lonNFijas - 1
                    If Trim(VSFG1.TextMatrix(i, 14)) = Trim(VSFG.TextMatrix(j, 1)) Then
                        VSFG.TextMatrix(j, 3) = Val(VSFG.TextMatrix(j, 3)) + Val(VSFG1.TextMatrix(i, 13))
                        TRET = TRET + FormatoD2(VSFG1.TextMatrix(i, 13))
                        j = lonNFijas
                    End If
                'If Trim(VSFG1.TextMatrix(j, 13)) <> "" And VSFG1.TextMatrix(j, 12) <> "Tot.Ret." Then
                '    TRET = TRET + FormatoD2(VSFG1.TextMatrix(j, 13))
                'End If
                Next j
            End If
        End If
                        'If Trim(VSFG1.TextMatrix(i, 13)) <> "" And VSFG1.TextMatrix(i, 12) <> "Tot.Ret." Then
                    'TRET = TRET + FormatoD2(VSFG1.TextMatrix(i, 13))
                'End If

    Next i
    txtValor.Tag = FormatoD2(TRET)
    txtValor = Format(Suma, "##0.00")
    LlenarVariableDescripcion
End Sub
Private Sub Limpiar()
    VSFG1.Clear 1
    VSFG1.Rows = 2
    VSFG.Clear 1
    VSFG.Rows = 2
    dcmbBeneficiario.Text = ""
    dcmbDeudorCh.Text = ""
    dcmbBanco.Text = ""
    txtDocumento = ""
    txtDescripcion = ""
    txtTotalHaber = 0
    txtTotalDebe = 0
    TxtTotal = 0
    txtValor = 0
    chkAnticipo.Value = 0
    'txtSaldoReal = 0
    'txtDisponible = 0
    'txtPrevisto = 0
    txtp = 0
    txtD = 0
    dcmbBanco.Text = ""
    txtNumero = ""
    dcmbTipo.Text = ""
    
    dtpFechaConta.Value = Format(HoyDia, "yyyy-MM-dd")
End Sub

Private Sub cmdAceptar_Click()
    Dim ElAsiento As String
    Dim maxpago As String, maxpagoAux As String, booPasar As Boolean, respuesta As Integer, booGrabar As Boolean
          
    '****************************************
        Dim Descripcion As String           '
        Dim FechaD As String                '
        Dim fechah As String                '
        Dim fechac As String                '
        Dim Pendiente As Integer            '
        Dim EstadoCH As String              '
        Dim campoAsiento As String          '
        Dim CHPost As String                '
        Dim nl As String                    '
    '****************************************
    If dcmbBancoE.MatchedWithList = False Then
        MsgBox "Seleccione una Banco"
        Exit Sub
    End If
    If dcmbCuenta.MatchedWithList = False Then
        MsgBox "Seleccione una Cuenta"
        Exit Sub
    End If
    maxpago = ""
    maxpagoAux = ""
    booPasar = False
    booGrabar = False
    If FormatoD2(txtTotalDebe.Text) <> FormatoD2(txtTotalHaber.Text) Then
        MsgBox "No esta cuadrado el comprobante de ingreso", vbInformation, "CONTABILIDAD"
        Exit Sub
    End If
    If dtpFecha.Value < dtpFechaCh.Value Then
        If dcmbDeudorCh.MatchedWithList = False Then
            MsgBox "debe escojer un deudor del documento", vbInformation, "CONTABILIDAD"
            Exit Sub
        End If
    End If
            
    clsdoc.Inicializar AdoConn, AdoConnMaster
    clscdo.Inicializar AdoConn, AdoConnMaster
'Comprueba que todos los datos esten ingresados
    ffch = Format(dtpFecha.Value, "yyyy-mm-dd")
    Fecha = Format(dtpFechaCh.Value, "yyyy-mm-dd")
    
    fechac = Format(dtpFechaConta.Value, "yyyy-mm-dd")  '*********
    fechac = ffch
    If (IsDate(ffch) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Pagos"
        Exit Sub
    End If
    

    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd

'    a = VSFG.Rows - 1
'    For i = 1 To a
'        For j = i + 1 To a
'            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) And VSFG.TextMatrix(i, 5) = VSFG.TextMatrix(j, 5) Then
'                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
'                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
'                VSFG.RemoveItem j
'                a = a - 1
'                j = j - 1
'            End If
'            If j >= a Then
'                Exit For
'            End If
'        Next j
'    Next i

    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber And optNCredito.Value = False Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "Pagos"
        Exit Sub
    ElseIf FormatoD2(txtSaldo.Text) < FormatoD2(txtValor.Text) And optNCredito.Value = True Then
        MsgBox "El Saldo de la Nota de Credito debe ser menor al abono de las Facturas", vbInformation, "Pagos"
        Exit Sub
    Else
        'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If VSFG.TextMatrix(1, 1) = "" Or txtDescripcion = "" Or dcmbBeneficiario = "" Or txtDocumento = "" Or (optBanco.Value = True And (dcmbBancoE = "" Or dcmbTipo.Text = "" Or txtNumero = "")) Then     '****************
            MsgBox "No estan ingresados todos los datos", vbInformation, "Pagos"
            Exit Sub
        Else
            '*************************************************
                If optBanco.Value = True Then
                    Descripcion = UCase("Depósito Banco: " + " " + dcmbBancoE.Text + " " + "No. de Documento:" + " " + txtNumero + " " + "Cantidad: " + " " + txtValor)
                    Pendiente = 0
                    EstadoCH = "COBRADO"
                    campoAsiento = "asi_numasiento"
                End If
            '*************************************************
        
        
            If optcheque.Value = True Then
                Dim intPend As Integer
                intPend = 0
                If ffch < Fecha And FormatoD2(txtValor.Text) <> 0 Then
                    intPend = 1
                End If
            End If
            'Ingreso de datos en la tabla pago
            n = VSFG1.Rows - 1
            
            While booPasar = False
                maxpago = InputBox("No. de Recibo", "No. de Recibo", maxpagoAux)
                If Trim(maxpago) <> "" Then
                    maxpago = strSucursal & strPtoFactura & Format(maxpago, "0000000")
                    strSql = " SELECT count(*) as Num FROM doc_pago " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND doc_pag_codigo LIKE '" & maxpago & "%'"
                    clsdoc.Ejecutar strSql
                    If clsdoc.adorec_Def("Num") <> 0 Then
                        respuesta = MsgBox("Ese recibo ya ha sido emitido, desea añadir el registro?", vbQuestion + vbYesNo, "Cobros")
                        If respuesta = vbYes Then
                            maxpago = maxpago & "-" & Chr(64 + FormatoD0(clsdoc.adorec_Def("Num")))
                            booPasar = True
                            booGrabar = True
                        Else
                            booPasar = False
                            booGrabar = False
                        End If
                    Else
                        booPasar = True
                        booGrabar = True
                    End If
                Else
                    booPasar = True
                    booGrabar = False
                End If
            Wend
            
         
            If booGrabar = True Then
                ElAsiento = "NULL"
                If optcheque.Value = True Then
                    strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, tip_doc_pag_codigo, ban_codigo, doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_doc ," & _
                             "                       per_codigo, doc_pag_valor, doc_pag_observacion, doc_pag_estado,doc_pag_pendiente,doc_pag_anticipo,per_codigo_ch, doc_pag_fechamod, doc_pag_usumod)" & _
                             " VALUES ('" & maxpago & "', '" & strEmpresa & "', '" & dcmbDocumento.BoundText & "', '" & dcmbBanco.BoundText & "', '" & txtDocumento & "','" & ffch & "', '" & Fecha & "'," & _
                             "         '" & dcmbBeneficiario.BoundText & "', '" & txtValor & "', '" & UCase(txtDescripcion) & "', 'GIRADO','" & intPend & "','" & chkAnticipo.Value & "','" & dcmbDeudorCh.BoundText & "', CURRENT_TIMESTAMP, '" & strUsuario & "') "
                ElseIf optefectivo.Value = True Then
                    strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_doc, doc_pag_fecha_efec, " & _
                             "                       per_codigo, doc_pag_valor, doc_pag_observacion,doc_pag_estado,doc_pag_pendiente,doc_pag_anticipo,per_codigo_ch, doc_pag_fechamod, doc_pag_usumod)" & _
                             " VALUES ('" & maxpago & "', '" & strEmpresa & "', '" & txtDocumento & "', '" & ffch & "', '" & ffch & "','" & ffch & "'," & _
                             "         '" & dcmbBeneficiario.BoundText & "', '" & txtValor & "', '" & UCase(txtDescripcion.Text) & "', 'GIRADO','" & intPend & "','" & chkAnticipo.Value & "','" & dcmbDeudorCh.BoundText & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                Else
                    maxpago = 0
                    Dim clsInventa As New clsInventario
                    clsInventa.Inicializar AdoConn, AdoConnMaster
                    If optNCredito.Value = True Then
                        clsInventa.strTipo = "DCL"
                        clsInventa.strIE = "I"
                        clsInventa.strDoc = dcmbNota.BoundText
                        clsInventa.ModificaIng , , , , , , , , , , , , , , , FormatoD2(txtSaldo.Tag) + FormatoD2(txtValor.Text)
                    Else
                        clsInventa.strTipo = "DPV"
                        clsInventa.strIE = "E"
                        clsInventa.strDoc = dcmbNota.BoundText
                        clsInventa.ModificaEgr , , , , , , , , , , , , , , , FormatoD2(txtSaldo.Tag) + FormatoD2(txtValor.Text)
                    End If
                    Set clsInventa = Nothing
                    ElAsiento = "'" & dcmbNota.Tag & "'"
                End If
                
                clsdoc.Ejecutar strSql, "M"
                
                If chkAnticipo.Value = 0 Then
                    
                    For i = 1 To n
                        If VSFG1.TextMatrix(i, 0) <> VSFG1.TextMatrix(i - 1, 0) Then
                            k = VSFG1.TextMatrix(i, 11)
                            If (VSFG1.TextMatrix(i, 11) <> "" Or VSFG1.TextMatrix(i, 11) <> "0") And VSFG1.TextMatrix(i, 1) <> "0" Then
                                'Calcula el máximo codigo de pago para la cuenta
                                 strSql = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                                          " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                                          "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                                          "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                                          " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'C'" & _
                                          " GROUP BY pago.emp_codigo"
                                clsCod.Ejecutar strSql
                                If clsCod.adorec_Def.EOF Then
                                    maxpag = 1
                                Else
                                    maxpag = clsCod.adorec_Def("pag") + 1
                                End If
                                Dim ValorPago As Double
                                If optNDebito.Value = True Then
                                    ValorPago = FormatoD2(VSFG1.TextMatrix(i, 11)) * -1
                                Else
                                    ValorPago = FormatoD2(VSFG1.TextMatrix(i, 11))
                                End If
                                
                                strSql = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, " & _
                                         " pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                                         " VALUES ('" & strEmpresa & "', '" & Val(VSFG1.TextMatrix(i, 2)) & "', 'C', '" & Val(maxpag) & "', '" & ffch & "', '" & ValorPago & "', " & _
                                         " '" & txtDocumento & "', '" & UCase(txtDescripcion) & "', " & _
                                         " '" & maxpago & "'," & ElAsiento & ",CURRENT_TIMESTAMP, '" & strUsuario & "') "
                                clsPag.Ejecutar strSql, "M"
                                dblRet = 0
                                If VSFG1.TextMatrix(i, 12) <> "" And Val(VSFG1.TextMatrix(i, 13)) <> 0 And Trim(VSFG1.TextMatrix(i, 15)) <> "" And Val(VSFG1.TextMatrix(i, 16)) <> 0 Then
                                    strSql = " INSERT INTO comprobante_retencion (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,com_ret_total,com_ret_fecha,com_ret_serie,com_ret_numero,com_ret_autorizacion,com_ret_fechamod,com_ret_usumod) VALUES ('" & _
                                             strEmpresa & "','" & VSFG1.TextMatrix(i, 2) & "','C',0,'" & Format(dtpFechaR.Value, "yyyy-mm-dd") & "','" & txtSerieR.Text & "','" & txtDocumentoR.Text & "','" & txtAutorizacionR.Text & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                                    clsPgd.Ejecutar strSql, "M"
                                    strSql = " INSERT INTO det_comp_ret (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_com_ret_fechamod,det_com_ret_usumod) VALUES ('" & _
                                             strEmpresa & "','" & VSFG1.TextMatrix(i, 2) & "','C','" & VSFG1.TextMatrix(i, 12) & "','" & VSFG1.TextMatrix(i, 13) * 100# / VSFG1.TextMatrix(i, 15) & "','" & VSFG1.TextMatrix(i, 15) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                                    clsPgd.Ejecutar strSql, "M"
                                    dblRet = FormatoD2(VSFG1.TextMatrix(i, 13))
                                    banfin = False
                                    If i + 1 <= VSFG1.Rows - 1 Then
                                        If VSFG1.TextMatrix(i, 2) <> VSFG1.TextMatrix(i + 1, 2) Then
                                            banfin = True
                                        End If
                                    Else
                                        banfin = True
                                    End If
                                    While banfin = False
                                        If VSFG1.TextMatrix(i + 1, 12) <> "" And Val(VSFG1.TextMatrix(i + 1, 13)) <> 0 Then
                                            strSql = " INSERT INTO det_comp_ret (emp_codigo,cue_p_c_codigo,cue_p_c_tipo,ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_com_ret_fechamod,det_com_ret_usumod) VALUES ('" & _
                                                     strEmpresa & "','" & VSFG1.TextMatrix(i + 1, 2) & "','C','" & VSFG1.TextMatrix(i + 1, 12) & "','" & VSFG1.TextMatrix(i + 1, 13) * 100# / VSFG1.TextMatrix(i + 1, 15) & "','" & VSFG1.TextMatrix(i + 1, 15) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                                            clsPgd.Ejecutar strSql, "M"
                                            dblRet = dblRet + FormatoD2(VSFG1.TextMatrix(i + 1, 13))
                                        End If
                                        i = i + 1
                                        If i = VSFG1.Rows - 1 Then
                                            banfin = True
                                        End If
                                        If banfin = False Then
                                            If VSFG1.TextMatrix(i, 2) <> VSFG1.TextMatrix(i + 1, 2) Then
                                                banfin = True
                                            End If
                                        End If
                                    Wend
                                    strSql = " UPDATE comprobante_retencion SET com_ret_total='" & dblRet & "' " & _
                                             " WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_codigo='" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_tipo='C'"
                                    clsPgd.Ejecutar strSql, "M"
                                End If
                                If intPend = 0 Then
                                    If FormatoD2(VSFG1.TextMatrix(i, 10)) <= FormatoD2(VSFG1.TextMatrix(i, 11)) + FormatoD2(dblRet) + 0.005 And optNDebito.Value = False Then
                                        strSql = " SELECT MAX(pag_fecha) as fec " & _
                                                 " FROM pago " & _
                                                 " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' " & _
                                                 " AND emp_codigo = '" & strEmpresa & "' " & _
                                                 " AND pag_observacion NOT LIKE 'ANULADO%' " & _
                                                 " GROUP BY emp_codigo"
                                        clsPgd.Ejecutar strSql, "M"
                                        strSql = " UPDATE cuenta_p_c " & _
                                                 " SET cue_p_c_fechapago='" & clsPgd.adorec_Def("fec") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                                 " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & VSFG1.TextMatrix(i, 2) & "' AND cue_p_c_egr_codigo = '" & VSFG1.TextMatrix(i, 4) & "' AND emp_codigo = '" & strEmpresa & "' "
                                        clsPgd.Ejecutar strSql, "M"
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
               
                If optcheque.Value = True Or optefectivo.Value = True Then
                    With VSFG
                        For i = 1 To .Rows - 1
                            'Ingresa el detalle del asiento del egreso
                            If .TextMatrix(i, 1) = "" Then
                                Exit For
                            Else
                                If i = 1 And Not (Val(.TextMatrix(i, 3)) = 0 And Val(.TextMatrix(i, 4)) = 0) Then
                                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                                         " VALUES ('" & strEmpresa & "','" & maxpago & "',0, '*', '" & .TextMatrix(i, 5) & "','" & FormatoD2(.TextMatrix(i, 3)) & "', '" & FormatoD2(.TextMatrix(i, 4)) & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                                clsSql.Ejecutar strSql, "M"
                                ElseIf Not (Val(.TextMatrix(i, 3)) = 0 And Val(.TextMatrix(i, 4)) = 0) Then
                                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                                         " VALUES ('" & strEmpresa & "','" & maxpago & "',0, '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 5) & "','" & FormatoD2(.TextMatrix(i, 3)) & "', '" & FormatoD2(.TextMatrix(i, 4)) & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                                clsSql.Ejecutar strSql, "M"
                                End If
                            End If
                        Next i
                    End With
                    
                    
                    
                End If
                
                
                '**************************************************************
                Dim clsAsientoE As New clsContable
                clsAsientoE.Inicializar AdoConn, AdoConnMaster
                clsAsientoE.NuevoAsiento "I", fechac, 0, 0, FormatoD2(txtTotalDebe), Descripcion
                
                Descripcion = Descripcion & vbNewLine & "CLI: " & dcmbBeneficiario.Text & " DOC: " & dcmbDocumento.Text & "(" & dcmbBanco.Text & ") No:" & txtDocumento.Text & " ABONO: " & txtValor.Text & Replace(" OBS:" & UCase(txtDescripcion.Text), vbNewLine, " ")
                'Actualiza asientos en pagos
                strSql = " UPDATE pago " & _
                         " SET asi_numasiento='" & clsAsientoE.NumAsiento & _
                         "' , pag_fechamod= CURRENT_TIMESTAMP, pag_usumod='" & strUsuario & "' " & _
                         " WHERE doc_pag_codigo= '" & maxpago & "' AND emp_codigo = '" & strEmpresa & "' " & _
                         " AND cue_p_c_tipo='C' "
                clsPag.Ejecutar strSql, "M"
                'Actualiza la tabla doc_pago
                strSql = " UPDATE doc_pago " & _
                         " SET doc_pag_fecha_efec='" & fechac & _
                         "'," & campoAsiento & "='" & clsAsientoE.NumAsiento & _
                         "',doc_pag_pendiente='" & Pendiente & "', doc_pag_estado = '" & EstadoCH & "' , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                         " WHERE doc_pag_codigo= '" & maxpago & "' AND emp_codigo = '" & strEmpresa & "' "
                clsPag.Ejecutar strSql, "M"
                
                strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo," & _
                         " max(doc_pago.doc_pag_fecha_doc) as fecha,cuenta_p_c.cue_p_c_valor,COALESCE(sum(p2.pag_monto),0),COALESCE(com_ret_total,0)," & _
                         " cuenta_p_c.cue_p_c_valor-COALESCE(sum(p2.pag_monto),0)-COALESCE(com_ret_total,0) as saldo " & _
                         " FROM pago as p1 INNER JOIN cuenta_p_c ON cuenta_p_c.cue_p_c_codigo=p1.cue_p_c_codigo " & _
                         " AND cuenta_p_c.cue_p_c_tipo=p1.cue_p_c_tipo " & _
                         " AND cuenta_p_c.emp_codigo=p1.emp_codigo " & _
                         " INNER JOIN pago as p2 ON cuenta_p_c.cue_p_c_codigo=p2.cue_p_c_codigo " & _
                         " AND cuenta_p_c.cue_p_c_tipo=p2.cue_p_c_tipo " & _
                         " AND cuenta_p_c.emp_codigo=p2.emp_codigo " & _
                         " INNER JOIN doc_pago ON p2.doc_pag_codigo=doc_pago.doc_pag_codigo " & _
                         " AND p2.emp_codigo=doc_pago.emp_codigo " & _
                         " AND doc_pago.doc_pag_pendiente=0 AND doc_pago.doc_pag_estado!='ANULADO' " & _
                         " LEFT JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                         " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                         " AND cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                         " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                         " AND p1.doc_pag_codigo='" & maxpago & "' " & _
                         " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_valor,com_ret_total "
                clsPag.Ejecutar strSql, "M"
                While Not clsPag.adorec_Def.EOF
                    If (FormatoD2(clsPag.adorec_Def("saldo")) <= 0) Then
                        strSql = " UPDATE cuenta_p_c " & _
                                 " SET cue_p_c_fechapago='" & clsPag.adorec_Def("fecha") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                 " WHERE cue_p_c_tipo= 'C' " & _
                                 " AND cue_p_c_codigo= '" & clsPag.adorec_Def("cue_p_c_codigo") & _
                                 "' AND cue_p_c_egr_codigo = '" & clsPag.adorec_Def("cue_p_c_egr_codigo") & _
                                 "' AND emp_codigo = '" & strEmpresa & "' "
                        clsSql.Ejecutar strSql, "M"
                    End If
                    clsPag.adorec_Def.MoveNext
                Wend
                
                clsAsientoE.ModificarAsiento FormatoD2(txtTotalDebe), FormatoD2(txtTotalHaber), , , , Descripcion
                'ingreso del detalle del asiento
                With VSFG
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 1) = "" Then
                            Exit For
                        Else
                            If FormatoD2(.TextMatrix(i, 3)) <> 0 Or FormatoD2(.TextMatrix(i, 4)) <> 0 Then
                                clsAsientoE.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                            End If
                        End If
                    Next i
                End With
                
                If optBanco.Value = True Then
                    'GENERACION DE LA NOTA DE CREDITO
                    'Calcula el código de la Nota de Crédito
                    strSql = " SELECT cta_ban_saldoreal, cta_ban_saldoprevisto " & _
                              " FROM cta_banco " & _
                              " WHERE cta_ban_numero = '" & dcmbCuenta.Text & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsSql.Ejecutar strSql
                    If Not clsSql.adorec_Def.EOF Then
                        saldoreal = clsSql.adorec_Def("cta_ban_saldoreal") + txtValor
                        saldoPrevisto = clsSql.adorec_Def("cta_ban_saldoprevisto") + txtValor
                    Else
                        saldoreal = txtValor
                        saldoPrevisto = txtValor
                    End If
                    'Guarda los datos de la Nota de Crédito
                    
                    strSql = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto,asi_numasiento,not_d_c_conciliado , not_d_c_fechamod, not_d_c_usumod) " & _
                             " VALUES ('C','" & AutoNumero_Cuenta & "', '" & dcmbCuenta.Text & "', '" & dcmbBancoE.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & txtNumero.Text & "','" & fechac & "','" & Descripcion & "','" & txtValor.Text & "','" & clsAsientoE.NumAsiento & "',0, CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsSql.Ejecutar strSql, "M"
                    'Actualiza los valores de los saldos
                    strSql = " UPDATE cta_banco " & _
                             " SET cta_ban_saldoreal= '" & saldoreal & "',cta_ban_saldoprevisto= '" & saldoPrevisto & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                             " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBancoE.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
                    clsSql.Ejecutar strSql, "M"
    
                End If
                
                
                
                '**************************************************************
                
                
                
                MsgBox " Los datos han sido ingresados", vbInformation, "Ingresos"
                
                
                
                
                
                If optcheque.Value = True Or optefectivo.Value = True Then
                    'drptReciboCaja.Tag = maxpago
                    If chkAnticipo.Value = 1 Then
                        'drptReciboCaja.EsAnticipo = True
                    Else
                        'drptReciboCaja.EsAnticipo = False
                    End If
                    'drptReciboCaja.Show
                    '''MsgBox "Cobro Ingresado"
                    Dim RepCobro As New frmReporte
                        RepCobro.strNumero = maxpago
                        RepCobro.strReporte = "rptReciboCaja"
                        RepCobro.Show
                Else
                    frmReporte.strAsiento = dcmbNota.Tag
                    frmReporte.strReporte = "rptAsiento"
                    frmReporte.Show
                End If
                
                '****************************************************
'                    'Impresion de Comprobante de Ingreso
'                    Dim rptNuevo As New frmReporte
'                    rptNuevo.strAsiento = clsAsientoE.NumAsiento
'                    rptNuevo.strReporte = "rptAsiento"
'                    rptNuevo.Show
                    
                    Dim rptCompIng As New frmReporte
                    rptCompIng.strAsiento = clsAsientoE.NumAsiento
                    rptCompIng.strReporte = "rptComprobanteIngreso"
                    rptCompIng.Show
                    
                    Set clsAsientoE = Nothing
                '******************************************************
                
                
                Set clsAsiento = Nothing
                Limpiar
                dcmbDocumento = ""
                dcmbAsiento = ""
                dcmbNota = ""
                txtSaldo.Text = ""
            End If
           
        End If
         
    End If
           
   
End Sub


Private Function AutoNumero_Cuenta() As String
    'Pone el número de cuenta por cobrar / pagar siguiente
    Set clsNumCuentas = New clsConsulta
    clsNumCuentas.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT COALESCE(max(not_d_c_codigo),0) as num_cuenta" & _
             " FROM nota_d_c " & _
             " WHERE tip_not_d_c = 'C' AND emp_codigo = '" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsNumCuentas.Ejecutar strSql
    If Not clsNumCuentas.adorec_Def.EOF Then
        Var_NumCuenta = Val(clsNumCuentas.adorec_Def("num_cuenta")) + 1
    Else
        Var_NumCuenta = 1
    End If
    AutoNumero_Cuenta = Var_NumCuenta
End Function

Private Sub Llenar_Grid(ByVal Row)
'limpia el grid para los asientos
If Row = 1 Then
    VSFG.TextMatrix(Row, 1) = "*"
    VSFG.TextMatrix(Row, 2) = "CAJA"
End If
If optBanco.Value = True Then
    strSql = " SELECT cta_codigo,cta_nombre " & _
             " FROM ctaconta " & _
             " WHERE cta_codigo = '" & dcmbCuenta.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
End If
'Cuenta del Banco
clsSql.Ejecutar strSql
'coloca los datos de los asientos de los documentos seleccionados

'cuentas con * se coloca la cuenta del banco escogido para realizar el depósito
        For j = 1 To VSFG.Rows - 1
            If VSFG.TextMatrix(j, 1) = "*" Then
                If clsSql.adorec_Def.RecordCount > 0 Then
                     VSFG.TextMatrix(j, 1) = clsSql.adorec_Def("cta_codigo")
                     VSFG.TextMatrix(j, 2) = clsSql.adorec_Def("cta_nombre")
                Else
                    If j > 1 Then
                        VSFG.TextMatrix(j, 1) = ""
                        VSFG.TextMatrix(j, 2) = ""
                    End If
                End If
            End If
        Next j


'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd
 
    CalcuTotal
End Sub



Private Sub LlenarVariableDescripcion()
    Dim Coma As String
    Dim TextoInicio As String
    Dim TextoRetencion As String
    Dim NumItems As Integer
    NumItems = 0
    Descripcion = ""
    TextoInicio = "FACTURA: "
    For i = 1 To VSFG1.Rows - 1
        If Val(VSFG1.TextMatrix(i, 1)) = -1 Then
            If VSFG1.TextMatrix(i, 2) <> VSFG1.TextMatrix(i - 1, 2) Then
                NumItems = NumItems + 1
                If (NumItems > 1) Then
                    Coma = ", "
                    TextoInicio = ""
                End If
                If VSFG1.TextMatrix(i, 13) <> "" And VSFG1.TextMatrix(i, 12) <> "Tot.Ret." Then
                    TextoRetencion = " (CON RETENCIÓN)"
                    
                Else
                    TextoRetencion = ""
                End If
                If FormatoD2(VSFG1.TextMatrix(i, 10)) = FormatoD2(VSFG1.TextMatrix(i, 11)) Then
                    txtAbono = " CANCELA"
                ElseIf FormatoD2(VSFG1.TextMatrix(i, 11)) <> 0 Then
                    txtAbono = " ABONA"
                End If
                Descripcion = TextoInicio & Descripcion & Coma & VSFG1.TextMatrix(i, 4) & txtAbono & TextoRetencion
                
            End If
        End If
    Next i
    PonerDescripcion2
End Sub

Private Sub PonerDescripcion2()
    Dim Cadena1 As String
    Dim Cadena2 As String
    Dim Cadena3 As String
    Dim Cadena4 As String
    Dim Cadena5 As String
    Dim Cadena6 As String
    
    If optcheque.Value = True Then
        If dcmbDocumento <> "" Then
            Cadena1 = dcmbDocumento & " "
        End If
        If txtDocumento <> "" Then
            Cadena2 = txtDocumento & " "
        End If
        'Cadena2 = Cadena2 & "(" & FechaDocumento & ") - "
        If dcmbBanco <> "" Then
            Cadena3 = dcmbBanco & " "
        End If
        If dtpFecha.Value < dtpFechaCh.Value Then
            Cadena6 = vbNewLine & "DOCUMENTO POR COBRAR DE: " & Me.dcmbDeudorCh.Text & " "
        End If
    ElseIf optefectivo.Value = True Then
        Cadena1 = "EFECTIVO"
    ElseIf optNCredito.Value = True Then
        Cadena1 = "NOTA DE CRÉDITO " & dcmbNota.Text & " (" & dcmbNota.BoundText & ")"
    Else
        Cadena1 = "NOTA DE DÉBITO " & dcmbNota.Text & " (" & dcmbNota.BoundText & ")"
    End If
    If Descripcion <> "" Then
        Cadena4 = Descripcion & " - " & vbNewLine
    End If
    If dcmbBeneficiario <> "" Then
        Cadena5 = dcmbBeneficiario & " - " & vbNewLine
    End If
    txtDescripcion = Cadena5 & Cadena4 & Cadena1 & Cadena2 & Cadena3 & Cadena6
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub dcmbBeneficiario_Validate(Cancel As Boolean)
    On Error Resume Next
    
    optefectivo.Value = True
    'Llenar_Grid 1
    If dcmbBeneficiario.MatchedWithList = True Or dcmbBeneficiario.Tag = "SI" Then
        txtValor = 0
        txtDocumento = ""
        dcmbDocumento = ""
        dcmbBancoE = ""
        dcmbCuenta = ""
        txtNumero = ""
        dtpFechaConta.Value = HoyDia
        dcmbBanco = ""
        dcmbNota = ""
        txtSaldo = ""
        txtFpag = ""
        t = "P"
        txtFpag.Visible = False
        lblfp.Visible = False
        If Me.optcliente.Value = True Then
            t = "C"
        End If
        cmdAceptar.Enabled = True
        VSFG1.Enabled = True
        strPersona = "'" & dcmbBeneficiario.BoundText & "'"
        strSql = " SELECT per_codigo_rel " & _
                 " FROM persona_relacion " & _
                 " WHERE per_codigo = '" & dcmbBeneficiario.BoundText & "' AND emp_codigo = '" & strEmpresa & "' "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            While Not clsSql.adorec_Def.EOF
                strPersona = strPersona & ",'" & clsSql.adorec_Def("per_codigo_rel") & "'"
                clsSql.adorec_Def.MoveNext
            Wend
        End If
        
        strSql = " SELECT com_ret_fecha,COALESCE(com_ret_serie,'') as com_ret_serie,COALESCE(com_ret_numero,'') as com_ret_numero,COALESCE(com_ret_autorizacion,'') as com_ret_autorizacion" & _
                 " FROM cuenta_p_c INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE per_codigo IN (" & strPersona & ") AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
                 " ORDER BY com_ret_fecha "
                 ' DESC LIMIT 1
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            txtAutorizacionR.Text = clsSql.adorec_Def("com_ret_autorizacion")
            txtSerieR.Text = clsSql.adorec_Def("com_ret_serie")
        Else
            txtAutorizacionR.Text = ""
            txtSerieR.Text = ""
        End If
        txtDocumentoR.Text = ""
        
        
     'Consulta para el grid sobre las cuentas por pagar del beneficiario seleccionado
        strSql = " SELECT ' ' as a,'0' as b, cuenta_p_c.cue_p_c_codigo, CONCAT(cue_p_c_fra_cuenta, '/' , cue_p_c_tot_cuenta ) as cue_p_c_fra_cuenta, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,DATEDIFF(DAY, cue_p_c_fechapropuesta, CURRENT_TIMESTAMP) AS dven, cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d, ' ' as e,IIF(com_ret_total IS NULL,' ','Tot.Ret.'),IIF(com_ret_total IS NULL,'0',com_ret_total),' ',' ',IIF(comprobante_retencion.com_ret_total IS NULL,'1','0') as f " & _
                 " FROM  (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo)" & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE per_codigo IN (" & strPersona & ") AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cue_p_c_fra_cuenta,cue_p_c_tot_cuenta,cue_p_c_egr_codigo,cue_p_c_descripcion,cue_p_c_fechaemision,cue_p_c_fechapropuesta,cue_p_c_valor,com_ret_total HAVING cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0)>0 " & _
                 " ORDER BY cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.EOF = False Then
            Valor = clsSql.adorec_Def("cue_p_c_valor")
            Set VSFG1.DataSource = clsSql.adorec_Def.DataSource
             VSFG1.ColDataType(1) = flexDTBoolean
            'ponerBotones
        Else
            Valor = 0
            VSFG1.Clear 1
            VSFG1.Rows = 2
        
        End If
        strSql = " SELECT ret_codigo, ret_nombre, ret_ctacontacli,ret_porcentaje " & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " AND ret_ctacontacli!='' " & _
                 " AND ret_activo=1 " & _
                 " ORDER BY ret_codigo"
        clsRet.Ejecutar strSql
    
         VSFG1.ColComboList(12) = "#;|" & VSFG1.BuildComboList(clsRet.adorec_Def, "*ret_codigo,ret_porcentaje,ret_nombre", "ret_codigo")
      'Consulta el saldo de la cuenta
      n = VSFG1.Rows - 1
      For i = 1 To n
        VSFG1.TextMatrix(i, 0) = i
      Next i
      
        strSql = " SELECT DISTINCT ret_ctacontacli,cta_nombre " & _
                 " FROM retencion INNER JOIN ctaconta ON retencion.emp_codigo=ctaconta.emp_codigo AND retencion.ret_ctacontacli=ctaconta.cta_codigo " & _
                 " WHERE retencion.emp_codigo='" & strEmpresa & "' " & _
                 " AND ret_activo=1 "
        clsSql.Ejecutar strSql
        clsSql.adorec_Def.MoveFirst
        VSFG.Rows = 1
        While Not clsSql.adorec_Def.EOF
            VSFG.AddItem "" & vbTab & clsSql.adorec_Def("ret_ctacontacli") & vbTab & clsSql.adorec_Def("cta_nombre") & vbTab & "0.00" & vbTab & "0.00"
            clsSql.adorec_Def.MoveNext
        Wend
        VSFG.AddItem "" & vbTab & "*" & vbTab & "CAJA" & vbTab & txtValor & vbTab & "0.00", 1
        lonNFijas = VSFG.Rows
        
        'Buscar cuenta del cliente de CxC
        If Me.chkAnticipo.Value = 0 Then
            strSql = " SELECT IIF(cat_p_ctaconta IS NULL OR cat_p_ctaconta='',par_texto,cat_p_ctaconta) as par_texto,tip_ped_ptofac " & _
                     " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                     " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                     " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                     " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                     " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                     " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                     " AND per_codigo='" & dcmbBeneficiario.BoundText & "'"
        Else
            strSql = " SELECT IIF(cat_p_ctaconta_ant IS NULL OR cat_p_ctaconta_ant='',par_texto,cat_p_ctaconta_ant) as par_texto,tip_ped_ptofac " & _
                     " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                     " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                     " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                     " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                     " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                     " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                     " AND per_codigo='" & dcmbBeneficiario.BoundText & "'"
        End If
        clsSql.Ejecutar strSql
        If clsSql.adorec_Def.RecordCount > 0 Then
            strPtoFactura = clsSql.adorec_Def("tip_ped_ptofac")
            VSFG.AddItem ""
            Me.VSFG.TextMatrix(VSFG.Rows - 1, 1) = clsSql.adorec_Def(0)
            Me.VSFG.TextMatrix(VSFG.Rows - 1, 3) = 0
            Me.VSFG.TextMatrix(VSFG.Rows - 1, 4) = 0
            FilaCxC = VSFG.Rows - 1
        End If
        If chkAnticipo.Value = 1 Then
            txtValor.Text = "0.00"
            txtValor.Locked = False
            VSFG1.Enabled = False
            VSFG1.Cell(flexcpForeColor, 1, 1, VSFG1.Rows - 1, 16) = &H8000000F
        Else
            txtValor.Text = "0.00"
            pagos
            txtValor.Locked = True
            VSFG1.Enabled = True
            VSFG1.Cell(flexcpForeColor, 1, 1, VSFG1.Rows - 1, 16) = &H0&
        End If
        If Me.optcliente.Value = True Then
            t = "C"
            txtFpag.Visible = True
            lblfp.Visible = True
            strSql = " SELECT for_pag_nombre " & _
                     " FROM forma_pago " & _
                     " INNER JOIN persona " & _
                     " ON persona.emp_codigo=forma_pago.emp_codigo " & _
                     " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
                     " WHERE forma_pago.emp_codigo='" & strEmpresa & "' " & _
                     " AND persona.per_codigo='" & dcmbBeneficiario.BoundText & "' "
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                txtFpag.Text = clsSql.adorec_Def(0)
            End If
        End If
    End If
    LlenarDatosDeDeposito
End Sub


Private Sub Form_Activate()
 Dim strComparar As String

'     consulta para saber los  bancos existentes
    strSql = " SELECT ban_codigo, ban_nombre " & _
             " FROM banco " & _
             " ORDER BY ban_codigo"
    clsBan.Ejecutar strSql
    dcmbBeneficiario.Tag = "NO"
    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        dcmbBanco = ""
    End If
    
    strSql = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSql

    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBancoE.RowSource = clsBan.adorec_Def.DataSource
        dcmbBancoE.ListField = "ban_nombre"
        dcmbBancoE.BoundColumn = "ban_codigo"
    Else
        dcmbBancoE = ""
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

Private Sub cargarGZDir()
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


'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa las clases para hacer distintas consultas
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsCod.Inicializar AdoConn, AdoConnMaster
    clsPgd.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsdoc.Inicializar AdoConn, AdoConnMaster
    clscdo.Inicializar AdoConn, AdoConnMaster
    clsRet.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsTip.Inicializar AdoConn, AdoConnMaster
    txtr = 0
    txtD = 0
    txtp = 0
    
    
    
    If dcmbBeneficiario.Text = "" Then
        cmdAceptar.Enabled = False
        VSFG1.Enabled = False
    End If
    
    
    strSql = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'C'" & _
             " ORDER BY tip_not_codigo"
    clsTip.Ejecutar strSql
    If clsTip.adorec_Def.EOF = False Then
    Set dcmbTipo.RowSource = clsTip.adorec_Def.DataSource
    dcmbTipo.ListField = "tip_not_nombre"
    dcmbTipo.BoundColumn = "tip_not_codigo"
'    dcmbTipo.Text = clsTip.adorec_Def("tip_not_nombre")
'    txtDescripciont.Text = clsTip.adorec_Def("descripcion")
    Else
        dcmbTipo.Text = ""
        dcmbTipo.BoundText = ""
        txtDescripcion = ""
    End If
    
    
    strSql = " SELECT ret_codigo, ret_nombre, ret_ctacontacli,ret_porcentaje" & _
                 " FROM retencion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " AND ret_ctacontacli!='' " & _
                 " AND ret_activo=1 " & _
                 " ORDER BY ret_codigo"
     clsSql.Ejecutar strSql

     VSFG1.ColComboList(12) = VSFG1.BuildComboList(clsSql.adorec_Def, "*ret_codigo,ret_porcentaje,ret_nombre", "ret_codigo")
    
    strSql = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsSql.Ejecutar strSql

     VSFG.ColComboList(5) = VSFG.BuildComboList(clsSql.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
     
     strSql = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSql

     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre")
     

' Asigna codigos de cuenta y nombres en el grid
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                 clsCta.Filtrar ("cta_codigo = '" & .TextMatrix(Row, 1) & "'")
                     .TextMatrix(Row, 2) = clsCta.adorec_Def("cta_nombre")
                 clsCta.QuitarFiltro
             End If

             If Col = 2 Then
                 clsCta.Filtrar ("cta_nombre = '" & .TextMatrix(Row, 2) & "'")
                     .TextMatrix(Row, 1) = clsCta.adorec_Def("cta_codigo")
                 clsCta.QuitarFiltro
             End If
         End If
    End With

    'Pone la fecha actual en los combos
    
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaCh.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaR.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaConta.Value = Format(HoyDia, "yyyy-MM-dd")
    'Seleccionamos el proveedor de la tabla persona (P), que esta por defecto
    cargarTipoPedido
    cargarGZDir
    
    OptCliente_Click
    'Consulta para saber tipod de documentos de pago
    dcmbDocumento.Enabled = False
    Me.dcmbBanco.Enabled = False
    strSql = " SELECT tip_doc_pag_codigo, tip_doc_pag_nombre " & _
             " FROM tipo_doc_pago "
    clsdoc.Ejecutar strSql
    
    Set dcmbDocumento.RowSource = clsdoc.adorec_Def.DataSource
    dcmbDocumento.ListField = "tip_doc_pag_nombre"
    dcmbDocumento.BoundColumn = "tip_doc_pag_codigo"
    
End Sub

Private Sub optcheque_Click()
    txtDocumento = ""
    dcmbBanco.Enabled = True
    dcmbDocumento.Enabled = True
    dcmbNota.Enabled = False
    
    dcmbDeudorCh = ""
    dcmbDeudorCh.Locked = True
'    dcmbCuenta.Enabled = True
'    txtTotalDebe = 0
'    txtTotalHaber = 0
'    VSFG.Clear 1
'    VSFG.Rows = 2
    
    'VSFG.TextMatrix(1, 3) = txtValor
    'VSFG.TextMatrix(FilaCxC, 3) = 0
    'VSFG.TextMatrix(FilaCxC, 4) = txtValor + txtValor.Tag
    pagos
    chkAnticipo.Enabled = True
    LlenarVariableDescripcion
    LlenarDatosDeDeposito
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaCh.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub OptCliente_Click()
  
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiario.Text = ""
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
    If cmbGerente.BoundText <> "-1" And cmbGerente.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref='" & cmbGerente.BoundText & "'"
    End If
    If cmbDirector.BoundText <> "-1" And cmbDirector.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref2='" & cmbDirector.BoundText & "'"
    End If
    If cmbEmprendedor.BoundText <> "-1" And cmbEmprendedor.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref3='" & cmbEmprendedor.BoundText & "'"
    End If
    If cmbEjecutivo.BoundText <> "-1" And cmbEjecutivo.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref4='" & cmbEjecutivo.BoundText & "'"
    End If
    If cmbN5.BoundText <> "-1" And cmbN5.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref5='" & cmbN5.BoundText & "'"
    End If
    If cmbN6.BoundText <> "-1" And cmbN6.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref6='" & cmbN6.BoundText & "'"
    End If
    If cmbN7.BoundText <> "-1" And cmbN7.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref7='" & cmbN7.BoundText & "'"
    End If
    If cmbN8.BoundText <> "-1" And cmbN8.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref8='" & cmbN8.BoundText & "'"
    End If
    If cmbN9.BoundText <> "-1" And cmbN9.BoundText <> "" Then
         strSql = strSql & " AND per_codigo_ref9='" & cmbN9.BoundText & "'"
    End If
    strSql = strSql & " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSql
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
        Set dcmbDeudorCh.RowSource = clsPer.adorec_Def.DataSource
        dcmbDeudorCh.ListField = "nombre"
        dcmbDeudorCh.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub optefectivo_Click()
    dcmbBanco.Enabled = False
    dcmbDocumento.Enabled = False
    dcmbDocumento = ""
    dcmbBanco = ""
    dcmbDeudorCh = ""
    dcmbDeudorCh.Locked = True
    dcmbNota.Enabled = False
'    dcmbCuenta.Enabled = False
'    dcmbCuenta = ""
'    txtDocumento.Enabled = True
'    txtDocumento = ""
'     VSFG.Clear 1
'    VSFG.Rows = 2
'     txtTotalDebe = 0
'    txtTotalHaber = 0
    If FilaCxC <= VSFG.Rows - 1 Then
        'VSFG.TextMatrix(1, 3) = txtValor
        'VSFG.TextMatrix(FilaCxC, 3) = 0
        'VSFG.TextMatrix(FilaCxC, 4) = txtValor
        pagos
        chkAnticipo.Enabled = True
        LlenarVariableDescripcion
    End If
    LlenarDatosDeDeposito
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaCh.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub optNCredito_Click()
    txtDocumento = ""
    dcmbBanco.Enabled = False
    dcmbDocumento.Enabled = False
    dcmbNota.Enabled = True
    
    dcmbDeudorCh = ""
    dcmbDeudorCh.Locked = True
    
    clsNot.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ing_codigo,ing_fecha,CONCAT(ing_serie,'-',ing_numero) as num,ing_saldo,ing_numasiento,ing_total-ing_saldo as sal " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='DCL' " & _
             " AND per_codigo IN (" & strPersona & ") " & _
             " AND ing_anulado=0 " & _
             " AND ing_total-ing_saldo!=0 " & _
             " ORDER BY ing_codigo"
    clsNot.Ejecutar strSql
    dcmbNota.ListField = "num"
    dcmbNota.BoundColumn = "ing_codigo"
    Set dcmbNota.RowSource = clsNot.adorec_Def.DataSource
    chkAnticipo.Value = 0
    chkAnticipo.Enabled = False
    LlenarVariableDescripcion
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaCh.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub optNDebito_Click()
    txtDocumento = ""
    dcmbBanco.Enabled = False
    dcmbDocumento.Enabled = False
    
    dcmbDeudorCh = ""
    dcmbDeudorCh.Locked = True
    
    VSFG.TextMatrix(1, 3) = 0
    VSFG.TextMatrix(FilaCxC, 3) = txtValor
    VSFG.TextMatrix(FilaCxC, 4) = 0
    chkAnticipo.Value = 0
    chkAnticipo.Enabled = False
    LlenarVariableDescripcion
End Sub

Private Sub optproveedor_Click()
    p = 1
    Frame1.Caption = "Proveedor"
    dcmbBeneficiario.Text = ""
    strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre) as nombre " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSql
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "nombre"
        dcmbBeneficiario.BoundColumn = "per_codigo"
        Set dcmbDeudorCh.RowSource = clsPer.adorec_Def.DataSource
        dcmbDeudorCh.ListField = "nombre"
        dcmbDeudorCh.BoundColumn = "per_codigo"
    End If
    dtpFecha.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaCh.Value = Format(HoyDia, "yyyy-mm-dd")
End Sub

Private Sub TextTotal_Change()
 TxtTotal = FormatoD2(txtTotalDebe - txtTotalHaber)
End Sub

Private Sub txtDocumento_Change()
    LlenarVariableDescripcion
End Sub

Private Sub txtTotalDebe_Change()
    txtTotalDebe = FormatoD2(txtTotalDebe)
     TxtTotal = FormatoD2(txtTotalDebe - txtTotalHaber)
End Sub

Private Sub txtTotalHaber_Change()
 txtTotalHaber = FormatoD2(txtTotalHaber)
  TxtTotal = FormatoD2(txtTotalDebe - txtTotalHaber)
End Sub


Private Sub txtValor_Change()
    If txtValor.Locked = True Then
        txtValor = FormatoD2(txtValor)
    End If
    'Poner el valor en la cuenta del cliente
    If FilaCxC < VSFG.Rows Then
        If optNDebito.Value = False Then
            VSFG.TextMatrix(FilaCxC, 3) = 0
            VSFG.TextMatrix(FilaCxC, 4) = FormatoD2(txtValor) + FormatoD2(txtValor.Tag)
        Else
            VSFG.TextMatrix(FilaCxC, 3) = txtValor
            VSFG.TextMatrix(FilaCxC, 4) = 0
        End If
    End If
    If optcheque.Value = True Or optefectivo.Value = True Then
        VSFG.TextMatrix(1, 3) = txtValor
    Else
        VSFG.TextMatrix(1, 3) = 0
    End If
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
    txtValor = FormatoD2(txtValor)
    VSFG.TextMatrix(1, 3) = txtValor
End Sub

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

    ' only interesetd in left button
    If Button <> 1 Then Exit Sub

    ' get cell that was clicked
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub

    If (c <> 0 Or r <= (lonNFijas - 1)) Then Exit Sub

    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - x
        If d > imgBtnDn.Width Then Exit Sub
        If r > lonNFijas - 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Cobros"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)

        'Recorro el FlexGrid para poner números a las filas

        If respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            PonerBotones
            CalcuTotal
        Else
            VSFG.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <> OldRow Then
        If Year(dtpFechaConta.Value) >= 2018 And VSFG.Rows > 1 Then
            If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes

    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            VSFG.AddItem ""
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
        If Row > 0 And Row < lonNFijas Then
'            If dcmbCuenta.Enabled = True Then
                If Col = 1 Then
                    Cancel = True
                End If
                If Col = 2 Then
                    Cancel = True
                End If
'            End If
            If Col = 3 Then
               Cancel = True
            End If
            If Col = 4 Then
                Cancel = True
            End If
        End If
    
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    'Verifica que se ingrese la cuenta contable en el grid
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
        MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
        VSFG.TextMatrix(Row, 3) = 0
        VSFG.TextMatrix(Row, 4) = 0
    ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe
        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = intDato
        End If
        'Verifica que solo se ingresen números tanto en el Debe como en el Haber
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = 0
        End If
        CalcuTotal
    End If
End Sub


Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    d = format(HoyDia, "yyyy-MM-dd")
'    dia = Mid(d, 9, 2)
'    mes = Mid(d, 6, 2)
'    año = Mid(d, 1, 4)
'    ffch = Format(cmbAño.Text + "-" + cmbMes + "-" + cmbDia.Text, "yyyy-mm-dd")
'    m = Mid(ffch, 6, 2)

'    If Val(cmbDia.Text) > dia Or m > mes Or Val(cmbAño.Text) > año Then
'            txtDisponible.Text = txtd
'    Else
'            txtDisponible.Text = Val(txtd) - Val(VSFG.TextMatrix(1, 4))
'    End If
'    txtPrevisto.Text = txtp - Val(VSFG.TextMatrix(1, 4))

If Row > lonNFijas - 1 And lonNFijas > 0 Then

'' Asigna codigos de cuenta y nombres en el grid
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                 clsCta.Filtrar ("cta_codigo = '" & .TextMatrix(Row, 1) & "'")
                     .TextMatrix(Row, 2) = clsCta.adorec_Def("cta_nombre")
                 clsCta.QuitarFiltro
             End If

             If Col = 2 Then
                 clsCta.Filtrar ("cta_nombre = '" & .TextMatrix(Row, 2) & "'")
                     .TextMatrix(Row, 1) = clsCta.adorec_Def("cta_codigo")
                 clsCta.QuitarFiltro
             End If
         End If
    End With
End If
CalcuTotal
End Sub

Private Sub VSFG_Validate(Cancel As Boolean)
    If Year(dtpFechaConta.Value) >= 2018 And VSFG.Rows > 1 Then
        If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim strComparar As String
If Col >= 12 And Row > 0 And VSFG1.TextMatrix(VSFG1.Row, 12) <> "" And VSFG1.TextMatrix(VSFG1.Row, 13) <> "" Then
    'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, se cree otra fila y ponga los botones correspondientes
    If VSFG1.Rows > Row + 1 Then
        If (VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row + 1, 0) And VSFG1.TextMatrix(Row + 1, 12) <> "" And VSFG1.TextMatrix(Row + 1, 13) <> "") Or VSFG1.TextMatrix(Row, 0) <> VSFG1.TextMatrix(Row + 1, 0) Then
            clsRet.adorec_Def.MoveFirst
            strComparar = "ret_codigo = '" & VSFG1.TextMatrix(Row, 12) & "'"
            clsRet.adorec_Def.Find strComparar
            If Not IsNull(clsRet.adorec_Def("ret_ctacontacli")) Then
                VSFG1.TextMatrix(Row, 14) = clsRet.adorec_Def("ret_ctacontacli")
                VSFG1.TextMatrix(Row, 15) = clsRet.adorec_Def("ret_porcentaje")
            Else
                MsgBox "no esta definida la Cuenta contable"
            End If
            If VSFG1.TextMatrix(Row, 0) <> VSFG1.TextMatrix(Row + 1, 0) Then
                VSFG1.AddItem VSFG1.TextMatrix(Row, 0) & vbTab & VSFG1.TextMatrix(Row, 1) & vbTab & VSFG1.TextMatrix(Row, 2) & vbTab & VSFG1.TextMatrix(Row, 3) & vbTab & VSFG1.TextMatrix(Row, 4) & vbTab & VSFG1.TextMatrix(Row, 5) & vbTab & VSFG1.TextMatrix(Row, 6) & vbTab & VSFG1.TextMatrix(Row, 7) & vbTab & VSFG1.TextMatrix(Row, 8) & vbTab & VSFG1.TextMatrix(Row, 9) & vbTab & VSFG1.TextMatrix(Row, 10) & vbTab & VSFG1.TextMatrix(Row, 11) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & VSFG1.TextMatrix(Row, 16), Row + 1
            End If
            'Une filas del grid
            VSFG1.MergeCells = flexMergeRestrictRows
            VSFG1.MergeCol(0) = True: VSFG1.MergeCol(1) = True: VSFG1.MergeCol(2) = True: VSFG1.MergeCol(3) = True: VSFG1.MergeCol(4) = True: VSFG1.MergeCol(5) = True: VSFG1.MergeCol(6) = True: VSFG1.MergeCol(7) = True: VSFG1.MergeCol(8) = True: VSFG1.MergeCol(9) = True: VSFG1.MergeCol(10) = True: VSFG1.MergeCol(11) = True: VSFG1.MergeCol(12) = False: VSFG1.MergeCol(13) = False: VSFG1.MergeCol(14) = False
        End If
    Else
        If Row + 1 = VSFG1.Rows Then
            clsRet.adorec_Def.MoveFirst
            strComparar = "ret_codigo = '" & VSFG1.TextMatrix(Row, 12) & "'"
            clsRet.adorec_Def.Find strComparar
            If Not IsNull(clsRet.adorec_Def("ret_ctacontacli")) Then
                VSFG1.TextMatrix(Row, 14) = clsRet.adorec_Def("ret_ctacontacli")
                VSFG1.TextMatrix(Row, 15) = clsRet.adorec_Def("ret_porcentaje")
            Else
                MsgBox "no esta definida la Cuenta contable"
            End If
            VSFG1.AddItem VSFG1.TextMatrix(Row, 0) & vbTab & VSFG1.TextMatrix(Row, 1) & vbTab & VSFG1.TextMatrix(Row, 2) & vbTab & VSFG1.TextMatrix(Row, 3) & vbTab & VSFG1.TextMatrix(Row, 4) & vbTab & VSFG1.TextMatrix(Row, 5) & vbTab & VSFG1.TextMatrix(Row, 6) & vbTab & VSFG1.TextMatrix(Row, 7) & vbTab & VSFG1.TextMatrix(Row, 8) & vbTab & VSFG1.TextMatrix(Row, 9) & vbTab & VSFG1.TextMatrix(Row, 10) & vbTab & VSFG1.TextMatrix(Row, 11) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & VSFG1.TextMatrix(Row, 16), Row + 1
            'Une filas del grid
            VSFG1.MergeCells = flexMergeRestrictRows
            VSFG1.MergeCol(0) = True: VSFG1.MergeCol(1) = True: VSFG1.MergeCol(2) = True: VSFG1.MergeCol(3) = True: VSFG1.MergeCol(4) = True: VSFG1.MergeCol(5) = True: VSFG1.MergeCol(6) = True: VSFG1.MergeCol(7) = True: VSFG1.MergeCol(8) = True: VSFG1.MergeCol(9) = True: VSFG1.MergeCol(10) = True: VSFG1.MergeCol(11) = True: VSFG1.MergeCol(12) = False: VSFG1.MergeCol(13) = False: VSFG1.MergeCol(14) = False
        End If
    End If
End If
If Col = 11 Then
    If VSFG1.TextMatrix(Row, 12) = "" Then
        If Row > 1 And Row < VSFG1.Rows - 1 Then
            If VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row - 1, 0) Or VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row + 1, 0) Then
                VSFG1.RemoveItem Row
            End If
        ElseIf Row = 1 Then
            If VSFG1.Rows > Row + 1 Then
                If VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row + 1, 0) Then
                    VSFG1.RemoveItem Row
                End If
            End If
        ElseIf Row = VSFG1.Rows - 1 Then
            If VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row - 1, 0) Then
                VSFG1.RemoveItem Row
            End If
        End If
    End If
End If
    If Col = 10 Then
        'Verifica que solo se ingresen números en el campo Debe
        If Not IsNumeric(VSFG1.TextMatrix(Row, 11)) And VSFG1.TextMatrix(Row, 11) <> "" Then
            MsgBox "Ingrese solo números en el Valor de Pago.", vbInformation, "Pagos"
            VSFG1.TextMatrix(Row, 11) = 0
        End If
    End If
    If Row < VSFG1.Rows Then
        If Val(VSFG1.TextMatrix(Row, 11)) > Val(VSFG1.TextMatrix(Row, 10)) And optNDebito.Value = False Then
            If MsgBox("El valor a pagar es mayor al Saldo." & vbNewLine & "Esta seguro de que el pago es mayor?", vbCritical + vbYesNo, "Pagos") = vbNo Then
                VSFG1.Select Row, 11
                VSFG1.TextMatrix(Row, 11) = 0
            End If
        End If
    End If
    pagos
End Sub

Private Sub VSFG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim clsSqlCierre As New clsConsulta
    clsSqlCierre.Inicializar AdoConn, AdoConnMaster
    If VSFG1.TextMatrix(Row, 1) = "0" Or VSFG1.TextMatrix(Row, 1) = "" Then
        If Col >= 11 Then
            Cancel = True
        End If
    ElseIf Abs(VSFG1.TextMatrix(Row, 1)) = 1 Then
        If Col >= 12 And Val(VSFG1.TextMatrix(Row, 16)) = 0 Then
            Cancel = True
        End If
    End If
    If Col = 11 Then
        If VSFG1.TextMatrix(Row, 12) <> "Tot.Ret." Then
        If (VSFG1.TextMatrix(Row, 12) <> "" Or VSFG1.TextMatrix(Row, 13) <> "") Then
            Cancel = True
            'MsgBox "No puede registrar un cobro ya que esta registrando una retencion"
        End If
        End If
    ElseIf Col = 13 Or Col = 12 Then
        If FormatoD2(VSFG1.TextMatrix(Row, 11)) <> 0 Then
            Cancel = True
            'MsgBox "No puede registrar una retencion ya que esta registrando un cobro"
        ElseIf Left(VSFG1.TextMatrix(Row, 6), 7) < Left(HoyDia, 7) Then
        
            strSql = " SELECT COALESCE(COUNT(*),0) " & _
                     " FROM cierre_mes " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cie_mes_ano=" & Year(VSFG1.TextMatrix(Row, 6)) & " AND cie_mes_mes=" & Month(VSFG1.TextMatrix(Row, 6))
            clsSqlCierre.Ejecutar strSql
            
            If clsSqlCierre.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsSqlCierre.adorec_Def(0)) > 0 Then
                    MsgBox "El mes está cerrado por contabilidad", vbInformation, "Fecha"
                    Cancel = True
                End If
            End If
        
            'If FormatoD0(Right(Left(VSFG1.TextMatrix(Row, 6), 7), 2)) - FormatoD0(Right(Left(HoyDia, 7), 2)) < -1 Then
                'Cancel = True
                'MsgBox "No puede registrar retenciones ya esta pasado la fecha de esta factura"
            'End If
        End If
    End If
  
End Sub

'Private Sub dcmbcuenta_KeyPress(KeyAscii As Integer)
'    'Validación de caracteres ingresados para que solo ingrese números y el caracter "/"
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 8) Then
'            KeyAscii = 0
'    End If
'End Sub
Private Sub VSFG1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 2 Or NewCol = 3 Or NewCol = 4 Or NewCol = 5 Or NewCol = 6 Or NewCol = 7 Or NewCol = 8 Or NewCol = 9 Or NewCol = 10 Then
        If NewCol > OldCol Then
            SendKeys vbKeyTab
        ElseIf NewCol < OldCol Then
            SendKeys vbKeyLeft
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 12 Then
            a = 1
        End If
    If Col = 11 Then
        txtValor = 0
    End If
    If Col = 1 And Row > 0 Then
        If Abs(VSFG1.TextMatrix(Row, 1)) = 1 Then
            VSFG1.Select Row, 1, Row, 13
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HC0FFFF
            VSFG1.Select Row, 11
        ElseIf Abs(VSFG1.TextMatrix(Row, 1)) = 0 Then
            VSFG1.Select Row, 1, Row, 13
            VSFG1.FillStyle = flexFillRepeat
            VSFG1.CellBackColor = &HFFFFFF
            VSFG1.Select Row, 11
            VSFG1.TextMatrix(Row, 11) = ""
            VSFG1.TextMatrix(Row, 12) = ""
            VSFG1.TextMatrix(Row, 13) = ""
            VSFG1.TextMatrix(Row, 14) = ""
            If Row < VSFG1.Rows - 2 And VSFG1.TextMatrix(Row, 0) <> " " Then
                While VSFG1.TextMatrix(Row, 0) = VSFG1.TextMatrix(Row + 1, 0)
                    VSFG1.RemoveItem Row + 1
                    If Row = VSFG1.Rows - 1 Then Exit Sub
                Wend
            End If
        End If
    End If
    
End Sub

Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (VSFG.TextMatrix(VSFG.Row, 3) = "") Then
                VSFG.TextMatrix(VSFG.Row, 3) = 0
     ElseIf VSFG.TextMatrix(VSFG.Row, 4) = "" Then
                VSFG.TextMatrix(VSFG.Row, 4) = 0
     End If
End Sub

Private Sub dcmbasiento_Change()
    If dcmbAsiento.Text = "" Then
        dcmbDescripcion.Text = ""
    Else
        clsAsi.Actualizar
        clsAsi.Filtrar "tip_asi_codigo = '" & dcmbAsiento.BoundText & "'"
            dcmbDescripcion.Tag = "A"
            dcmbDescripcion = clsAsi.adorec_Def("descripcionasi")
        clsAsi.QuitarFiltro
        dcmbDescripcion.Tag = ""
    End If
End Sub

Private Sub dcmbDescripcion_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If dcmbAsiento.Tag <> "A" Then
        If dcmbDescripcion.MatchedWithList = True Then
            dcmbAsiento.BoundText = dcmbDescripcion.BoundText
        End If
    End If
End Sub


Private Sub dcmbDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbAsiento.BoundText = dcmbDescripcion.BoundText
    End If
End Sub

Private Sub dcmbdescripcion_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbAsiento.BoundText = dcmbDescripcion.BoundText
End Sub
