VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNuevoIngreso 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Mercaderia"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "frmNuevoIngreso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   8880
      TabIndex        =   56
      Top             =   9360
      Width           =   1095
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
      Left            =   6840
      TabIndex        =   42
      Top             =   0
      Width           =   3255
      Begin NEED2.dtpFecha dtpFechaR 
         Height          =   285
         Left            =   1680
         TabIndex        =   58
         Top             =   263
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Value           =   42009.4060532407
      End
      Begin VB.TextBox txtAutorizacionR 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSerieR 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtDocumentoR 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   1335
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
         TabIndex        =   46
         Top             =   300
         Width           =   585
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
         TabIndex        =   45
         Top             =   630
         Width           =   375
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
         TabIndex        =   44
         Top             =   1350
         Width           =   1425
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
         TabIndex        =   43
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Extras del Documento"
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
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   9975
      Begin VB.CheckBox chkUnContenedor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Un Contenedor"
         Height          =   255
         Left            =   5760
         TabIndex        =   64
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton btn_cargar 
         Caption         =   "Cargar Ingresos de ""NS"""
         Height          =   375
         Left            =   3600
         TabIndex        =   61
         Top             =   1320
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo dcbo_ingreso 
         Height          =   315
         Left            =   1560
         TabIndex        =   60
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   57
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Value           =   42009.4060532407
      End
      Begin VB.TextBox txtDcto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCaduca 
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1060
      End
      Begin VB.TextBox txtNumAux 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   615
         Width           =   1815
      End
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtAutorizacion 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpCaduca 
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
         Left            =   8400
         TabIndex        =   8
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         CustomFormat    =   "MM/yyy"
         Format          =   66191363
         CurrentDate     =   37463
      End
      Begin MSDataListLib.DataCombo CmbFpago 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   315
         Left            =   7935
         TabIndex        =   52
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo dcmbIVA 
         Height          =   315
         Left            =   8280
         TabIndex        =   62
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "% IVA:"
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
         Left            =   7680
         TabIndex        =   63
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label Label16 
         BackColor       =   &H00DDDDDD&
         Caption         =   "No. Recepción"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblV 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   6960
         TabIndex        =   51
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
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
         Left            =   6960
         TabIndex        =   50
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Caducidad del Doc"
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
         Left            =   6960
         TabIndex        =   49
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label lblFpago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         TabIndex        =   47
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label3 
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
         Left            =   3840
         TabIndex        =   34
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizacion"
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
         Left            =   3840
         TabIndex        =   33
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label9 
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
         Left            =   3840
         TabIndex        =   32
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Doc"
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
         TabIndex        =   31
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Auxiliar"
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
         TabIndex        =   30
         Top             =   630
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Ingreso"
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
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   6615
      Begin MSDataListLib.DataCombo dcmbCodP 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Top             =   660
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cmbTDoc 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   300
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblCodProveedor 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Persona"
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
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   9975
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   315
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "Copiar Desde"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.Frame Frame4 
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
         Height          =   2535
         Left            =   1155
         TabIndex        =   35
         Top             =   3600
         Width           =   7935
         Begin VB.TextBox txtCantidad 
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtObserv 
            Height          =   285
            Left            =   360
            MaxLength       =   250
            TabIndex        =   21
            Top             =   2040
            Width           =   7335
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   600
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1680
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
            Left            =   6480
            TabIndex        =   17
            Top             =   840
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1080
            Width           =   1215
         End
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   1215
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
            Height          =   1095
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   4305
            _cx             =   29302186
            _cy             =   29296523
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmNuevoIngreso.frx":030A
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
         Begin VB.Label lblCantidad 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5580
            TabIndex        =   54
            Top             =   270
            Width           =   675
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
            Left            =   420
            TabIndex        =   41
            Top             =   1800
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Left            =   5580
            TabIndex        =   40
            Top             =   1710
            Width           =   450
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
            Left            =   5580
            TabIndex        =   39
            Top             =   1350
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Var SubT:"
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
            Left            =   5580
            TabIndex        =   38
            Top             =   870
            Width           =   735
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
            Left            =   5580
            TabIndex        =   37
            Top             =   1110
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal:"
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
            Left            =   5580
            TabIndex        =   36
            Top             =   630
            Width           =   630
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfgDetalle 
         Height          =   3330
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9735
         _cx             =   101991187
         _cy             =   101979890
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   275
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmNuevoIngreso.frx":038A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGAbrir 
         Height          =   1260
         Left            =   120
         TabIndex        =   55
         Top             =   4440
         Visible         =   0   'False
         Width           =   345
         _cx             =   29295201
         _cy             =   29296814
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
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmNuevoIngreso.frx":04C0
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
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   120
         Picture         =   "frmNuevoIngreso.frx":0560
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   360
         Picture         =   "frmNuevoIngreso.frx":0696
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   9390
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5070
      TabIndex        =   23
      Top             =   9390
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblObserv 
      BackColor       =   &H00BAA892&
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
      Height          =   225
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   1410
   End
End
Attribute VB_Name = "frmNuevoIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso de mercadería a los depòsitos por concepto de         #
'#  importaciones se permite crear estos ingresos                               #
'#  frmIngImportacion  V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite ingresar los productos a los diferentes depòsitos       #
'#  de la compañía por concepto de importaciones , solo se permite el ingreso   #
'#  de tales datos para posteriormente actualizar las existencias.              #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    ingreso    : En esta tabla se almacenan los nuevos ingresos de mercadería #
'#    det_ingreso: En estatabla se almacena los detalles de cada ingreso        #
'#    persona    : Se consulta los proveedores de la empresa                    #
'#    deposito   : Se consulta los depositos o bodegas de la empresa            #
'#    producto   : Se consulta los productos de la empresa                      #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#               limpiarFxGD()   Permite borrar los datos que se encuentran     #
'#                               en el flexGrid para realizar un nuevo ingreso  #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private clsCon_Prd As New clsConsulta
Private clsCon_TipDoc As New clsConsulta
Private clsRecargos As New clsConsulta
Private clsFPago As New clsConsulta
Private strSql As String
Private ModPreCos As Boolean
Private PreCos As String
Private IngAsi As Boolean
Private DctoTotal As Boolean
Private cargado As Boolean
Private cargadoIXC As Boolean
Private strNumeroIXC As String
Private strTipoIXC As String
'Variables Globales
Public i_ingreso As Integer



Private Sub btn_cargar_Click()
    Dim clsAux As New clsConsulta
    Dim i As Long
    
    clsAux.Inicializar AdoConn, AdoConnMaster
    vsfgDetalle.Clear 1
    vsfgDetalle.Rows = 2
    cargado = True
    
    strSql = " SELECT ing_codigo,for_pag_codigo,ing_fecha FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IXC' " & _
             " AND ing_factura='" & dcbo_ingreso.BoundText & "'"
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        cargadoIXC = True
        strNumeroIXC = clsAux.adorec_Def("ing_codigo")
        dtpFecha.value = clsAux.adorec_Def("ing_fecha")
        strTipoIXC = "IXC"
        Me.CmbFpago.BoundText = clsAux.adorec_Def("for_pag_codigo")
        strSql = " SELECT dep_codigo,prd_codigo,det_ing_precio as precio,det_ing_cantidad as cant " & _
                 " FROM det_ingreso " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND ing_codigo=" & strNumeroIXC & " AND tip_ing_codigo='" & strTipoIXC & "' "
    Else
        cargadoIXC = False
        strNumeroIXC = "0"
        strTipoIXC = ""
        strSql = " SELECT contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,0 as precio,SUM(det_con_mer_cantidad) as cant " & _
                 " FROM det_recepcion_mercaderia INNER JOIN contenedor_mercaderia " & _
                 " ON det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                 " INNER JOIN producto " & _
                 " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
                 " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo " & _
                 " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_contenedor_mercaderia.mov_codigo=" & strNumeroIXC & " AND det_contenedor_mercaderia.tip_mov_codigo='" & strTipoIXC & "' AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_ingreso.BoundText & "'" & _
                 " GROUP BY contenedor_mercaderia.dep_codigo,det_contenedor_mercaderia.prd_codigo,prd_nombre HAVING SUM(det_con_mer_cantidad)!=0 " & _
                 " ORDER BY prd_nombre "
    End If
    clsAux.Ejecutar strSql
    i = 1
    While Not clsAux.adorec_Def.EOF
        
        vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
        vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        
        vsfgDetalle.TextMatrix(i, 1) = clsAux.adorec_Def("dep_codigo")
        vsfgDetalle.TextMatrix(i, 2) = clsAux.adorec_Def("prd_codigo")
        vsfgDetalle.TextMatrix(i, 4) = clsAux.adorec_Def("cant")
        vsfgDetalle.TextMatrix(i, 5) = clsAux.adorec_Def("precio")
        i = i + 1
        clsAux.adorec_Def.MoveNext
    Wend
End Sub

Private Sub cmbTDoc_Change()
    clsCon_TipDoc.Filtrar "tip_ing_codigo='" & cmbTDoc.BoundText & "' "
    If dcmbCodP.Tag <> "S" Then
        CargarPersonas clsCon_TipDoc.adorec_Def("tip_ing_persona")
    End If
    
    If Val(clsCon_TipDoc.adorec_Def("tip_ing_retencion")) = 1 Then
        fraRet.Visible = True
        If Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) = "P" Then
            strSql = " SELECT TOP 1 COALESCE(com_ret_serie,'') as com_ret_serie,COALESCE(com_ret_numero,'0')+1 as com_ret_numero,COALESCE(com_ret_autorizacion,'') as com_ret_autorizacion " & _
                     " FROM comprobante_retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND " & IIf(GeneraDocElec = 1, "com_ret_serie='001" & PtoEmiDocEle & "' AND ", "") & _
                     " cue_p_c_tipo='" & Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) & "' " & _
                     " ORDER BY com_ret_numero DESC "
            clsCon_Def.Ejecutar strSql
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                txtSerieR.Text = clsCon_Def.adorec_Def("com_ret_serie")
                txtDocumentoR.Text = clsCon_Def.adorec_Def("com_ret_numero")
                txtAutorizacionR.Text = clsCon_Def.adorec_Def("com_ret_autorizacion")
                dtpFechaR.value = Format(HoyDia, "yyyy-mm-dd")
            End If
        End If
    Else
        fraRet.Visible = False
    End If
    If Right(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) = "S" Then
        IngAsi = True
    Else
        IngAsi = False
    End If
    If Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) <> "N" Then
        lblFpago.Visible = True
        CmbFpago.Visible = True
    Else
        lblFpago.Visible = False
        CmbFpago.Visible = False
    End If
    CmbFpago.Tag = Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1)
    If Right(clsCon_TipDoc.adorec_Def("tip_ing_cos_pre"), 1) = "S" Then
        ModPreCos = True
    Else
        ModPreCos = False
    End If
    PreCos = Left(clsCon_TipDoc.adorec_Def("tip_ing_cos_pre"), 1)
    CargaProductos
    If clsCon_TipDoc.adorec_Def("tip_ing_persona") = "N" Then
        dcmbCodP.BoundText = "%"
    End If
    If Val(clsCon_TipDoc.adorec_Def("tip_ing_impuesto")) = 1 Then
        TxtIva.Enabled = True
    Else
        TxtIva.Enabled = False
        TxtIva.Text = 0
    End If
    If Val(clsCon_TipDoc.adorec_Def("tip_ing_recargo")) = 1 Then
        VSFGReca.Enabled = True
    Else
        VSFGReca.Enabled = False
        TxtRecargo = 0
    End If
    If clsCon_TipDoc.adorec_Def("tip_ing_numsri") = "F" Then
        txtSerie.Text = strSucursal & strPtoFactura
        txtDocumento.Text = ""
        txtAutorizacion.Text = strAutorFactura
        txtCaduca.Text = ""
        txtSerie.Locked = True
        txtDocumento.Locked = True
        txtAutorizacion.Locked = True
        dtpCaduca.Enabled = False
    ElseIf clsCon_TipDoc.adorec_Def("tip_ing_numsri") = "P" Then
    
        If dcmbCodP.BoundText <> "" Then
            strSql = " SELECT tip_ped_ptofac " & _
                     " FROM tipo_pedido inner join persona on tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
                     " WHERE tipo_pedido.emp_codigo='" & strEmpresa & "' AND per_codigo='" & dcmbCodP.BoundText & "' "
        
            clsCon_Def.Ejecutar strSql
            txtSerie.Text = clsCon_Def.adorec_Def("tip_ped_ptofac") & strSucursal
        Else
        
            txtSerie.Text = strPtoFactura & strSucursal
        End If
        
        strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
                 " FROM ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
                 " AND tip_ing_codigo='DCL' AND ing_codigo like '" & FormatoD0(txtSerie.Text) & "%' AND LEN(ing_codigo)>10 " & _
                 " ORDER BY ing_numero DESC,ing_fecha DESC,ing_codigo DESC "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtSerie.Text = clsCon_Def.adorec_Def("ing_serie")
            'txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
            txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
            If clsCon_Def.adorec_Def("ing_caduca") <> "00/0000" Then
                If clsCon_Def.adorec_Def("ing_caduca") <> "" Then
                    dtpCaduca.value = clsCon_Def.adorec_Def("ing_caduca")
                End If
                txtCaduca.Text = clsCon_Def.adorec_Def("ing_caduca")
            Else
                dtpCaduca.value = Format(HoyDia, "mm\/yyyy")
                txtCaduca.Text = ""
            End If
        Else
            txtSerie.Text = ""
            txtCaduca.Text = ""
            txtDocumento.Text = ""
            txtAutorizacion.Text = ""
            dtpCaduca.value = Format(HoyDia, "mm\/yyyy")
        End If
        txtSerie.Locked = False
        txtDocumento.Locked = False
        txtAutorizacion.Locked = False
        dtpCaduca.Enabled = True
    End If
    If cmbTDoc.BoundText = "DCL" Then
        cmbVendedor.Visible = True
        lblV.Visible = True
    Else
        cmbVendedor.Visible = False
        lblV.Visible = False
    End If
End Sub

Private Sub cmdAbrir_Click()
    Dim num As Integer
    
    Dim strPath As String
    Dim strLinea As String
    Dim Arch As String
    'Arch = cmbTDoc.Text & ".xls"
    VSFGAbrir.Clear 1
    VSFGAbrir.Rows = 1
    
    If vsfgDetalle.Rows > 1 Then
        strPath = Trim(App.Path)
        cmdArchivo.DialogTitle = "Abrir"
        'cmdArchivo.DefaultExt = strPath
        cmdArchivo.InitDir = strPath
        'cmdArchivo.FileName = Arch
        cmdArchivo.Filter = "Documento de Excel 2003-2007|*.xls|Documento de Excel 2007|*xlsx|Todos los Archivos|*.*"
        cmdArchivo.ShowOpen
        num = FreeFile
        Archivo = cmdArchivo.FileName
        If Archivo <> "" Then
            VSFGAbrir.LoadGrid Archivo, flexFileExcel
            'VSFGAbrir.Visible = True
            'VSFGAbrir.Width = 400
            vsfgDetalle.Rows = 1
            With VSFGAbrir
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" Then
                        vsfgDetalle.AddItem "", i
                        vsfgDetalle.TextMatrix(i, 1) = .TextMatrix(i, 0)
                        vsfgDetalle.TextMatrix(i, 2) = .TextMatrix(i, 1)
                        vsfgDetalle.TextMatrix(i, 4) = .TextMatrix(i, 2)
                        vsfgDetalle.TextMatrix(i, 7) = .TextMatrix(i, 3)
                        vsfgDetalle.TextMatrix(i, 0) = i
                        vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                        vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
                    End If
                Next i
                If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = "" Then
                    vsfgDetalle.RemoveItem vsfgDetalle.Rows - 1
                End If
            End With
        End If
    Else
        MsgBox "No se tiene información para guardar", vbInformation, "Guardar"
    End If

End Sub

Private Sub cmdAceptar_Click()
    Dim clsIngreso As New clsInventario
    Dim clsIngresoAux As New clsInventario
    Dim clsAsiento As New clsContable
    Dim clsCta As New clsCtaXx
    Dim i As Long
    Dim strObserv As String
    Dim booGuardar As Boolean
    Dim TotalRet As Double
    Dim cue_p_c_codigo As Double
    Dim strTipCompAsiento As String
    Dim booSinIva As Boolean
    Dim NumeroAsiento As String
    booSinIva = False
    
    If clsCon_TipDoc.adorec_Def("tip_ing_persona") <> "N" Then
        If dcmbCodP.MatchedWithList = False Then
            MsgBox "Seleccione una persona", vbInformation, "Personas"
            Exit Sub
        End If
    End If
    If fraRet.Visible = True Then
        If Trim(txtSerieR.Text) = "" Or Trim(txtDocumentoR.Text) = "" Or Trim(txtAutorizacionR.Text) = "" Then
            MsgBox "Llene los campos de la Retencion", vbInformation, "Retenciones"
            Exit Sub
        End If
    End If
    If txtSerie.Locked = False And txtDocumento.Locked = False And txtAutorizacion.Locked = False Then
        If Trim(txtSerie.Text) = "" Or Trim(txtDocumento.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
            MsgBox "Llene los campos del Documento", vbInformation, "Documento"
            Exit Sub
        End If
    End If
    If CmbFpago.Text = "" And cmbTDoc.BoundText = "COM" Then
        MsgBox "Llene Forma de Pago", vbInformation, "Documento"
        Exit Sub
    End If
    If cmbTDoc.BoundText = "COM" Then
        If MsgBox("La compra ingresada grava IVA?", vbQuestion + vbYesNo, "IVA") = vbNo Then
            booSinIva = True
            TxtIva.Text = 0
            TxtTotal.Text = FormatoD2(TxtSubTotal.Text) + FormatoD2(TxtRecargo.Text) - FormatoD2(txtDcto.Text)
        End If
    End If
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    If cmbTDoc.BoundText = "DCL" Then
        
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido inner join persona on tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " WHERE tipo_pedido.emp_codigo='" & strEmpresa & "' AND per_codigo='" & dcmbCodP.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        
        txtSerie.Text = clsCon_Def.adorec_Def("tip_ped_ptofac") & strSucursal
        
        strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
                 " FROM ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
                 " AND tip_ing_codigo='DCL' AND ing_codigo like '" & FormatoD0(txtSerie.Text) & "%' AND LEN(ing_codigo)>10 " & _
                 " ORDER BY ing_numero DESC,ing_fecha DESC,ing_codigo DESC "
        clsCon_Def.Ejecutar strSql
        
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            txtSerie.Text = clsCon_Def.adorec_Def("ing_serie")
            txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
            txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
        Else
            txtDocumento.Text = 1
            txtAutorizacion.Text = "0"
        End If
        If chkUnContenedor.value = 1 Then
            strContenedorRecurrente = "111"
        Else
            strContenedorRecurrente = ""
        End If
        'NOTA DE CREDITO
        If GeneraDocElec = 1 Then
            booGuardar = clsIngreso.NuevoIng(True, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , dcmbIVA.BoundText, IIf(cargadoIXC = True, strNumeroIXC, ""))
        Else
            booGuardar = clsIngreso.NuevoIng(True, cmbTDoc.BoundText, True, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, txtNumAux.Text, , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , dcmbIVA.BoundText, IIf(cargadoIXC = True, strNumeroIXC, ""))
        End If
    ElseIf cmbTDoc.BoundText = "COM" Then
        If cargado = True Then
            booGuardar = clsIngreso.NuevoIng(False, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, , , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , dcmbIVA.BoundText, IIf(cargadoIXC = True, strNumeroIXC, ""))
        Else
            booGuardar = clsIngreso.NuevoIng(True, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, , , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , dcmbIVA.BoundText, IIf(cargadoIXC = True, strNumeroIXC, ""))
        End If
    Else
        'booGuardar = clsIngreso.NuevoIng(cmbTDoc.BoundText, False, strSucursal, strPtoFacturaOriginal, txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, , , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text))
        'Para que salga bien la serie
        If cargado = True Then
            booGuardar = clsIngreso.NuevoIng(False, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, , , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , , IIf(cargadoIXC = True, strNumeroIXC, ""))
        Else
            booGuardar = clsIngreso.NuevoIng(True, cmbTDoc.BoundText, False, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, dcmbCodP.BoundText, dtpFecha.value, , , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , , IIf(cargadoIXC = True, strNumeroIXC, ""))
        End If
    End If
    If booGuardar = True Then
        If clsIngreso.strTipo = "COM" Then
            strTipCompAsiento = "C"
        ElseIf clsIngreso.strTipo = "DCL" Then
            strTipCompAsiento = "A"
        End If
        strObserv = UCase(cmbTDoc.Text & clsIngreso.strDoc & vbNewLine & "PERSONA: " & dcmbCodP.Text & vbNewLine & "DOCUMENTO: " & txtSerie.Text & Format(txtDocumento.Text, "0000000") & vbNewLine & TxtObserv.Text & IIf(cargadoIXC = True, vbNewLine & "INGRESO x CONTABILIDAR " & strNumeroIXC, ""))
        If IngAsi = True Then
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            clsAsiento.NuevoAsiento strTipCompAsiento, dtpFecha.value, 0, 0, TxtTotal.Text, strObserv
            clsIngreso.ModificaIng , , , , , , clsAsiento.NumAsiento
            NumeroAsiento = clsAsiento.NumAsiento
        End If
        With vsfgDetalle
            For i = 1 To .Rows - 1
                clsIngreso.NuevoDetIng .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD4(.TextMatrix(i, 4)), FormatoD8(.TextMatrix(i, 5)), FormatoD4(.TextMatrix(i, 8)), FormatoD4(.TextMatrix(i, 6)), Abs(FormatoD0(.TextMatrix(i, 9)))
            Next i
            InicializarContenedorRecurrente
        End With
        If cargadoIXC = True Then
            clsIngresoAux.Inicializar AdoConn, AdoConnMaster
            clsIngresoAux.AnularIng strNumeroIXC, "IXC", , "CONTABILIZADO EN " & clsIngreso.strTipo & " " & clsIngreso.strDoc
        End If
        If cargado = True And cargadoIXC = False Then
            strSql = " UPDATE det_contenedor_mercaderia " & _
                     " SET det_contenedor_mercaderia.tip_mov_codigo='" & clsIngreso.strTipo & "'," & _
                     " det_contenedor_mercaderia.mov_codigo='" & clsIngreso.strDoc & "'" & _
                     " FROM det_recepcion_mercaderia ,contenedor_mercaderia,det_contenedor_mercaderia WHERE det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " AND contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo AND det_contenedor_mercaderia.mov_codigo=0 AND det_contenedor_mercaderia.tip_mov_codigo='' " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " AND det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_ingreso.BoundText & "'"
            clsCon_Def.Ejecutar strSql
        End If
        With VSFGReca
            For i = 1 To .Rows - 1
                clsIngreso.NuevoDetIngRecargo .TextMatrix(i, 1), FormatoD2(.TextMatrix(i, 3))
            Next i
        End With
        clsIngreso.DetRetenciones

        If IngAsi = True And CmbFpago.Visible = True Then
            clsFPago.adorec_Def.MoveFirst
            strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
            clsFPago.adorec_Def.Find strComparar
            'Inserta un nuevo registro de la cuenta por cobrar*/
            clsCta.Inicializar AdoConn, AdoConnMaster
            If CmbFpago.Tag = "P" Then
                clsCta.NuevaCta CmbFpago.Tag, 1, "06", dtpFecha.value, Format(IIf(CmbFpago.Visible = True, DateAdd("d", clsFPago.adorec_Def("for_pag_tiempo"), dtpFecha.value), dtpFecha.value), "yyyy-MM-dd"), dcmbCodP.BoundText, strObserv, txtSerie.Text, txtDocumento.Text, txtAutorizacion.Text, txtCaduca.Text, FormatoD2(clsIngreso.dblTotalProd), FormatoD2(clsIngreso.dblTotalServ), FormatoD2(IIf(booSinIva = True, 0, clsIngreso.dblTotalProdIVA)), FormatoD2(IIf(booSinIva = True, 0, clsIngreso.dblTotalServIVA)), dcmbIVA.BoundText, FormatoD2(clsIngreso.dblIVA), FormatoD2(clsIngreso.dblSubTotal0 + IIf(booSinIva = False, 0, clsIngreso.dblTotalProdIVA + clsIngreso.dblTotalServIVA)), 0, 0, 0, FormatoD2(clsIngreso.dblTotal), clsAsiento.NumAsiento
                strSql = " BEGIN TRAN "
                clsCon_Def.Ejecutar strSql, "M"
                strSql = " SELECT TOP 1 COALESCE(com_ret_serie,'') as com_ret_serie,COALESCE(com_ret_numero,'0')+1 as com_ret_numero,COALESCE(com_ret_autorizacion,'') as com_ret_autorizacion " & _
                         " FROM comprobante_retencion WITH (TABLOCKX) " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND " & IIf(GeneraDocElec = 1, "com_ret_serie='001" & PtoEmiDocEle & "' AND ", "") & " cue_p_c_tipo='" & Left(clsCon_TipDoc.adorec_Def("tip_ing_cx_p_c"), 1) & "' " & _
                         " ORDER BY com_ret_numero DESC "
                clsCon_Def.Ejecutar strSql, "M"
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    txtSerieR.Text = clsCon_Def.adorec_Def("com_ret_serie")
                    txtDocumentoR.Text = clsCon_Def.adorec_Def("com_ret_numero")
                    txtAutorizacionR.Text = clsCon_Def.adorec_Def("com_ret_autorizacion")
                    'dtpFechaR.Value = Format(HoyDia, "yyyy-mm-dd")
                End If
                
                clsCta.IngRetencionPersonaIng clsIngreso, txtSerieR.Text, txtDocumentoR.Text, txtAutorizacionR.Text, Format(dtpFechaR.value, "yyyy-mm-dd")
                
                clsCta.IngAsientoIng clsAsiento, clsIngreso
                DocElectronico "07", (clsCta.strNoCta)
                MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
                
                'Setear informacion de la tabla Intermediate
'                Call Actualizar(i_ingreso)
                
                If fraRet.Visible = True Then
                    clsCta.VerRet
                End If
            ElseIf CmbFpago.Tag = "S" Then
                clsCta.IngAsientoIng clsAsiento, clsIngreso
                MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
            End If
        End If
        
        If clsIngreso.strTipo = "DCL" Then
            DocElectronico "04", (clsIngreso.strDoc)
            'Aplica la Nota de Credito
            clsCta.Inicializar AdoConn, AdoConnMaster
            clsCta.strPersona = clsIngreso.strPersona
            clsCta.strTipoCta = "C"
            clsCta.AplicaNC clsIngreso.strDoc, "2011-01-01", dtpFecha.value
        End If
        
        If clsIngreso.strTipo = "DCL" Then
            Dim rpTNC As New frmReporte
            rpTNC.strNumero = clsIngreso.strDoc
            rpTNC.strReporte = "rptNotaCredito"
            rpTNC.Show
            
            
            Dim rpTNC2 As New frmReporte
            rpTNC2.strNumero = clsIngreso.strDoc
            rpTNC2.strReporte = "rptNotaCreditoUbicacion"
            rpTNC2.Show
            
            Dim rpMov1 As New frmReporte
            rpMov1.strNumero = clsIngreso.strDoc
            rpMov1.strTipo = clsIngreso.strTipo
            rpMov1.strReporte = "rptDetalleAdjunto"
            rpMov1.Show

        Else
            Dim rpMov As New frmReporte
            rpMov.strNumero = clsIngreso.strDoc
            rpMov.strTipo = clsIngreso.strTipo
            rpMov.strReporte = "rptIngresoMercaderia"
            rpMov.Show
            

        End If
        If clsIngreso.strTipo = "COM" Then
            Dim rpAsi As New frmReporte
            rpAsi.strAsiento = NumeroAsiento
            rpAsi.strNumero = clsCta.strNoCta
            rpAsi.strTipo = "P"
            rpAsi.strReporte = "rptRetencionDiario"
            rpAsi.Show
            
            
        End If
        If cargado = True Then
            strSql = " UPDATE det_recepcion_mercaderia ,contenedor_mercaderia,det_contenedor_mercaderia " & _
                     " SET det_recepcion_mercaderia.tip_mov_codigo='" & clsIngreso.strTipo & "', " & _
                     " det_recepcion_mercaderia.mov_codigo='" & clsIngreso.strDoc & "' " & _
                     " WHERE det_recepcion_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_recepcion_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " AND contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo_origen=0 " & _
                     " AND det_recepcion_mercaderia.mov_codigo=0" & _
                     " AND det_recepcion_mercaderia.tip_mov_codigo=''" & _
                     " WHERE det_recepcion_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_recepcion_mercaderia.rec_mer_codigo LIKE  '" & dcbo_ingreso.BoundText & "%'"
            'clsCon_Def.Ejecutar strSql, "M"
        End If
        Unload Me
    End If
    Set clsCta = Nothing
    Set clsAsiento = Nothing
End Sub

Private Sub Actualizar(p_actualizar As Integer)
    '
    ' Si ya esta abierta la conexion la setea.
    Set cnn = Nothing
    Set rst = Nothing
    '
    ' Crear los objetos
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    If Trim(s_instanciaSQL) <> "" Then
        
        With cnn
            .CursorLocation = adUseClient
            .Open cadena_conexion
        End With
    
        ' abrir el recordset indicando la tabla a la que queremos acceder
        rst.Open "UPDATE ingreso set ing_estado=" & 2 & " where ing_id = " & p_actualizar, cnn, adOpenDynamic, adLockOptimistic
    End If
End Sub


Private Sub cmdCopiar_Click()
    frmVerMovimiento.Tag = "I"
    frmVerMovimiento.cmdCopiar.Visible = True
    frmVerMovimiento.cmdAnular.Visible = False
End Sub


Private Sub dcbo_ingreso_Validate(Cancel As Boolean)
 If (dcbo_ingreso.MatchedWithList = True) Then
     i_ingreso = CInt(dcbo_ingreso.BoundText)
 End If
End Sub

Private Sub dcmbCodP_Validate(Cancel As Boolean)
    dcmbCodP.Tag = "S"
    cargado = False
    chkUnContenedor.value = 0
    chkUnContenedor.Visible = False
    strSql = " SELECT rec_mer_codigo,CONCAT(rec_mer_codigo,' ()') as n " & _
             " FROM recepcion_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND per_codigo like '" & dcmbCodP.BoundText & "%' " & _
             " AND rec_mer_codigo NOT IN (" & _
             " SELECT DISTINCT ing_factura " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='IXC' AND ing_anulado=0" & _
             " ) " & _
             " UNION" & _
             " SELECT rec_mer_codigo,CONCAT(rec_mer_codigo,' (',ing_codigo,')') as n " & _
             " FROM recepcion_mercaderia INNER JOIN ingreso " & _
             " ON recepcion_mercaderia.emp_codigo=ingreso.emp_codigo " & _
             " AND recepcion_mercaderia.per_codigo=ingreso.per_codigo " & _
             " AND recepcion_mercaderia.rec_mer_codigo=ingreso.ing_factura " & _
             " AND ingreso.tip_ing_codigo='IXC' and ing_anulado=0 " & _
             " WHERE recepcion_mercaderia.emp_codigo='" & strEmpresa & "' " & _
             " AND recepcion_mercaderia.per_codigo like '" & dcmbCodP.BoundText & "%' " & _
             " ORDER BY rec_mer_codigo"
    clsCon_Def.Ejecutar strSql
    
    Set dcbo_ingreso.RowSource = clsCon_Def.adorec_Def.DataSource
    dcbo_ingreso.ListField = "n"
    dcbo_ingreso.BoundColumn = "rec_mer_codigo"
    
    
    If chkUnContenedor.value = 1 And cmbTDoc.BoundText = "DCL" Then
        strContenedorRecurrente = "111"
    Else
        strContenedorRecurrente = ""
    End If
    
    If cmbTDoc.BoundText = "DCL" Or cmbTDoc.BoundText = "COM" Then
        Dim Gerente As String, Director As String, entra As Boolean, tipNeg As String
        entra = False: Gerente = "": Director = ""
        Dim Vend As String, ForPag As String
        
        If dcmbCodP.BoundText <> "" Then
            
        
            strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,tip_ped_codigo,per_codigo_ref,per_codigo_ref2,for_pag_codigo,ven_codigo " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_codigo='" & dcmbCodP.BoundText & "' "
            clsCon_Def.Ejecutar strSql
            
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
            
                tipNeg = clsCon_Def.adorec_Def(1)
                Gerente = clsCon_Def.adorec_Def(2)
                Director = clsCon_Def.adorec_Def(3)
                Vend = clsCon_Def.adorec_Def(5)
                ForPag = clsCon_Def.adorec_Def(4)
                
                
    
                
                If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                    MsgBox "El Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                    'cmdAceptar.Enabled = False
                    'entra = True
                End If
                
                
                strSql = " SELECT tip_ped_ptofac " & _
                         " FROM tipo_pedido inner join persona on tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
                         " WHERE tipo_pedido.emp_codigo='" & strEmpresa & "' AND per_codigo='" & dcmbCodP.BoundText & "' "
                clsCon_Def.Ejecutar strSql
                
                txtSerie.Text = clsCon_Def.adorec_Def("tip_ped_ptofac") & strSucursal
                
                strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
                         " FROM ingreso " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
                         " AND tip_ing_codigo='" & cmbTDoc.BoundText & "' AND ing_codigo like '" & FormatoD0(txtSerie.Text) & "%' " & _
                         " AND LEN(ing_codigo)>10 AND ing_codigo NOT IN ( " & _
                            " SELECT i_s.ing_codigo FROM ingreso_salto i_s " & _
                            " WHERE i_s.emp_codigo='" & strEmpresa & "' AND i_s.tip_ing_codigo='" & cmbTDoc.BoundText & "' " & _
                            " AND i_s.ing_codigo like '" & FormatoD0(txtSerie.Text) & "%')" & _
                         " ORDER BY ing_numero DESC,ing_fecha DESC,ing_codigo DESC "
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
                    txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
                End If
            End If
            If entra = False Then
                strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND per_codigo='" & Gerente & "' " & _
                         " AND tip_ped_codigo='" & tipNeg & "' "
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                        MsgBox "El Gerente de Zona del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                        'cmdAceptar.Enabled = False
                        'entra = True
                    End If
                End If
            End If
            
            If entra = False Then
                strSql = " SELECT IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                         " FROM persona " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cat_p_tipo='C' " & _
                         " AND per_codigo='" & Director & "' " & _
                         " AND tip_ped_codigo='" & tipNeg & "' "
                clsCon_Def.Ejecutar strSql
                If clsCon_Def.adorec_Def.RecordCount > 0 Then
                    If FormatoD0(clsCon_Def.adorec_Def(0)) = 1 Then
                        MsgBox "El Director del Cliente está BLOQUEADO." & vbNewLine & vbNewLine & "Va a continuar con la Transaccion", vbCritical, "Bloqueado"
                        'cmdAceptar.Enabled = False
                        'entra = True
                    End If
                End If
            End If
            
            If cmbTDoc.BoundText = "DCL" Then
                cmbVendedor.BoundText = Vend
                CmbFpago.BoundText = ForPag
            End If
            
        End If
    End If
    If entra = False Then
        cmbTDoc_Change
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = True
    End If
    dcmbCodP.Tag = ""
    chkUnContenedor.value = 0
    chkUnContenedor.Visible = False
End Sub

Private Sub dcmbIVA_Validate(Cancel As Boolean)
    TxtIva.Tag = dcmbIVA.Text
    CalculaTotal
End Sub

Private Sub dtpCaduca_Change()
    txtCaduca.Text = Format(dtpCaduca.value, "mm\/yyyy")
End Sub

Private Sub dtpFecha_Change()
    dtpFechaR.value = dtpFecha.value
End Sub

Private Sub Form_Activate()
    CargaProductos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 And cmbTDoc.BoundText = "DCL" Then
        If chkUnContenedor.Visible = True Then
            chkUnContenedor.Visible = False
        Else
            chkUnContenedor.Visible = True
        End If
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_TipDoc = Nothing
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim CadenaConexion As String
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_TipDoc.Inicializar AdoConn, AdoConnMaster
    clsRecargos.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster

    Dim au As Integer
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " WHERE cod_iva_enuso=1 "
     clsCon_Def.Ejecutar (strSql)
     au = clsCon_Def.adorec_Def("cod_iva_codigo")
     
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " ORDER BY cod_iva_porcentaje"
     clsCon_Def.Ejecutar (strSql)
     Set dcmbIVA.RowSource = clsCon_Def.adorec_Def.DataSource
     dcmbIVA.ListField = "cod_iva_porcentaje"
     dcmbIVA.BoundColumn = "cod_iva_codigo"
     dcmbIVA.BoundText = au

    dtpFecha.value = HoyDia
    dtpFechaR.value = HoyDia
    DctoTotal = False
        'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0

    'IVA
    strSql = " SELECT par_numero " & _
             " FROM parametro WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo='IVAC' "
    clsCon_TipDoc.Ejecutar strSql
    TxtIva.Tag = clsCon_TipDoc.adorec_Def("par_numero")
    'Carga los ingresos
    strSql = " SELECT tip_ing_codigo, tip_ing_nombre,tip_ing_impuesto,tip_ing_persona,tip_ing_cx_p_c,tip_ing_recargo,tip_ing_numsri,tip_ing_cos_pre,tip_ing_retencion " & _
             " FROM tipo_ingreso WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_TipDoc.Ejecutar strSql
    cmbTDoc.ListField = "tip_ing_nombre"
    cmbTDoc.BoundColumn = "tip_ing_codigo"
    Set cmbTDoc.RowSource = clsCon_TipDoc.adorec_Def.DataSource
    'Carga los depositos
    strSql = "SELECT dep_codigo, dep_nombre FROM deposito WHERE emp_codigo = '" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql
    vsfgDetalle.ColComboList(1) = vsfgDetalle.BuildComboList(clsCon_Def.adorec_Def, "*dep_codigo, dep_nombre", "dep_codigo")
    
    '****** VENDEDORES
    'Coloca los datos de los vendedores en un listado
    strSql = " SELECT ven_codigo, CONCAT(ven_apellido,' ',ven_nombre) as nombV " & _
             " FROM vendedor " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY nombV "
    clsCon_Def.Ejecutar (strSql)
    Set cmbVendedor.RowSource = clsCon_Def.adorec_Def.DataSource
    cmbVendedor.ListField = "nombV"
    cmbVendedor.BoundColumn = "ven_codigo"
    
    
    
    'Consulta los recargos que puede manejar una empresa
    strSql = " SELECT oca_codigo,oca_nombre,oca_precio " & _
             " FROM ocargos " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY oca_nombre "
    clsRecargos.Ejecutar (strSql)
    'Muestra los recargos en el combo del grid de recargos
    VSFGReca.ColComboList(1) = VSFGReca.BuildComboList(clsRecargos.adorec_Def, "*oca_codigo,oca_nombre")
    
    'Insertamos el botón de eliminar en cada una de las filas
    vsfgDetalle.Cell(flexcpPicture, 1, 0) = imgBtnUp
    vsfgDetalle.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsFPago.Ejecutar (strSql)
    Set CmbFpago.RowSource = clsFPago.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        End Select
End Sub

Private Sub TxtDesc_Validate(Cancel As Boolean)
    Dim i As Long
    Dim Dcto As Double
    Dim subST As Double
    DctoTotal = True
    subST = 0
    Dcto = FormatoD2(TxtDesc.Text)
    For i = 1 To vsfgDetalle.Rows - 1
        If FormatoD2(subST - FormatoD2(TxtSubTotal.Text)) = 0 Then Exit For
        vsfgDetalle.TextMatrix(i, 6) = FormatoD4((FormatoD2(vsfgDetalle.TextMatrix(i, 4)) * FormatoD2(vsfgDetalle.TextMatrix(i, 5))) / FormatoD2(FormatoD2(TxtSubTotal.Text) - subST) * FormatoD4(Dcto))
        'Dcto = Dcto - vsfgDetalle.TextMatrix(i, 6)
        'subST = subST + FormatoD2(vsfgDetalle.TextMatrix(i, 7))
    Next i
    DctoTotal = False
End Sub

Private Sub txtNumAux_Change()
    dcmbIVA.BoundText = RevisivaCodigoIVAFactura(txtNumAux.Text)
    TxtIva.Tag = dcmbIVA.Text
    CalculaTotal
End Sub

Private Sub vsfgDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col >= 8 Then
        Cancel = True
    End If
    If ModPreCos = False Then
        If Col = 5 Or Col = 6 Or Col = 7 Then
            Cancel = True
        End If
    End If
    If cargado = True And Not (Col = 5 Or Col = 7) Then
        Cancel = True
    End If
End Sub

Private Sub vsfgDetalle_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalle.MouseRow
    c = vsfgDetalle.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = 1) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalle.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
   
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalle.Cell(flexcpLeft, r, c) + vsfgDetalle.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "AHORA DEBE ELIMINAR ESTA FILA!"
    
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Ingreso de Importación"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        
        'Recorro el FlexGrid para almacenar los detalles del ingreso
        
        If respuesta = vbYes Then
            Dim i As Long
            vsfgDetalle.RemoveItem (r)
            For i = 1 To (vsfgDetalle.Rows - 1)
                vsfgDetalle.TextMatrix(i, 0) = i
                vsfgDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
                vsfgDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
            Next i
            CalculaTotal
        Else
            vsfgDetalle.Cell(flexcpPicture, r, c) = imgBtnUp
        
        End If
    Cancel = True
End Sub

Private Sub vsfgDetalle_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalle.TextMatrix(Row, 2) & "'"
        vsfgDetalle.ShowCell Row, 2
        If clsCon_Prd.adorec_Def.RecordCount > 0 Then
            vsfgDetalle.TextMatrix(Row, 3) = clsCon_Prd.adorec_Def("prd_nombre")
            vsfgDetalle.TextMatrix(Row, 4) = 0
            vsfgDetalle.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_precio")
            vsfgDetalle.TextMatrix(Row, 6) = 0
            vsfgDetalle.TextMatrix(Row, 7) = 0
            vsfgDetalle.TextMatrix(Row, 8) = clsCon_Prd.adorec_Def("prd_costo")
            vsfgDetalle.TextMatrix(Row, 9) = clsCon_Prd.adorec_Def("prd_iva")
        Else
            MsgBox "El producto " & vsfgDetalle.TextMatrix(Row, 2) & " no esta en la base", vbCritical, "Productos"
            cmdAceptar.Enabled = False
        End If
        If Me.cmbTDoc.BoundText = "DCL" Then
            vsfgDetalle.TextMatrix(Row, 1) = clsCon_Prd.adorec_Def("dep_codigo")
        End If
    ElseIf Col = 4 Or Col = 5 Or Col = 6 Then
        If FormatoD2(txtDcto.Text) <> 0 And Col <> 6 Then
            vsfgDetalle.TextMatrix(Row, 6) = FormatoD4(vsfgDetalle.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalle.TextMatrix(Row, 5)) * FormatoD2(txtDcto.Text) / 100
        End If
        vsfgDetalle.TextMatrix(Row, 7) = FormatoD2(FormatoD4(vsfgDetalle.TextMatrix(Row, 4)) * FormatoD4(vsfgDetalle.TextMatrix(Row, 5)) - FormatoD4(vsfgDetalle.TextMatrix(Row, 6)))
        CalculaTotal
    ElseIf Col = 7 Then
        If FormatoD4(vsfgDetalle.TextMatrix(Row, 4)) <> 0 Then
            vsfgDetalle.TextMatrix(Row, 5) = FormatoD8((FormatoD2(vsfgDetalle.TextMatrix(Row, 7)) + FormatoD4(vsfgDetalle.TextMatrix(Row, 6))) / FormatoD4(vsfgDetalle.TextMatrix(Row, 4)))
        Else
            vsfgDetalle.TextMatrix(Row, 5) = 0
        End If
        CalculaTotal
    End If
    If vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 2) <> "" And vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 3) <> "" And Val(vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 4)) <> 0 Then
        vsfgDetalle.AddItem ""
        vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 0) = vsfgDetalle.Rows - 1
        vsfgDetalle.Cell(flexcpPicture, vsfgDetalle.Rows - 1, 0) = imgBtnUp
        vsfgDetalle.Cell(flexcpPictureAlignment, vsfgDetalle.Rows - 1, 0) = flexAlignRightCenter
        If vsfgDetalle.Rows > 2 Then
            vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 1, 1) = vsfgDetalle.TextMatrix(vsfgDetalle.Rows - 2, 1)
        End If
    End If
    
End Sub

Private Sub CalculaCant()
    Dim i As Long, Total As Double
    Total = 0
    For i = 1 To vsfgDetalle.Rows - 1
        Total = Total + FormatoD4(vsfgDetalle.TextMatrix(i, 4))
    Next i
    txtCantidad.Text = FormatoD4(Total)
End Sub

Private Sub CalculaTotal()
    Dim i As Long
    Dim subIVA As Double
    Dim Suma As Double
    Dim subIVASDcto As Double
    Dim sumaSDcto As Double
    Suma = 0
    subIVA = 0
    sumaSDcto = 0
    subIVASDcto = 0
    If DctoTotal = False Then
        TxtDesc.Text = 0
    End If
    CalculaCant
    CalcuReca
    For i = 1 To vsfgDetalle.Rows - 1
        If Abs(FormatoD0(vsfgDetalle.TextMatrix(i, 9))) = 0 Then
            subIVA = FormatoD2(FormatoD2(subIVA) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)))
            subIVASDcto = FormatoD2(FormatoD2(subIVASDcto) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)) + FormatoD4(vsfgDetalle.TextMatrix(i, 6)))
        Else
            Suma = FormatoD2(FormatoD2(Suma) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)))
            sumaSDcto = FormatoD2(FormatoD2(sumaSDcto) + FormatoD2(vsfgDetalle.TextMatrix(i, 7)) + FormatoD4(vsfgDetalle.TextMatrix(i, 6)))
        End If
        If DctoTotal = False Then
            TxtDesc.Text = FormatoD2(FormatoD2(TxtDesc.Text) + FormatoD4(vsfgDetalle.TextMatrix(i, 6)))
        End If
    Next i
    TxtSubTotal.Text = sumaSDcto
    TxtRecargo.Tag = FormatoD2(FormatoD2(TxtRecargo.Text) + subIVA)
    TxtRecargo.Text = FormatoD2(FormatoD2(TxtRecargo.Text) + subIVASDcto)
    If TxtIva.Enabled = True Then
        TxtIva.Text = FormatoD2(FormatoD2(Suma) * Val(TxtIva.Tag) / 100#)
    End If
    TxtTotal.Text = FormatoD2(Suma) + FormatoD2(TxtIva.Text) + FormatoD2(TxtRecargo.Tag)
End Sub

Private Sub CargarPersonas(Tipo As String)
'Carga los personas
    If Tipo = "C" Then
        strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,' - ',tip_ped_codigo,' - ',per_ruc,')') per " & _
                 " FROM persona " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND cat_p_tipo LIKE '" & Tipo & "'" & _
                 " AND per_inactivo=0 " & _
                 " ORDER BY CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,' - ',tip_ped_codigo,' - ',per_ruc,')')"
    Else
        strSql = " SELECT per_codigo, CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,' - ',per_ruc,')') per " & _
                 " FROM persona " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                 " AND cat_p_tipo LIKE '" & Tipo & "'" & _
                 " AND per_inactivo=0 " & _
                 " ORDER BY CONCAT(per_apellido,' ',per_nombre,' (',cat_p_tipo,' - ',per_ruc,')')"
    End If
    clsCon_Def.Ejecutar strSql
    dcmbCodP.ListField = "per"
    dcmbCodP.BoundColumn = "per_codigo"
    Set dcmbCodP.RowSource = clsCon_Def.adorec_Def.DataSource
End Sub
Private Sub VSFGReca_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Aumenta una fila adicional en el grid de recargos en caso de ser necesario
    If OldCol = 2 And OldRow = VSFGReca.Rows - 1 And NewCol = 3 And VSFGReca.TextMatrix(OldRow, 1) <> "" Then
        VSFGReca.AddItem ""
        PonerBotones
    End If
End Sub

Private Sub VSFGReca_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 0 del recargo
    If Col = 2 Then
        Cancel = True
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
             CalculaTotal
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
    CalculaTotal
End Sub

Private Sub VSFGReca_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 And (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
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
    TxtRecargo = Format(Suma, "####0.00")
End Sub

Private Sub CargaProductos()
    strSql = " SELECT COALESCE(per_dcto,0) as per_dcto " & _
             " FROM persona " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND per_codigo='" & dcmbCodP.BoundText & "' "
    clsCon_Prd.Ejecutar strSql
    If clsCon_Prd.adorec_Def.RecordCount > 0 Then
        txtDcto.Text = FormatoD2(clsCon_Prd.adorec_Def("per_dcto"))
    Else
        txtDcto.Text = 0
    End If
    If PreCos = "C" Then
        'Carga los productos
        strSql = " SELECT producto.prd_codigo, prd_nombre,prd_costo,COALESCE(per_pro_precio,prd_costo) as prd_precio,prd_iva,coleccion.dep_codigo " & _
                 " FROM producto " & _
                 " INNER JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo AND producto.clc_codigo=coleccion.clc_codigo" & _
                 " LEFT JOIN persona_producto ON producto.emp_codigo=persona_producto.emp_codigo" & _
                 " AND producto.prd_codigo=persona_producto.prd_codigo AND persona_producto.per_codigo='" & dcmbCodP.BoundText & "'" & _
                 " WHERE producto.emp_codigo = '" & strEmpresa & "' " & _
                 " ORDER BY producto.emp_codigo,prd_codigo "
    Else
        'Carga los productos
        strSql = " SELECT producto.prd_codigo, prd_nombre,prd_costo,lis_pre_p_precio as prd_precio,prd_iva,coleccion.dep_codigo " & _
                 " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                 " INNER JOIN lista_precio_p ON categoria_p.emp_codigo=lista_precio_p.emp_codigo AND categoria_p.lis_pre_codigo=lista_precio_p.lis_pre_codigo " & _
                 " INNER JOIN producto ON lista_precio_p.emp_codigo=producto.emp_codigo AND lista_precio_p.prd_codigo=producto.prd_codigo AND producto.prd_baja=0 " & _
                 " INNER JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo AND producto.clc_codigo=coleccion.clc_codigo" & _
                 " WHERE persona.emp_codigo = '" & strEmpresa & "' " & _
                 " AND persona.per_codigo = '" & dcmbCodP.BoundText & "' " & _
                 " ORDER BY prd_codigo "
    End If
    clsCon_Prd.Ejecutar strSql
End Sub

Private Sub cmbProducto_Validate(Cancel As Boolean)
    vsfgDetalle.TextMatrix(vsfgDetalle.Row, 2) = cmbProducto.BoundText
    cmbProducto.Visible = False
    vsfgDetalle.SetFocus
    vsfgDetalle.Col = 2
    vsfgDetalle.EditCell
End Sub


Private Sub vsfgDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If vsfgDetalle.Col = 3 And KeyCode = vbKeyF4 And Trim(vsfgDetalle.TextMatrix(vsfgDetalle.Row, vsfgDetalle.Col)) <> "" And Len(Trim(vsfgDetalle.TextMatrix(vsfgDetalle.Row, vsfgDetalle.Col))) >= 2 Then
        strSql = " SELECT DISTINCT producto.prd_codigo, prd_nombre " & _
                 " FROM producto " & _
                 " Where producto.emp_codigo='" & strEmpresa & "' And prd_baja=0 " & _
                 " AND prd_nombre LIKE '" & Trim(vsfgDetalle.TextMatrix(vsfgDetalle.Row, vsfgDetalle.Col)) & "%' " & _
                 " ORDER BY producto.prd_nombre "
        clsAux.Ejecutar strSql
        cmbProducto = ""
        Set cmbProducto.RowSource = clsAux.adorec_Def.DataSource
        cmbProducto.ListField = "prd_nombre"
        cmbProducto.BoundColumn = "prd_codigo"
        cmbProducto.Visible = True
        cmbProducto.SetFocus
        
    End If
End Sub


