VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerCtaxc_p 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas por Cobrar / Pagar"
   ClientHeight    =   8430
   ClientLeft      =   7185
   ClientTop       =   1080
   ClientWidth     =   9465
   Icon            =   "frmVerCtaxc_p.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Height          =   7695
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9240
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   7440
         TabIndex        =   58
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFactura 
         Height          =   285
         Left            =   3720
         TabIndex        =   57
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame FrmTipo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente / Proveedor"
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
         Height          =   1215
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtNomPersona 
            Height          =   285
            Left            =   2880
            TabIndex        =   32
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtPersona 
            Height          =   285
            Left            =   2880
            TabIndex        =   31
            Top             =   360
            Width           =   4095
         End
         Begin VB.OptionButton OptCliente 
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
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptProveedores 
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
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
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
            Left            =   1920
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
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
            Left            =   1920
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox TxtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   7320
         Width           =   1815
      End
      Begin VB.TextBox TxtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Datos de la cuenta"
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
         Height          =   3255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   9015
         Begin VB.TextBox txtBaseICE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtPorcICE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtICE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   6030
            TabIndex        =   45
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtCaduca 
            Height          =   285
            Left            =   1560
            TabIndex        =   43
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtPorcIVA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAServ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtSTIVAProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSTcero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtTipoDoc 
            Height          =   285
            Left            =   1560
            TabIndex        =   33
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtAutorizacion 
            Height          =   285
            Left            =   1560
            TabIndex        =   29
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Left            =   1560
            TabIndex        =   28
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TxtFechaEmision 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   25
            Tag             =   "5"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtdocumento 
            Height          =   285
            Left            =   2640
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtFechaPropuesta 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   3
            Tag             =   "6"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtobservacion 
            Height          =   735
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Tag             =   "7"
            Top             =   2400
            Width           =   5655
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Base ICE:"
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
            Index           =   3
            Left            =   6720
            TabIndex        =   54
            Top             =   630
            Width           =   690
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "% ICE:"
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
            Index           =   2
            Left            =   6945
            TabIndex        =   52
            Top             =   990
            Width           =   465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Total ICE:"
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
            Index           =   1
            Left            =   6750
            TabIndex        =   50
            Top             =   1350
            Width           =   660
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Total IVA:"
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
            Left            =   3795
            TabIndex        =   48
            Top             =   2070
            Width           =   705
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Sustento:"
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
            Left            =   5040
            TabIndex        =   46
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Caducidad del Doc:"
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
            Left            =   0
            TabIndex        =   44
            Top             =   1350
            Width           =   1395
         End
         Begin VB.Label Label14 
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
            Index           =   0
            Left            =   3990
            TabIndex        =   42
            Top             =   1710
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Serv:"
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
            Left            =   3120
            TabIndex        =   40
            Top             =   1350
            Width           =   1380
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal IVA Prod:"
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
            Left            =   3135
            TabIndex        =   38
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal 0%:"
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
            Left            =   6495
            TabIndex        =   36
            Top             =   1710
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Doc."
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
            TabIndex        =   34
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Autorizacion:"
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
            Left            =   0
            TabIndex        =   30
            Top             =   990
            Width           =   1470
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de emisión:"
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
            TabIndex        =   27
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observación:"
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
            TabIndex        =   26
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
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
            Left            =   120
            TabIndex        =   18
            Top             =   630
            Width           =   1350
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de pago:"
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
            TabIndex        =   17
            Top             =   2055
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   6855
            TabIndex        =   16
            Top             =   2070
            Width           =   555
         End
      End
      Begin MSDataListLib.DataCombo DcmbCodCuenta 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGAsientos 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   9000
         _cx             =   15875
         _cy             =   2778
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerCtaxc_p.frx":030A
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
      Begin MSDataListLib.DataCombo DcmbCodCuentaNF 
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   7335
         Width           =   735
      End
      Begin VB.Label LblNumCuenta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta por Cobrar No.:"
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
         TabIndex        =   22
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdImpRet 
      Caption         =   "Imprimir Retencion"
      Height          =   375
      Left            =   4777
      TabIndex        =   12
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   3097
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6472
      TabIndex        =   13
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Nueva"
      Height          =   375
      Left            =   1417
      TabIndex        =   10
      Top             =   7920
      Width           =   1575
   End
End
Attribute VB_Name = "frmVerCtaxc_p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private clsPersona As New clsConsulta
Private clsCuentas As New clsConsulta
Private clsDetCuentas As New clsConsulta
Private clsSql As New clsConsulta
Private Var_NumCuenta As Integer

Private Sub cmdCargar_Click()
    frmCargaCuentaxPagar.Show
End Sub

Private Sub cmdConsultar_Click()
    Llena_CodCuenta
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPersona = Nothing
    Set clsCuentas = Nothing
    Set clsDetCuentas = Nothing
    Set clsSql = Nothing
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la seginda columna del grid de todas las filas
    For i = 1 To (VSFGAsientos.Rows - 1)
        VSFGAsientos.TextMatrix(i, 0) = i
    Next i
End Sub

Private Sub Llena_Persona()
    On Error GoTo Error
    Dim cls_Aux As New clsConsulta
    Dim strSql As String
    cls_Aux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT cuenta_p_c.cue_p_c_codigo as cue_p_c_codigo, cue_p_c_egr_codigo,cuenta_p_c.asi_numasiento, cuenta_p_c.per_codigo,cue_p_c_pagado, " & _
             " CONCAT(persona.per_apellido,' ',persona.per_nombre) as per_nombre,persona.cat_p_tipo,cuenta_p_c.cue_p_c_fechaemision, " & _
             " cuenta_p_c.cue_p_c_fechapropuesta,cuenta_p_c.cue_p_c_descripcion,cue_p_c_valor,cue_p_c_serie,cue_p_c_numero,cue_p_c_autorizacion,cue_p_c_caduca,tip_doc_cue_descripcion, " & _
             " cue_p_c_st_prod,cue_p_c_st_serv,cue_p_c_st_cero,cue_p_c_iva,cod_iva_porcentaje,cue_p_c_baseice,cue_p_c_ice,cod_ice_porcentaje" & _
             " From cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo" & _
             " INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo" & _
             " INNER JOIN codigo_iva ON cuenta_p_c.cod_iva_codigo=codigo_iva.cod_iva_codigo" & _
             " INNER JOIN codigo_ice ON cuenta_p_c.cod_ice_codigo=codigo_ice.cod_ice_codigo" & _
             " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='" & Me.Tag & "'" & _
             " and cue_p_c_codigo='" & DcmbCodCuenta.Text & _
             "' ORDER BY cuenta_p_c.cue_p_c_codigo "
    cls_Aux.Ejecutar strSql
    If Not cls_Aux.adorec_Def.EOF > 0 Then
        cls_Aux.adorec_Def.MoveFirst
    End If
    cls_Aux.adorec_Def.Find "cue_p_c_codigo = '" & DcmbCodCuenta.Text & "'", , adSearchForward
    If clsCuentas.adorec_Def.EOF = False Then
        txtPersona.Text = cls_Aux.adorec_Def("per_codigo")
        txtNomPersona.Text = cls_Aux.adorec_Def("per_nombre")
        txtObservacion.Text = cls_Aux.adorec_Def("cue_p_c_descripcion")
        txtSerie.Text = cls_Aux.adorec_Def("cue_p_c_serie")
        txtAutorizacion.Text = cls_Aux.adorec_Def("cue_p_c_autorizacion")
        txtTipoDoc.Text = cls_Aux.adorec_Def("tip_doc_cue_descripcion")
        txtDocumento.Text = cls_Aux.adorec_Def("cue_p_c_numero")
        txtCaduca.Text = cls_Aux.adorec_Def("cue_p_c_caduca")
        txtSTIVAProd.Text = cls_Aux.adorec_Def("cue_p_c_st_prod")
        txtSTIVAServ.Text = cls_Aux.adorec_Def("cue_p_c_st_serv")
        txtPorcIVA.Text = cls_Aux.adorec_Def("cod_iva_porcentaje")
        txtIVA.Text = cls_Aux.adorec_Def("cue_p_c_iva")
        txtBaseICE.Text = cls_Aux.adorec_Def("cue_p_c_baseice")
        txtPorcICE.Text = cls_Aux.adorec_Def("cod_ice_porcentaje")
        txtICE.Text = cls_Aux.adorec_Def("cue_p_c_ice")
        txtSTcero.Text = cls_Aux.adorec_Def("cue_p_c_st_cero")
    'Llena combos de fecha
        TxtFechaEmision.Text = cls_Aux.adorec_Def("cue_p_c_fechaemision")
        TxtFechaPropuesta.Text = cls_Aux.adorec_Def("cue_p_c_fechapropuesta")
        txtValor.Text = Format(cls_Aux.adorec_Def("cue_p_c_valor"), "##0.00")
        txtDocumento.Tag = IIf(IsNull(cls_Aux.adorec_Def("asi_numasiento")), "", cls_Aux.adorec_Def("asi_numasiento"))
        If cls_Aux.adorec_Def("cue_p_c_pagado") = 1 Then
            cmdEliminar.Enabled = False
        Else
            cmdEliminar.Enabled = True
        End If
        If cls_Aux.adorec_Def("cat_p_tipo") = "C" Then
            optcliente.Value = True
            Optproveedores.Value = False
            Optproveedores.Enabled = False
        ElseIf cls_Aux.adorec_Def("cat_p_tipo") = "P" Then
            Optproveedores.Value = True
            optcliente.Value = False
            optcliente.Enabled = False
        End If
    Else
        txtPersona.Text = ""
        txtNomPersona.Text = ""
        txtObservacion.Text = ""
        txtSerie.Text = ""
        txtAutorizacion.Text = ""
        txtTipoDoc.Text = ""
        txtDocumento.Text = ""
        txtCaduca.Text = ""
        txtSTIVAProd.Text = ""
        txtSTIVAServ.Text = ""
        txtSTcero.Text = ""
        txtIVA.Text = ""
        'Llena combos de fecha
        TxtFechaEmision.Text = ""
        TxtFechaPropuesta.Text = ""
        txtValor.Text = Format(0, "##0.00")
    End If
    Exit Sub
Error:
    txtPersona.Text = ""
    txtNomPersona.Text = ""
    txtObservacion.Text = ""
    txtSerie.Text = ""
    txtAutorizacion.Text = ""
    txtTipoDoc.Text = ""
    txtDocumento.Text = ""
    txtCaduca.Text = ""
    txtSTIVAProd.Text = ""
    txtSTIVAServ.Text = ""
    txtSTcero.Text = ""
    txtIVA.Text = ""
    'Llena combos de fecha
    TxtFechaEmision.Text = ""
    TxtFechaPropuesta.Text = ""
    txtValor.Text = Format(0, "##0.00")
End Sub

Private Sub cmdAceptar_Click()
    Var_Tipo_Cuenta = Me.Tag
    frmCtaxc_p.Tag = Var_Tipo_Cuenta
    frmCtaxc_p.Show
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If vbYes = MsgBox("Está seguro(a) de eliminar la cuenta seleccionada?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar") Then
        strSql = " DELETE FROM det_asiento " & _
                 " WHERE asi_numasiento='" & txtDocumento.Tag & "' AND emp_codigo='" & strEmpresa & "'"
        clsPersona.Ejecutar (strSql), "M"
        strSql = " DELETE FROM asiento " & _
                 " WHERE asi_numasiento='" & txtDocumento.Tag & "' AND emp_codigo='" & strEmpresa & "'"
        clsPersona.Ejecutar (strSql), "M"
        strSql = " DELETE FROM cuenta_p_c " & _
                 " WHERE cue_p_c_codigo='" & DcmbCodCuenta.Text & "' AND emp_codigo='" & strEmpresa & _
                 "' AND cue_p_c_tipo='" & Me.Tag & "'"
        clsPersona.Ejecutar (strSql), "M"
        strSql = " DELETE FROM comprobante_retencion WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_codigo='" & DcmbCodCuenta.Text & "' AND cue_p_c_tipo='" & Me.Tag & "'"
        clsPersona.Ejecutar (strSql), "M"
        strSql = " DELETE FROM det_comp_ret WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_codigo='" & DcmbCodCuenta.Text & "' AND cue_p_c_tipo='" & Me.Tag & "'"
        clsPersona.Ejecutar (strSql), "M"
        DcmbCodCuenta.Text = ""
        Llena_CodCuenta
        VSFGAsientos.Clear 1
        VSFGAsientos.Rows = 2
    End If
End Sub


Private Sub cmdImpRet_Click()
    Dim frmReten As New frmReporte
    
    
'frmReten.strNumero = 111783
'frmReten.strAsiento = "2017D0016984"
'frmReten.strTipo = "P"
'frmReten.strReporte = "rptRetencionDiario"
'frmReten.Show
'frmReten.Form_Activate
'frmReten.VSPrint.PrintDoc

    frmReten.strNumero = DcmbCodCuenta.Text
    frmReten.strAsiento = VSFGAsientos.Tag
    frmReten.strTipo = Me.Tag
    frmReten.Atencion = "Fecha de Pago: " & Format(TxtFechaPropuesta.Text, "yyyy-MM-dd")
    frmReten.strReporte = "rptRetencionDiario"
    frmReten.Show
    frmReten.Form_Activate
    frmReten.VSPrint.PrintDoc
    If UCase(Left(txtTipoDoc.Text, 11)) = "LIQUIDACION" Then
        Dim frmLiqui As New frmReporte
        frmLiqui.strNumero = DcmbCodCuenta.Text
        frmLiqui.strTipo = Me.Tag
        frmLiqui.strReporte = "rptLiquidacionCompras"
        frmLiqui.Show
    End If
End Sub

Private Sub DcmbCodCuenta_Change()
    'LLena datos de cuenta escogida
    DcmbCodCuenta.Tag = "A"
    Llena_Persona
    Dim strSql As String
    
'    strSql = " SELECT tip_asi_codigo " & _
'             " FROM cuenta_p_c " & _
'             " WHERE emp_codigo = '" & strEmpresa & "' AND cue_p_c_codigo = '" & DcmbCodCuenta.Text & "' AND cue_p_c_tipo = '" & Me.Tag & "'"
'    clsSql.Ejecutar strSql
'    If Not clsSql.adorec_Def.EOF Then
'        DcmbCodCuenta.Tag = clsSql.adorec_Def("tip_asi_codigo")
'    Else
'        DcmbCodCuenta.Tag = ""
'    End If
    LLena_CombosGrid
    DcmbCodCuenta.Tag = ""
End Sub

Private Sub DcmbCodCuenta_KeyPress(KeyAscii As Integer)
    'permite poner solo numeros en el combo de cuentas
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "C" Then
        LblNumCuenta.Caption = "Cuenta por Cobrar No.:"
        Me.Caption = "Cuentas por Cobrar"
    ElseIf Me.Tag = "P" Then
        LblNumCuenta.Caption = "Cuenta por Pagar No.:"
        Me.Caption = "Cuentas por Pagar"
    End If
     
     'LLena list box de cuentas
     Llena_CodCuenta

End Sub
Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    clsPersona.Inicializar AdoConn, AdoConnMaster
    clsCuentas.Inicializar AdoConn, AdoConnMaster
    clsDetCuentas.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    
End Sub

'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
Private Sub Llena_CodCuenta()

    'LLena list box de cuenta
    If Trim(txtCodigo.Text) <> "" Or Trim(txtFactura.Text) <> "" Then
        strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.asi_numasiento,CONCAT(cue_p_c_serie,FORMAT(cue_p_c_numero,'0000000'),'-',cue_p_c_autorizacion) as cue_p_c_egr_codigo " & _
                 " From cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo" & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='" & Me.Tag & "' "
        If Trim(txtCodigo.Text) <> "" Then
            strSql = strSql & " AND cuenta_p_c.cue_p_c_codigo='" & txtCodigo.Text & "' "
        End If
        If Trim(txtFactura.Text) <> "" Then
            strSql = strSql & " AND CONCAT(cue_p_c_serie,FORMAT(cue_p_c_numero,'0000000'),'-',cue_p_c_autorizacion) like '" & txtFactura.Text & "%' "
        End If
        strSql = strSql & " ORDER BY CONCAT(cue_p_c_serie,FORMAT(cue_p_c_numero,'0000000'),'-',cue_p_c_autorizacion) "
        clsCuentas.Ejecutar (strSql)
        
        Set DcmbCodCuenta.RowSource = clsCuentas.adorec_Def.DataSource
        DcmbCodCuenta.ListField = "cue_p_c_codigo"
        DcmbCodCuenta.BoundColumn = "asi_numasiento"
        Set DcmbCodCuentaNF.RowSource = clsCuentas.adorec_Def.DataSource
        DcmbCodCuentaNF.ListField = "cue_p_c_egr_codigo"
        DcmbCodCuentaNF.BoundColumn = "cue_p_c_codigo"
    End If
End Sub
Private Sub LLena_CombosGrid()
    'Llena grid
    strSql = " SELECT det_asiento.cta_codigo, ctaconta.cta_nombre,det_asiento.det_asi_debe,det_asiento.det_asi_haber,cen_cos_nombre" & _
                 " FROM det_asiento INNER JOIN ctaconta ON det_asiento.emp_codigo = ctaconta.emp_codigo" & _
                 "                     AND det_asiento.cta_codigo = ctaconta.cta_codigo " & _
                 " LEFT JOIN centro_costo ON det_asiento.emp_codigo = centro_costo.emp_codigo" & _
                 "                     AND det_asiento.cen_cos_codigo = centro_costo.cen_cos_codigo " & _
                 " WHERE det_asiento.asi_numasiento='" & DcmbCodCuenta.BoundText & "' AND det_asiento.emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY det_asiento.cta_codigo"
    clsDetCuentas.Ejecutar (strSql)
    VSFGAsientos.Tag = DcmbCodCuenta.BoundText
    If Not clsDetCuentas.adorec_Def.EOF Then
        Set VSFGAsientos.DataSource = clsDetCuentas.adorec_Def.DataSource
        PonerBotones
        Calcula_Total
    Else
        VSFGAsientos.Clear 1
        VSFGAsientos.Rows = 2
    End If
    
End Sub

Private Sub OptCliente_Click()
    'Llena_Cliente
End Sub

Private Sub Optproveedores_Click()
    'Llena_Proveedor
End Sub


Private Sub Calcula_Total()
        'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    
    'Calcula total debe
    
    For i = 1 To VSFGAsientos.Rows - 1
        SumaDebe = SumaDebe + Val(VSFGAsientos.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    
    'Calcula total haber
    
    For i = 1 To VSFGAsientos.Rows - 1
        SumaHaber = SumaHaber + Val(VSFGAsientos.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")
End Sub
Private Sub DcmbCodCuentaNF_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If DcmbCodCuenta.Tag <> "A" Then
        If DcmbCodCuentaNF.MatchedWithList = True Then
            DcmbCodCuenta.Text = DcmbCodCuentaNF.BoundText
        End If
    End If
End Sub


Private Sub DcmbCodCuentaNF_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        DcmbCodCuenta.Text = DcmbCodCuentaNF.BoundText
    End If
End Sub

Private Sub DcmbCodCuentaNF_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    DcmbCodCuenta.Text = DcmbCodCuentaNF.BoundText
End Sub


