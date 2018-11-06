VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerAsientoVista 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Asientos"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "frmVerAsientoVista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular Asiento"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txtTipo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   1095
         Width           =   3255
      End
      Begin VB.CheckBox chkFiltroTipo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Tipo de Asiento"
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
         Left            =   3840
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroValor 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por valor"
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
         Left            =   5040
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2385
         Width           =   1695
      End
      Begin VB.CheckBox chkFiltroNumero 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por número de asiento"
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
         Left            =   3840
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroD 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por descripción"
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
         Left            =   7320
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   460
         Width           =   2295
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar / Recargar"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   42
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtD 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   7320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtNum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtValor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         MaxLength       =   24
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Height          =   1695
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   210
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
            Enabled         =   0   'False
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
            Left            =   480
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   705
            Width           =   1815
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmVerAsientoVista.frx":030A
            Left            =   1320
            List            =   "frmVerAsientoVista.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
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
            Left            =   480
            TabIndex        =   33
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            Format          =   58392579
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
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
            Left            =   1920
            TabIndex        =   34
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            Format          =   58392579
            CurrentDate     =   37463
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkFiltroCuenta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por cuenta contable"
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
         Left            =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   315
         Left            =   240
         TabIndex        =   46
         Top             =   2895
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Asiento"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3840
         TabIndex        =   53
         Top             =   855
         Width           =   3255
      End
      Begin VB.Label lblD 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7320
         TabIndex        =   50
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Asiento"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5040
         TabIndex        =   48
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Contable"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   2655
         Width           =   4560
      End
   End
   Begin VB.TextBox txtNumAsiento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DDDDDD&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4160
      Width           =   2295
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Información"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar Asiento"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo Asiento"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir Asiento"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar Asiento"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   420
      TabIndex        =   13
      Top             =   4320
      Width           =   9615
      Begin VB.TextBox TxtTotal2Haber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox TxtTotal2Debe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Tag             =   "7"
         Top             =   3360
         Width           =   9135
      End
      Begin VB.CheckBox chkMayorizado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Mayorizado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkRevisado 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Revisado"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox TxtTotal1Haber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox TxtTotal1Debe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2400
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   9360
         _cx             =   16510
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
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVerAsientoVista.frx":039E
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
      Begin VB.Label lblModificado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado por:"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblRealizado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realizado por:"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total registrado:"
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
         Left            =   3720
         TabIndex        =   19
         Top             =   2805
         Width           =   1365
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Asiento:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suma total:"
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
         Left            =   4170
         TabIndex        =   14
         Top             =   2445
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   9000
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dcmbAsiento 
      Height          =   315
      Left            =   3075
      TabIndex        =   22
      Top             =   3900
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblCorrelativo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3120
      TabIndex        =   24
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label lblAsientos 
      Alignment       =   2  'Center
      BackColor       =   &H00000050&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ningún asiento encontrado"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3075
      TabIndex        =   23
      Top             =   3660
      Width           =   3960
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   6120
      Picture         =   "frmVerAsientoVista.frx":04BF
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   6360
      Picture         =   "frmVerAsientoVista.frx":05EB
      ToolTipText     =   "Elimina una Fila"
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmVerAsientoVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private clsAsiento As New clsContable
Private Hacer As Boolean
Private HacerConsulta As Boolean
Public HacerActivate As Boolean
Public Modificando As Boolean
Private FechaAsiento As String

Private HacerFecha As Boolean
Private CuentasCargadas As Boolean
Private FechaI As Variant
Private FechaF As Variant

Private Sub chkCorrelativo_Click()
    If chkCorrelativo.value = 1 Then
        lblNum.Caption = "Número de Asiento Correlativo"
    Else
        lblNum.Caption = "Número de Asiento"
    End If
    LlenarComboAsientos
End Sub

Private Sub chkFechas_Click()
    If chkFechas.value = 1 Then
        Label1.Caption = "Fecha Inicial"
        Label2.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Label1.Caption = "Fecha"
        Label2.Enabled = False
        Fecha2.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkFiltroCuenta_Click()
    If chkFiltroCuenta.value = 1 Then
        lblCuenta.Enabled = True
        dcmbCuenta.Enabled = True
        If CuentasCargadas = False Then
            CargarCuentas
        End If
    Else
        lblCuenta.Enabled = False
        dcmbCuenta.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub CargarCuentas()
    Dim strSql As String
    Screen.MousePointer = vbHourglass
    strSql = " SELECT cta_codigo, CONCAT(cta_codigo,' - ',cta_nombre) AS cta_nombre " & _
             " FROM ctaconta WHERE emp_codigo ='" & strEmpresa & "' AND cta_subcta=0 ORDER BY cta_codigo"
    clsSql.Ejecutar strSql
    Set dcmbCuenta.RowSource = clsSql.adorec_Def.DataSource
    dcmbCuenta.ListField = "cta_nombre"
    dcmbCuenta.BoundColumn = "cta_codigo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        dcmbCuenta.BoundText = clsSql.adorec_Def(0)
        If dcmbCuenta.MatchedWithList = True Then
        End If
    End If
    CuentasCargadas = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkFiltroD_Click()
    If chkFiltroD.value = 1 Then
        lbld.Enabled = True
        txtD.Enabled = True
    Else
        lbld.Enabled = False
        txtD.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkFiltroNumero_Click()
    If chkFiltroNumero.value = 1 Then
        lblNum.Enabled = True
        txtNum.Enabled = True
    Else
        lblNum.Enabled = False
        txtNum.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkFiltroFecha_Click()
    If chkFiltroFecha.value = 1 Then
        Frame1.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        If Option1.value = True Then
            lblMes.Enabled = True
            cmbMesI.Enabled = True
        ElseIf Option2.value = True Then
            Fecha1.Enabled = True
            Label1.Enabled = True
            Fecha1.Enabled = True
            chkFechas.Enabled = True
            If chkFechas.value = 1 Then
                Label2.Enabled = True
                Fecha2.Enabled = True
            End If
        End If
    Else
        Frame1.Enabled = False
        
        Fecha2.Enabled = False
        Label1.Enabled = False
        Fecha1.Enabled = False
        Label2.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        
        Option1.Enabled = False
        Option2.Enabled = False
        lblMes.Enabled = False
        cmbMesI.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkFiltroTipo_Click()
    If chkFiltroTipo.value = 1 Then
        lblTipo.Enabled = True
        txtTipo.Enabled = True
    Else
        lblTipo.Enabled = False
        txtTipo.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkFiltroValor_Click()
    If chkFiltroValor.value = 1 Then
        lblValor.Enabled = True
        txtValor.Enabled = True
    Else
        lblValor.Enabled = False
        txtValor.Enabled = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub chkMayorizado_Click()
    If Hacer = True Then Exit Sub
    If chkMayorizado.value = 1 Then
        Hacer = True
        chkMayorizado.value = 0
        Hacer = False
    ElseIf chkMayorizado.value = 0 Then
        Hacer = True
        chkMayorizado.value = 1
        Hacer = False
    End If
End Sub

Private Sub chkRevisado_Click()
    If Hacer = True Then Exit Sub
    If chkRevisado.value = 1 Then
        Hacer = True
        chkRevisado.value = 0
        Hacer = False
    ElseIf chkRevisado.value = 0 Then
        Hacer = True
        chkRevisado.value = 1
        Hacer = False
    End If
End Sub

Private Sub CambiarFecha()
    If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1"
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal
        DiaFinal = DiaFinal - 1
    Wend
    cmdBuscar(0).Enabled = True
End Sub


Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub cmdAnular_Click()
    Dim Motivo As String
    Dim anula As Boolean
    Motivo = ""
    While Motivo = ""
        Motivo = InputBox("Motivo de Anulacion", "Contabilidad")
        If Motivo = "" Then
            If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el Asiento?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                anula = False
                Motivo = "NO ANULAR"
            End If
        Else
            anula = True
        End If
    Wend
    If anula = True Then
        clsAsiento.NumAsiento = Right(dcmbAsiento, 14)
        clsAsiento.AnularAsientoYOtros UCase(Motivo), txtDescripcion.Text
    End If
    dcmbasiento_Change
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    LlenarComboAsientos
End Sub

Private Sub cmdEliminar_Click()
    Dim ElTipo As String
    Dim NumeroDocumentos As String
    ElTipo = "asiento contable"

    If MsgBox("¿Está seguro de eliminar el " & ElTipo & " número " & Right(dcmbAsiento, 14) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Pregunta - Eliminar") = vbNo Then Exit Sub
    clsAsiento.NumAsiento = Right(dcmbAsiento, 14)
    clsAsiento.EliminarAsiento True, True
    MsgBox "Asiento " & dcmbAsiento & " eliminado.", vbInformation, "Eliminar"
    'Set clsAsiento = Nothing
    LlenarComboAsientos
End Sub

Private Sub cmdImprimir_Click()
    frmReporte.strAsiento = dcmbAsiento
    frmReporte.strReporte = "rptAsiento"
    frmReporte.Show
End Sub

Private Sub cmdInfo_Click()
    MsgBox clsAsiento.VerMasDatos, vbInformation, "Contabilidad"
End Sub

Private Sub cmdModificar_Click()
    frmAsiento.Tag = "M"
    frmAsiento.txtAsiento = Right(dcmbAsiento, 14)
    frmAsiento.Fecha1 = FechaAsiento
    frmAsiento.chkRevisado.value = Me.chkRevisado.value
    frmAsiento.txtDescripcion = Me.txtDescripcion
    frmAsiento.TxtTotal1Debe = Me.TxtTotal1Debe
    frmAsiento.TxtTotal1Haber = Me.TxtTotal1Haber
    frmAsiento.VSFG.Rows = Me.VSFG.Rows
    For i = 1 To VSFG.Rows - 1
        frmAsiento.VSFG.TextMatrix(i, 1) = Me.VSFG.TextMatrix(i, 1)
        frmAsiento.VSFG.TextMatrix(i, 2) = Me.VSFG.TextMatrix(i, 2)
        frmAsiento.VSFG.TextMatrix(i, 3) = Me.VSFG.TextMatrix(i, 3)
        frmAsiento.VSFG.TextMatrix(i, 4) = Me.VSFG.TextMatrix(i, 4)
        frmAsiento.VSFG.TextMatrix(i, 5) = Me.VSFG.TextMatrix(i, 5)
        frmAsiento.VSFG.TextMatrix(i, frmAsiento.VSFG.Cols - 1) = Me.VSFG.TextMatrix(i, Me.VSFG.Cols - 1)
        frmAsiento.VSFG.Cell(flexcpBackColor, i, 1, i, frmAsiento.VSFG.Cols - 1) = Me.VSFG.Cell(flexcpBackColor, i, 1, i, 1)
    Next i
    frmAsiento.Show
End Sub

Private Sub cmdNuevo_Click()
    frmAsiento.Tag = "N"
    frmAsiento.Show
    frmAsiento.Manual = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbasiento_Change()
    If HacerConsulta = True Then
        clsAsiento.NumAsiento = Right(dcmbAsiento, 14)
        LlenarDetalleAsiento
        If clsAsiento.Eliminable = False Then
            cmdEliminar.Enabled = False
        Else
            cmdEliminar.Enabled = True
        End If
        
    End If
End Sub

Private Sub dcmbCuenta_Change()
    cmdBuscar(0).Enabled = True
End Sub

Private Sub Fecha1_Change()
    If chkFechas.value = 0 Then
        Fecha2 = Fecha1
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub Fecha2_Change()
    cmdBuscar(0).Enabled = True
End Sub

Private Sub Form_Activate()
    If HacerActivate = True Then
        If Modificando = True Then
            LlenarDetalleAsiento
            Modificando = False
        Else
            If cmdBuscar(0).Enabled = False Then
                LlenarComboAsientos
            End If
        End If
        HacerActivate = False
    End If
    cmdBuscar(0).Enabled = True
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    HacerActivate = False
    
    CuentasCargadas = False
    HacerFecha = True
    'Selecciona el mes actual
    Fecha1 = HoyDia
    Fecha2 = HoyDia
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsAsiento.Inicializar AdoConn, AdoConnMaster
End Sub

Private Sub LlenarComboAsientos()
    Dim strSql As String
    Screen.MousePointer = vbHourglass
    HacerConsulta = False
    dcmbAsiento = ""
    HacerConsulta = True
    strSql = " SELECT CONCAT(asiento.asi_numasiento) AS correlativo "
    strSql = strSql & " FROM asiento "
    If chkFiltroValor.value = 1 Then
        txtValor = FormatoD2(txtValor)
    End If
    
    If chkFiltroValor.value = 1 Or chkFiltroCuenta.value = 1 Then
        strSql = strSql & " INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento "
    End If
    
    strSql = strSql & " WHERE asiento.emp_codigo='" & strEmpresa & "' "
    
    If chkFiltroFecha.value = 1 Then
        If Option1.value = True Then
            strSql = strSql & " AND asi_fecha BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
        ElseIf Option2.value = True Then
            strSql = strSql & " AND asi_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
        End If
    End If
    If chkFiltroD.value = 1 Then
        strSql = strSql & " AND asi_descripcion LIKE '%" & txtD & "%'"
    End If
    If chkFiltroTipo.value = 1 Then
        txtTipo = UCase(Left(txtTipo, 1))
        strSql = strSql & " AND asiento.asi_numasiento LIKE '%" & txtTipo & "%'"
    End If
    If chkFiltroNumero.value = 1 Then
        txtNum = Trim(txtNum)
        strSql = strSql & " AND asiento.asi_numasiento LIKE '%" & txtNum & "'"
    End If
    If chkFiltroValor.value = 1 Then
        strSql = strSql & " AND (asi_totaldebe = " & txtValor & " OR asi_totalhaber = " & txtValor & _
                    " OR det_asi_debe = " & txtValor & " OR det_asi_haber = " & txtValor & " )   "
    End If
    
    If chkFiltroCuenta.value = 1 Then
        strSql = strSql & " AND cta_codigo = '" & dcmbCuenta.BoundText & "'"
    End If
    
    If chkFiltroValor.value = 1 Or chkFiltroCuenta.value = 1 Then
        strSql = strSql & " GROUP BY asiento.asi_numasiento"
    End If
    
    strSql = strSql & " ORDER BY asiento.asi_numasiento DESC"
    clsSql.Ejecutar strSql
    Set dcmbAsiento.RowSource = clsSql.adorec_Def.DataSource
    dcmbAsiento.ListField = "correlativo"
    dcmbAsiento.BoundColumn = "correlativo"
    If clsSql.adorec_Def.RecordCount > 0 Then
        HacerConsulta = False
        dcmbAsiento.BoundText = clsSql.adorec_Def("correlativo")
        If dcmbAsiento.MatchedWithList = True Then
        End If
        HacerConsulta = True
        If clsSql.adorec_Def.RecordCount = 1 Then
            lblAsientos.Caption = "1 asiento encontrado"
        Else
            lblAsientos.Caption = clsSql.adorec_Def.RecordCount & " asientos encontrados"
        End If
        dcmbAsiento.SetFocus
        dcmbasiento_Change
        
    Else
        lblAsientos.Caption = "Ningún asiento encontrado"
    End If
    
    'LlenarDetalleAsiento
    'cmdBuscar(0).Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub LlenarDetalleAsiento()
    Dim ElDebe As Double
    Dim ElHaber As Double
    Dim UsuMod As String
    Dim FechaMod As String
    Dim strSql As String
    If dcmbAsiento = "" Then
        Limpiar
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    strSql = " SELECT asi_fecha,COALESCE(asi_descripcion,'') AS asi_descripcion,asi_revisado,asi_mayorizado, asi_totaldebe, asi_totalhaber, asi_fechamod, asi_usumod " & _
             " FROM asiento " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento = '" & Right(dcmbAsiento, 14) & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        HacerConsulta = False
        
        HacerConsulta = True
        txtNumAsiento.Text = dcmbAsiento.Text
        lblfecha.Caption = "Fecha Asiento: " & clsSql.adorec_Def("asi_fecha")
        FechaAsiento = clsSql.adorec_Def("asi_fecha")
        lblRealizado.Caption = "Realizado por: " & clsSql.adorec_Def("asi_usumod") & " el " & Left(clsSql.adorec_Def("asi_fechamod"), 10) & " a las " & Mid(clsSql.adorec_Def("asi_fechamod"), 12, 8)
        FechaMod = clsSql.adorec_Def("asi_fechamod")
        UsuMod = clsSql.adorec_Def("asi_usumod")
        txtDescripcion.Text = clsSql.adorec_Def("asi_descripcion")
        Hacer = True
        chkRevisado.value = clsSql.adorec_Def("asi_revisado")
        chkMayorizado.value = clsSql.adorec_Def("asi_mayorizado")
        
        If chkMayorizado.value = 1 Then
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
        Else
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        End If
        cmdInfo.Enabled = True
        Hacer = False
        TxtTotal2Debe = Format(clsSql.adorec_Def("asi_totaldebe"), "####0.00")
        TxtTotal2Haber = Format(clsSql.adorec_Def("asi_totalhaber"), "####0.00")
        cmdImprimir.Enabled = True
        strSql = " SELECT det_asiento.cta_codigo, cta_nombre, det_asi_debe, det_asi_haber,COALESCE(cen_cos_nombre,'') as cen_cos_nombre, det_asi_fechamod, det_asi_usumod,'0' as modi " & _
                 " FROM det_asiento INNER JOIN ctaconta ON ctaconta.cta_codigo=det_asiento.cta_codigo " & _
                 " AND ctaconta.emp_codigo=det_asiento.emp_codigo" & _
                 " LEFT JOIN centro_costo ON det_asiento.emp_codigo=centro_costo.emp_codigo" & _
                 " AND det_asiento.cen_cos_codigo=centro_costo.cen_cos_codigo" & _
                 " WHERE det_asiento.emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento = '" & Right(dcmbAsiento, 14) & "' " & _
                 " ORDER BY det_asiento.cta_codigo"
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        If clsSql.adorec_Def.EOF = False Then
            lblModificado.Caption = "Modificado por: " & clsSql.adorec_Def("det_asi_usumod") & " el " & Left(clsSql.adorec_Def("det_asi_fechamod"), 10) & " a las " & Mid(clsSql.adorec_Def("det_asi_fechamod"), 12, 8)
        Else
            lblModificado.Caption = "Modificado por:"
        End If
        ElDebe = 0
        ElHaber = 0
        For i = 1 To VSFG.Rows - 1
            If clsAsiento.Modificable(VSFG.TextMatrix(i, 1)) = True Then
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 1
                VSFG.Cell(flexcpBackColor, i, 1, i, VSFG.Cols - 1) = &HFFFFFF
            Else
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 0
                VSFG.Cell(flexcpBackColor, i, 1, i, VSFG.Cols - 1) = &HC0FFFF
            End If
            ElDebe = ElDebe + VSFG.TextMatrix(i, 3)
            ElHaber = ElHaber + VSFG.TextMatrix(i, 4)
        Next i
        TxtTotal1Debe = Format(ElDebe, "####0.00")
        TxtTotal1Haber = Format(ElHaber, "####0.00")
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Limpiar()
    VSFG.Clear 1
    VSFG.Rows = 1
    TxtTotal1Debe = "0.00"
    TxtTotal1Haber = "0.00"
    TxtTotal2Debe = "0.00"
    TxtTotal2Haber = "0.00"
    txtNumAsiento.Text = ""
    Hacer = True
    chkRevisado.value = 0
    chkMayorizado.value = 0
    Hacer = False
    txtDescripcion.Text = ""
    cmdModificar.Enabled = False
    cmdImprimir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdInfo.Enabled = False
    lblfecha.Caption = "Fecha Asiento:"
    lblRealizado.Caption = "Realizado por:"
    lblModificado.Caption = "Modificado por:"
    FechaAsiento = ""
    Frame2.Caption = "Asiento"
End Sub

Private Sub Option1_Click()
    If Option1.value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Label1.Enabled = False
        Fecha1.Enabled = False
        Label2.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        cmdBuscar(0).Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        Label1.Enabled = True
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.value = 1 Then
            Label2.Enabled = True
            Fecha2.Enabled = True
        End If
        cmdBuscar(0).Enabled = True
    End If
End Sub

Private Sub txtD_Change()
    'cmdBusca(0).Enabled = True
End Sub

Private Sub txtNum_Change()
    cmdBuscar(0).Enabled = True
End Sub

Private Sub txtValor_Change()
    cmdBuscar(0).Enabled = True
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub
