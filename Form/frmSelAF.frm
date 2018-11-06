VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelAF 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activos Fijos"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelAF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdVender 
      Caption         =   "&Vender"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdBaja 
      Caption         =   "&Dar De Baja"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Activos Fijos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5880
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   10215
      Begin VSFlex8Ctl.VSFlexGrid VSFDep 
         Height          =   1455
         Left            =   120
         TabIndex        =   52
         Top             =   3600
         Width           =   4335
         _cx             =   7646
         _cy             =   2566
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelAF.frx":030A
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
      Begin VB.TextBox txtValor_dep2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtValorT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   32
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtValorR 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtFecha 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8880
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtAsiento 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8880
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkVendido 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Vendido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4560
         TabIndex        =   25
         Top             =   3600
         Width           =   5535
         Begin VB.TextBox txtProveedor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   3465
         End
         Begin VB.TextBox txtFactura 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Factura"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   3720
            TabIndex        =   48
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proveedor"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4560
         TabIndex        =   21
         Top             =   4680
         Width           =   5535
         Begin VB.TextBox txtFechaB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   1665
         End
         Begin VB.TextBox txtFechaD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   1665
         End
         Begin VB.TextBox txtFechaA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Baja"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   3720
            TabIndex        =   51
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Depreciación"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   50
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Adquisición"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox txtValor_dep 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtUbicacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtCustodio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CheckBox chkBaja 
         BackColor       =   &H00DDDDDD&
         Caption         =   "De baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtMarca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtVidautil 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1800
         _ExtentX        =   3175
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
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   315
         Left            =   2025
         TabIndex        =   1
         Top             =   600
         Width           =   3870
         _ExtentX        =   6826
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
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo por Depreciar"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7080
         TabIndex        =   46
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vida Útil"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ubicación"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   44
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Custodio"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Depreciación Reval. (-)"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7920
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Depreciación (-)"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Revalorización (+)"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7920
         TabIndex        =   40
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo inicial (+)"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   37
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   36
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDDDDD&
         Caption         =   "Total:"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   5070
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDDDDD&
         Caption         =   "Años"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   3120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSelAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion del Activo_Fijo y poder modificar,                  #
'#  crear o eliminar Activo_Fijos                                               #
'#  frmSelAFV1.0                                                                #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los Activo_Fijos que al momento estan                #
'#  ingresados en el sistema. Desde esta ventana se puede crear un nuevo        #
'#  Activo_Fijo, modificar o eliminar los Activo_Fijos ya creados.              #
'#  Desde esta ventana se llama a la ventana frmActivoFijo en la que crea       #
'#  y modifica los Activo_Fijos                                                 #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    tipo_ativo_Fijo: En esta tabla se almacenan y se sacan los tipos de       #
'#               activos Fijos que se pueden asignar a los activos fijos        #
'#               con su respectivo codigo.                                      #
'#    marca_activo_Fijo: En esta tabla se almacenan y se sacan las marcas de    #
'#               de los activos fijos con sus nombres y codigos.                #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private clsDet As New clsConsulta
Private clsDelete As New clsConsulta
Private clsDep_mod As New clsConsulta
Private clsDep As New clsConsulta
Private clsSql As New clsConsulta
Private Hacer As Boolean
Private Cuenta1 As String
Private Cuenta2 As String
Private Cuenta3 As String
Private Cuenta4 As String
Private DeBaja As Boolean
Private HacerActivate As Boolean


Private Sub chkBaja_Click()
    If Hacer = True Then Exit Sub
    If chkBaja.value = 1 Then
        Hacer = True
        chkBaja.value = 0
        Hacer = False
    ElseIf chkBaja.value = 0 Then
        Hacer = True
        chkBaja.value = 1
        Hacer = False
    End If
End Sub

Private Sub chkVendido_Click()
    If Hacer = True Then Exit Sub
    If chkVendido.value = 1 Then
        Hacer = True
        chkVendido.value = 0
        Hacer = False
    ElseIf chkVendido.value = 0 Then
        Hacer = True
        chkVendido.value = 1
        Hacer = False
    End If
End Sub

Private Sub cmdBaja_Click()
    
    'If MsgBox("¿Está seguro de dar de baja el activo fijo " & Me.dcmbnombre & "?" & Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Dar De Baja") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    DeBaja = True
    Set frmAsiento.Objeto = txtAsiento
    'Set frmAsiento.Objeto1 = txtFecha
    frmAsiento.Tag = "N"
    frmAsiento.ActivoFijo = True
    frmAsiento.VSFG.Rows = 5
    frmAsiento.VSFG.TextMatrix(1, 1) = Cuenta1
    frmAsiento.VSFG.TextMatrix(1, 3) = 0
    frmAsiento.VSFG.TextMatrix(1, 4) = Me.txtValor
    frmAsiento.VSFG.TextMatrix(2, 1) = Cuenta2
    frmAsiento.VSFG.TextMatrix(2, 3) = Me.txtValor_dep
    frmAsiento.VSFG.TextMatrix(2, 4) = 0
    frmAsiento.VSFG.TextMatrix(3, 1) = Cuenta3
    frmAsiento.VSFG.TextMatrix(3, 3) = 0
    frmAsiento.VSFG.TextMatrix(3, 4) = Me.txtValorR
    frmAsiento.VSFG.TextMatrix(4, 1) = Cuenta4
    frmAsiento.VSFG.TextMatrix(4, 3) = Me.txtValor_dep2
    frmAsiento.VSFG.TextMatrix(4, 4) = 0
    frmAsiento.TxtTotal1Haber = Me.txtValor
    frmAsiento.TxtTotal1Debe = Me.txtValor_dep
    frmAsiento.txtDescripcion = "DADA DE BAJA DEL ACTIVO FIJO: " & Me.dcmbNombre & " (CÓDIGO: " & Me.dcmbCodigo & ")"
    strSql = " SELECT MAX(asi_fecha) FROM asiento INNER JOIN det_activo_fijo" & _
             " ON asiento.asi_numasiento=det_activo_fijo.asi_numasiento AND asiento.emp_codigo=det_activo_fijo.emp_codigo " & _
             " WHERE det_activo_fijo.emp_codigo='" & strEmpresa & "' AND act_fij_codigo='" & Me.dcmbCodigo & "'" & _
             " GROUP BY asiento.emp_codigo"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        frmAsiento.FechaMinima = clsSql.adorec_Def(0)
    End If
    frmAsiento.Show
    Screen.MousePointer = vbDefault
    strSql = " UPDATE activo_fijo " & _
             " SET act_fij_baja = '1' " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND act_fij_codigo='" & Me.dcmbCodigo & "'"
    clsSql.Ejecutar strSql, "M"
    BuscarActivos
End Sub

Private Sub cmdEliminar_Click()
    'If VerificarFechaContable(Me.txtFechaA) = False Then Exit Sub
    Dim strSql As String
    
    ' Consulta para conocer si hay activo fijos en Detalle adquisicion
    strSql = " SELECT count(det_act_fij_codigo) As Num " & _
             " FROM det_activo_fijo " & _
             " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsDet.Ejecutar (strSql)
    ' Si existen activos fijos no se eliminan
    If clsDet.adorec_Def("Num") > 0 Then
        Dim Mensaje As String
        If clsDet.adorec_Def("Num") = 1 Then
            Mensaje = "1 depreciación relacionada"
        Else
            Mensaje = clsDet.adorec_Def("Num") & " depreciaciones relacionadas"
        End If
        MsgBox "No puede eliminar este activo fijo. Hay " & Mensaje & ".", vbInformation, "Eliminación"
        Exit Sub
    ' Si no existen activos fijos se elimina
'        Else
'        ' Consulta para conocer si hay activos fijos en depreciacion de Activos Fijos
'        strSql = " SELECT count(*) As Ing " & _
'                 " FROM depreciacion_activo " & _
'                 " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
'                 " AND emp_codigo='" & strEmpresa & "'"
'        clsCon_Def.Ejecutar (strSql)
'        ' Si existen Activos Fijos no se elimina
'        If clsCon_Def.adorec_Def("Ing") > 0 Then
'            MsgBox "No Puede eliminar este Activo Fijo, esta depreciado en alguna area", vbInformation, "Eliminación"
'                ' Si no existen Activo Fijo se elimina
    Else
        Mensaje = "¿Está seguro de eliminar el Activo Fijo " & Me.dcmbNombre & "?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "Activos Fijos"   ' Define el título.
        respuesta = MsgBox(Mensaje, Estilo, Título)
        If respuesta = vbYes Then
            
            clsDelete.Inicializar AdoConn, AdoConnMaster
            strSql = " DELETE " & _
                           " FROM depreciacion_activo " & _
                           " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
                           " AND emp_codigo='" & strEmpresa & "'"
            clsDelete.Ejecutar strSql, "M"
            'Se elimina  Activos Fijos
            strSql = " DELETE " & _
                           " FROM activo_fijo " & _
                           " WHERE act_fij_codigo = '" & dcmbCodigo.Text & "' " & _
                           " AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar strSql, "M"
            'MsgBox "Activo Fijo eliminado", vbInformation, "Eliminación"
            'Consulta los Activos Fijos que estan disponibles
            limpiar
            BuscarActivos
        End If
    End If
End Sub

Private Sub limpiar()
    dcmbCodigo.Text = ""

End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de un Activo_fijo, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del activo_fijo que se modificará
    Screen.MousePointer = vbHourglass
    Dim i As Integer
    Dim intPos As Integer
    Dim strCodAux As String
    
    frmActivoFijo.Show
    frmActivoFijo.Tag = "M"
    frmActivoFijo.codigoAF = Me.dcmbCodigo.Text
    frmActivoFijo.txtCodigo.Text = Me.dcmbCodigo.Text
    frmActivoFijo.txtNombre.Text = Me.dcmbNombre.Text
    frmActivoFijo.txtDescripcion.Text = Me.txtDescripcion.Text
    frmActivoFijo.dcmbTipo.Text = Me.txtTipo.Text
    frmActivoFijo.dcmbMarca.Text = Me.txtMarca.Text
    frmActivoFijo.chkBaja.value = Me.chkBaja.value
    frmActivoFijo.chkVendido.value = Me.chkVendido.value
    frmActivoFijo.txtValor.Text = Me.txtValor.Text
    frmActivoFijo.txtValor_dep.Text = Me.txtValor_dep.Text
    frmActivoFijo.txtValorR.Text = Me.txtValorR.Text
    If Trim(txtFechaA) <> "" Then frmActivoFijo.FechaA = txtFechaA
    If Trim(txtFechaB) <> "" Then frmActivoFijo.FechaB = txtFechaB
    If Trim(txtFechaD) <> "" Then frmActivoFijo.FechaD = txtFechaD
    
    frmActivoFijo.dcmbTipo.Text = Me.txtTipo.Text
    frmActivoFijo.dcmbMarca.Text = Me.txtMarca.Text
    frmActivoFijo.txtCustodio.Text = Me.txtCustodio.Text
    frmActivoFijo.txtUbicacion.Text = Me.txtUbicacion.Text
    frmActivoFijo.txtVidaUtil.Text = Me.txtVidaUtil.Text
    If Trim(Me.txtProveedor.Tag) <> "" Then frmActivoFijo.dcmbProveedor.BoundText = Me.txtProveedor.Tag
    frmActivoFijo.dcmbFactura.BoundText = Me.txtFactura.Tag
    
    HacerActivate = True
    Screen.MousePointer = vbdafault
End Sub

Private Sub cmdNuevo_Click()
' Crea un nuevo producto, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nuevo producto
    Screen.MousePointer = vbHourglass
    frmActivoFijo.Show
    frmActivoFijo.Tag = "N"
    HacerActivate = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVender_Click()
    'If MsgBox("¿Está seguro de vender el activo fijo " & Me.dcmbnombre & "?" & Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Vender Activo") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    DeBaja = False
    Set frmAsiento.Objeto = txtAsiento
    Set frmAsiento.Objeto1 = txtFecha
    frmAsiento.Tag = "N"
    frmAsiento.ActivoFijo = True
    frmAsiento.VSFG.Rows = 5
    frmAsiento.VSFG.TextMatrix(1, 1) = Cuenta1
    frmAsiento.VSFG.TextMatrix(1, 3) = 0
    frmAsiento.VSFG.TextMatrix(1, 4) = Me.txtValor
    frmAsiento.VSFG.TextMatrix(2, 1) = Cuenta2
    frmAsiento.VSFG.TextMatrix(2, 3) = Me.txtValor_dep
    frmAsiento.VSFG.TextMatrix(2, 4) = 0
    frmAsiento.VSFG.TextMatrix(3, 1) = Cuenta3
    frmAsiento.VSFG.TextMatrix(3, 3) = 0
    frmAsiento.VSFG.TextMatrix(3, 4) = Me.txtValorR
    frmAsiento.VSFG.TextMatrix(4, 1) = Cuenta4
    frmAsiento.VSFG.TextMatrix(4, 3) = Me.txtValor_dep2
    frmAsiento.VSFG.TextMatrix(4, 4) = 0
    frmAsiento.TxtTotal1Haber = Me.txtValor
    frmAsiento.TxtTotal1Debe = Me.txtValor_dep
    frmAsiento.txtDescripcion = "VENTA DEL ACTIVO FIJO: " & Me.dcmbNombre & " (CÓDIGO: " & Me.dcmbCodigo & ")"
    
    strSql = " SELECT MAX(asi_fecha) FROM asiento INNER JOIN det_activo_fijo" & _
             " ON asiento.asi_numasiento=det_activo_fijo.asi_numasiento AND asiento.emp_codigo=det_activo_fijo.emp_codigo " & _
             " WHERE det_activo_fijo.emp_codigo='" & strEmpresa & "' AND act_fij_codigo='" & Me.dcmbCodigo & "'" & _
             " GROUP BY asiento.emp_codigo"
    clsSql.Ejecutar (strSql)
    If clsSql.adorec_Def.RecordCount > 0 Then
        frmAsiento.FechaMinima = clsSql.adorec_Def(0)
    End If
    frmAsiento.Show
    
    Screen.MousePointer = vbDefault
    strSql = " UPDATE activo_fijo " & _
             " SET act_fij_vendido = '1' " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND act_fij_codigo='" & Me.dcmbCodigo & "'"
    clsSql.Ejecutar strSql, "M"
    BuscarActivos
End Sub

Private Sub dcmbCodigo_Change()
'Chequea el activo fijo seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
        If dcmbCodigo.Text = "" Then
            Call borrar_datos
            Call limpiarFxGD
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        clsCon_Def.Actualizar
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = " act_fij_codigo = '" & dcmbCodigo.Text & "' "
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("act_fij_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            txtDescripcion.Text = clsCon_Def.adorec_Def("act_fij_descripcion")
            
            txtValor.Text = clsCon_Def.adorec_Def("act_fij_valor")
            txtValor_dep.Text = clsCon_Def.adorec_Def("act_fij_depreciado")
            txtValorR.Text = clsCon_Def.adorec_Def("act_fij_revalorizado")
            txtValor_dep2.Text = clsCon_Def.adorec_Def("act_fij_depreciado2")
            txtValorT = FormatoD2(txtValor) - FormatoD2(txtValor_dep) + FormatoD0(txtValorR) - FormatoD2(txtValor_dep2)
            
            Hacer = True
            chkBaja.value = clsCon_Def.adorec_Def("act_fij_baja")
            chkVendido.value = clsCon_Def.adorec_Def("act_fij_vendido")
            Hacer = False
            If chkBaja.value = 1 Or chkVendido.value = 1 Then
                Me.cmdBaja.Enabled = False
                Me.cmdVender.Enabled = False
            Else
                Me.cmdBaja.Enabled = True
                Me.cmdVender.Enabled = True
            End If
            txtFechaA.Text = clsCon_Def.adorec_Def("act_fij_fecha_adq")
            If IsNull(clsCon_Def.adorec_Def("act_fij_fecha_baja")) = False Then
                txtFechaB.Text = clsCon_Def.adorec_Def("act_fij_fecha_baja")
            Else
                txtFechaB.Text = ""
            End If
            If IsNull(clsCon_Def.adorec_Def("act_fij_fecha_dep")) = False Then
                txtFechaD.Text = clsCon_Def.adorec_Def("act_fij_fecha_dep")
            Else
                txtFechaD.Text = ""
            End If
            txtTipo.Text = clsCon_Def.adorec_Def("tip_act_nombre")
            txtMarca.Text = clsCon_Def.adorec_Def("mar_act_fij_nombre")
            txtCustodio.Text = clsCon_Def.adorec_Def("act_fij_custodio")
            txtUbicacion.Text = clsCon_Def.adorec_Def("act_fij_ubicacion")
            txtVidaUtil.Text = clsCon_Def.adorec_Def("act_fij_vida_util")
            
            txtProveedor.Text = clsCon_Def.adorec_Def("proveedor")
            txtProveedor.Tag = clsCon_Def.adorec_Def("CodProveedor")
            txtFactura.Text = clsCon_Def.adorec_Def("factura")
            txtFactura.Tag = clsCon_Def.adorec_Def("cue_p_c_codigo")
            Cuenta1 = clsCon_Def.adorec_Def("tip_act_ctaconta")
            Cuenta2 = clsCon_Def.adorec_Def("tip_act_ctaconta2")
            Cuenta3 = clsCon_Def.adorec_Def("tip_act_ctaconta3")
            Cuenta4 = clsCon_Def.adorec_Def("tip_act_ctaconta4")
            cmdNuevo.Enabled = True
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            Call borrar_datos
            Call limpiarFxGD
        End If
        dcmbCodigo.Tag = ""
        
        'llenar flexgrid
        strSql = " SELECT area.are_codigo,area.are_nombre,COALESCE(cen_cos_nombre,'') as cen_cos_nombre,depreciacion_activo.dep_act_porcentaje " & _
                 " FROM area INNER JOIN depreciacion_activo ON area.emp_codigo =depreciacion_activo.emp_codigo AND area.are_codigo =depreciacion_activo.are_codigo " & _
                 " LEFT JOIN centro_costo ON depreciacion_activo.emp_codigo =centro_costo.emp_codigo AND depreciacion_activo.cen_cos_codigo =centro_costo.cen_cos_codigo " & _
                 " WHERE area.emp_codigo = '" & strEmpresa & "' AND act_fij_codigo =  '" & dcmbCodigo.Text & "' " & _
                 " ORDER BY area.are_nombre"
        clsDep_mod.Ejecutar (strSql)
        
        If (clsDep_mod.adorec_Def.RecordCount > 0) Then
            TxtTotal.Text = 0
            Set VSFDep.DataSource = clsDep_mod.adorec_Def.DataSource
            VSFDep.Enabled = True
            Call CalcuTotal
        Else
            VSFDep.Enabled = False
           Call limpiarFxGD
           TxtTotal.Text = " "
        End If
    Screen.MousePointer = vbDefault
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub dcmbNombre_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbNombre.Text = "" Then
        Call borrar_datos
    End If
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub

Private Sub Form_Activate()
'Centra esta forma dentro de la forma MDI
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        clsDet.Inicializar AdoConn, AdoConnMaster
        clsDep.Inicializar AdoConn, AdoConnMaster
        clsDep_mod.Inicializar AdoConn, AdoConnMaster
        clsSql.Inicializar AdoConn, AdoConnMaster
    If HacerActivate = True Then
        BuscarActivos
        HacerActivate = False
        dcmbCodigo_Change
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    On Error GoTo errhandler
        
       
        Hacer = False
        HacerActivate = True
        
    Exit Sub
       
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub BuscarActivos()
    strSql = " SELECT activo_fijo.act_fij_codigo, activo_fijo.act_fij_nombre, marca_activo_fijo.mar_act_fij_nombre, tipo_activo.tip_act_nombre, " & _
            " activo_fijo.act_fij_descripcion, activo_fijo.act_fij_vida_util, activo_fijo.act_fij_fecha_adq, activo_fijo.act_fij_fecha_baja, " & _
            " activo_fijo.act_fij_fecha_dep, activo_fijo.act_fij_fecha_baja, activo_fijo.act_fij_valor, activo_fijo.act_fij_revalorizado, " & _
            " activo_fijo.act_fij_custodio , activo_fijo.act_fij_depreciado, activo_fijo.act_fij_depreciado2, activo_fijo.act_fij_ubicacion, activo_fijo.act_fij_baja, IFNULL(concat(cue_p_c_serie,' ',cue_p_c_egr_codigo),'') AS factura, IFNULL(concat(per_apellido,' ',per_nombre),'') AS proveedor, IFNULL(cuenta_p_c.cue_p_c_codigo,'') AS cue_p_c_codigo, IFNULL(cuenta_p_c.per_codigo,'') as CodProveedor, act_fij_vendido, tip_act_ctaconta, tip_act_ctaconta2, tip_act_ctaconta3, tip_act_ctaconta4" & _
            " FROM activo_fijo INNER JOIN " & _
            " tipo_activo ON activo_fijo.tip_act_codigo = tipo_activo.tip_act_codigo AND " & _
            " activo_fijo.emp_codigo = tipo_activo.emp_codigo INNER JOIN " & _
            " marca_activo_fijo ON activo_fijo.mar_act_fij_codigo = marca_activo_fijo.mar_act_fij_codigo " & _
            " LEFT JOIN cuenta_p_c ON cuenta_p_c.cue_p_c_codigo=activo_fijo.cue_p_c_codigo AND cuenta_p_c.emp_codigo=activo_fijo.emp_codigo AND cuenta_p_c.cue_p_c_tipo='P'" & _
            " LEFT JOIN persona ON cuenta_p_c.per_codigo=persona.per_codigo AND cuenta_p_c.emp_codigo=persona.emp_codigo AND persona.cat_p_tipo='P'" & _
            " WHERE (activo_fijo.emp_codigo = '" & strEmpresa & "') " & _
            " ORDER BY activo_fijo.act_fij_nombre "
        clsCon_Def.Ejecutar (strSql)
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False

        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If clsCon_Def.adorec_Def.RecordCount = 1 Then
                Label6.Caption = "Nombre - 1 registro"
            Else
                Label6.Caption = "Nombre - " & clsCon_Def.adorec_Def.RecordCount & " registros"
            End If
            Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbCodigo.ListField = "act_fij_codigo"
            Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
                dcmbNombre.ListField = "act_fij_nombre"
                dcmbNombre.BoundColumn = "act_fij_codigo"
                dcmbNombre.Enabled = True
                dcmbCodigo.Enabled = True
                dcmbCodigo.BoundText = clsCon_Def.adorec_Def("act_fij_codigo")
        Else
            Set dcmbCodigo.RowSource = Nothing
            Set dcmbNombre.RowSource = Nothing
                dcmbNombre.Enabled = False
                dcmbCodigo.Enabled = False
        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
Public Sub borrar_datos()
        
            dcmbNombre.Text = ""
            txtDescripcion.Text = ""
            txtValor.Text = ""
            txtValor_dep.Text = ""
            txtValorR.Text = ""
            txtValorT.Text = ""
            chkBaja.value = 0
            txtFechaA.Text = ""
            txtFechaB.Text = ""
            txtFechaD.Text = ""
            txtTipo.Text = ""
            txtMarca.Text = ""
            txtCustodio.Text = ""
            txtUbicacion.Text = ""
            txtVidaUtil.Text = ""
            txtProveedor.Text = ""
            txtFactura.Text = ""
        
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
        cmdBaja.Enabled = False
        cmdVender.Enabled = False
End Sub


Private Sub txtAsiento_Change()
    'Luego de dar de bajar el asiento
    If Trim(txtAsiento) <> "" Then
        'MsgBox txtAsiento
        If DeBaja = True Then
            NuevoDetalleActivo Me.dcmbCodigo, "BAJA", txtAsiento, 0, 0, 0
            strSql = " UPDATE activo_fijo " & _
                " SET act_fij_baja=1, act_fij_fecha_baja='" & txtFecha & "'" & _
                " WHERE act_fij_codigo='" & Me.dcmbCodigo & "' " & _
                " AND emp_codigo='" & strEmpresa & "'"
        Else
            NuevoDetalleActivo Me.dcmbCodigo, "VENTA", txtAsiento, 0, 0, 0
            strSql = " UPDATE activo_fijo " & _
                " SET act_fij_vendido=1, act_fij_fecha_baja='" & txtFecha & "'" & _
                " WHERE act_fij_codigo='" & Me.dcmbCodigo & "' " & _
                " AND emp_codigo='" & strEmpresa & "'"
        End If
        clsSql.Ejecutar strSql, "M"
        dcmbCodigo_Change
    End If
End Sub

Private Sub txtValor_Change()
txtValor.Text = Format(Val(txtValor.Text), "##0.00")
End Sub
Private Sub txtValor_dep_Change()
txtValor_dep.Text = Format(Val(txtValor_dep.Text), "##0.00")
End Sub
Private Sub CalcuTotal()
   'Calcula total
    Dim Subtotal As Double
    Total = 0
    For i = 1 To (VSFDep.Rows - 1)
        Total = Total + Val(VSFDep.TextMatrix(i, 4))
    Next i
    TxtTotal.Text = Val(Total)
End Sub
Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Integer
    VSFDep.Tag = "N"
    
    VSFDep.Clear 1
    VSFDep.Rows = 2
    VSFDep.Tag = "T"
    
End Sub

Private Sub txtValorR_Change()
    txtValorR.Text = Format(Val(txtValorR.Text), "##0.00")
End Sub
