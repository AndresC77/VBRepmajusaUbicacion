VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComprobanteEgresoComun 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobante de Egresos Comunes"
   ClientHeight    =   8910
   ClientLeft      =   1170
   ClientTop       =   1950
   ClientWidth     =   9720
   Icon            =   "frmComprobanteEgresoComun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9720
   Begin VB.CheckBox chkNombre2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "El cheque sale con otro nombre"
      Height          =   375
      Left            =   6720
      TabIndex        =   51
      Top             =   8400
      Width           =   2760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Egreso Común"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   9495
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   2160
         TabIndex        =   52
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
      End
      Begin VB.TextBox txtTotal 
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
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   7830
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   5670
         Width           =   7335
      End
      Begin VB.TextBox txtTotalHaber 
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
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   7830
         Width           =   1275
      End
      Begin VB.TextBox txtTotalDebe 
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
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   7830
         Width           =   1275
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
         Left            =   120
         TabIndex        =   1
         Top             =   1005
         Value           =   -1  'True
         Width           =   1335
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1005
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   4
         Top             =   990
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Cliente/Proveedor"
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
         Height          =   1935
         Left            =   120
         TabIndex        =   36
         Top             =   1470
         Width           =   9255
         Begin VB.CommandButton cmdpersona 
            Caption         =   "&Ver más datos"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtEmail 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtTelefono 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtDireccion 
            Enabled         =   0   'False
            Height          =   525
            Left            =   6600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtRuc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1080
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo dcmbNombre 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   705
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbBeneficiario 
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   345
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
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
            Left            =   5400
            TabIndex        =   42
            Top             =   757
            Width           =   405
         End
         Begin VB.Label lblTelefono 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono/fax:"
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
            Left            =   5400
            TabIndex        =   41
            Top             =   397
            Width           =   960
         End
         Begin VB.Label lbldireccion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
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
            Left            =   5400
            TabIndex        =   40
            Top             =   1117
            Width           =   720
         End
         Begin VB.Label lblruc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruc:"
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
            Left            =   600
            TabIndex        =   39
            Top             =   1117
            Width           =   330
         End
         Begin VB.Label lblnombre 
            AutoSize        =   -1  'True
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
            Height          =   210
            Left            =   600
            TabIndex        =   38
            Top             =   757
            Width           =   600
         End
         Begin VB.Label lblBeneficiario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beneficiario:"
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
            Left            =   600
            TabIndex        =   37
            Top             =   397
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Banco"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   28
         Top             =   3510
         Width           =   9255
         Begin VB.TextBox txtp 
            Height          =   285
            Left            =   6600
            TabIndex        =   20
            Top             =   1335
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtd 
            Height          =   285
            Left            =   6600
            TabIndex        =   18
            Top             =   855
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   1920
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtPrevisto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   1095
            Width           =   1815
         End
         Begin VB.TextBox txtDisponible 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   615
            Width           =   1815
         End
         Begin VB.TextBox txtSaldoReal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   255
            Width           =   1815
         End
         Begin VB.TextBox txtCheque 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   960
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbBanco 
            Height          =   315
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor del cheque:"
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
            Left            =   285
            TabIndex        =   35
            Top             =   1357
            Width           =   1275
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Previsto:"
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
            Left            =   4920
            TabIndex        =   34
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Disponible:"
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
            Left            =   4920
            TabIndex        =   33
            Top             =   630
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Real:"
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
            Left            =   4920
            TabIndex        =   32
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Left            =   360
            TabIndex        =   31
            Top             =   652
            Width           =   1245
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de cheque:"
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
            Left            =   360
            TabIndex        =   30
            Top             =   997
            Width           =   1095
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
            Left            =   360
            TabIndex        =   29
            Top             =   292
            Width           =   510
         End
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Hacer Comprobante Manualmente"
         Text            =   "Hacer Comprobante Manualmente"
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   6270
         Width           =   9240
         _cx             =   16298
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
         FormatString    =   $"frmComprobanteEgresoComun.frx":030A
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
      Begin MSDataListLib.DataCombo dcmbDescripcion 
         Height          =   315
         Left            =   6840
         TabIndex        =   3
         Top             =   630
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin NEED2.dtpFecha dtpFechaCH 
         Height          =   315
         Left            =   6840
         TabIndex        =   53
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Egreso Común:"
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
         TabIndex        =   49
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label lblDescripcion 
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
         Left            =   120
         TabIndex        =   48
         Top             =   5430
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
         Left            =   3600
         TabIndex        =   47
         Top             =   7845
         Width           =   855
      End
      Begin VB.Label lbldescripcion1 
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
         Left            =   5280
         TabIndex        =   46
         Top             =   675
         Width           =   900
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   7560
         Picture         =   "frmComprobanteEgresoComun.frx":03E4
         ToolTipText     =   "Elimina una Fila"
         Top             =   5670
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   7800
         Picture         =   "frmComprobanteEgresoComun.frx":051A
         Top             =   5670
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del comprobante:"
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
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label8 
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
         Left            =   5280
         TabIndex        =   44
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Cheque:"
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
         Height          =   285
         Left            =   5280
         TabIndex        =   43
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3233
      TabIndex        =   25
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4913
      TabIndex        =   26
      Top             =   8400
      Width           =   1575
   End
End
Attribute VB_Name = "frmComprobanteEgresoComun"
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
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private clsCom As New clsConsulta
Private clsPer As New clsConsulta
Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsEgr As New clsConsulta
Private strSQL As String
Dim Persona As String
Dim ff As Variant
Dim ffch As Variant
Dim m As String
Dim p As String
Private booCambiar As Boolean
Private intDato As Variant
Private numComp As String
Private numAsi As String

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
    Set clsCom = Nothing
    Set clsPer = Nothing
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsEgr = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub
Private Sub saldodisponible()
    'Calcula el saldo disponible de la cuenta bancaria
    
    strSQL = " SELECT  sum(com_egr_ch_valor) as valor" & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND com_egr_ch_estado = 'GIRADO' AND cta_ban_numero = '" & dcmbCuenta.Text & "' AND ban_codigo = '" & dcmbBanco.BoundText & "'  AND com_egr_ch_fecha <= CURRENT_TIMESTAMP " & _
             " GROUP BY cta_ban_numero "
    clsSql.Ejecutar strSQL
    
    If Not IsNull(clsSql.adorec_Def("valor")) And clsSql.adorec_Def.EOF = False Then
        Valor = clsSql.adorec_Def("valor")
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
    Else
        Valor = 0
        disponible = Val(txtSaldoReal) - Val(Valor)
        txtDisponible = disponible
        txtD = disponible
    End If
End Sub

Private Sub CalcuTotal()
   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    
    'Calcula total debe
    
    For i = 1 To VSFG.Rows - 1
        SumaDebe = SumaDebe + Val(VSFG.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    
    'Calcula total haber
    
    For i = 1 To VSFG.Rows - 1
        SumaHaber = SumaHaber + Val(VSFG.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")
    TxtTotal.Text = FormatoD2(txtTotalDebe.Text) - FormatoD2(txtTotalHaber.Text)
End Sub
Private Sub Limpiar()
    dcmbBanco = ""
    dcmbCuenta = ""
    dcmbBeneficiario = ""
    dcmbNombre = ""
    txtCheque = ""
    txtSaldoReal = ""
    txtPrevisto = ""
    txtDisponible = ""
    txtD = 0
    txtp = 0
    txtValor = 0
    txtValor.Text = FormatoD2(txtValor.Text)
    txtCodigo = txtCodigo + 1
    txtDescripcion = ""
    txtTotalDebe = 0
    'txtTotalDebe.Text = Formatod2(txtTotalDebe.Text)
    txtTotalHaber = 0
    TxtTotal = 0
    'txtTotalHaber.Text = Formatod2(txtTotalHaber.Text)
    VSFG.Clear 1
    VSFG.Row = 2

End Sub

Private Sub cmdAceptar_Click()
    Dim ff As String
    'Comprueba que todos los datos esten ingresados
    ff = Format(dtpFecha.Value, "yyyy-mm-dd")
    ffch = Format(dtpFechaCh.Value, "yyyy-mm-dd")
    If (IsDate(ff) = False) And (IsDate(ffch) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Egresos Comunes"
        Exit Sub
    End If
    
    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd
    
    a = VSFG.Rows - 1
    For i = 1 To a
        For j = i + 1 To a
            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) And VSFG.TextMatrix(i, 5) = VSFG.TextMatrix(j, 5) Then
                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
                VSFG.RemoveItem j
                a = a - 1
                j = j - 1
            End If
            If j >= a Then
                Exit For
            End If
        Next j
    Next i
    
    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "Comprobante de Egreso"
        txtValor.SetFocus
        Exit Sub
    Else
          'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If dcmbBanco = "" And dcmbCuenta = "" Then
            MsgBox "No estan ingresados los datos de la cuenta", vbInformation, "Comprobante de Egreso"
            dcmbBanco.SetFocus
            Exit Sub
        End If
        If txtCodigo = "" Or VSFG.TextMatrix(1, 1) = "" Or txtDescripcion = "" Or dcmbBeneficiario = "" Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "Ingreso"
            Exit Sub
        Else
            Dim Nombre2 As String
            If chkNombre2.Value = 1 Then
                Nombre2 = Trim(UCase(InputBox("Ingrese el nombre que saldrá en el cheque: ", "Nombre del cheque")))
                Me.txtDescripcion = UCase(Me.txtDescripcion) & " GIRADO A: " & Nombre2
            End If
            
            If Trim(Nombre2) = "" Then
                Nombre2 = "NULL"
            Else
                Nombre2 = "'" & Nombre2 & "'"
            End If
            If numComp = "" And numAsi = "" Then
                strSQL = " SELECT COALESCE(max(com_egr_codigo),0) as egr  " & _
                         " FROM comp_egreso " & _
                         " where emp_codigo='" & strEmpresa & "'" & _
                         " GROUP BY emp_codigo"
                clsEgr.Ejecutar strSQL
                txtCodigo.Text = clsEgr.adorec_Def("egr") + 1 'valor del código del egreso comun +1
            Else
                txtCodigo.Text = numComp
                strSQL = " DELETE FROM comp_egreso WHERE com_egr_codigo='" & txtCodigo.Text & "' and emp_codigo='" & strEmpresa & "' "
                clsSql.Ejecutar strSQL, "M"
            End If
            strSQL = " SELECT COALESCE(max(com_egr_codigo),0) as num " & _
                     " FROM comp_egreso" & _
                     " WHERE emp_codigo = '" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
            clsCom.Ejecutar strSQL
            
            Dim clsAsiento As New clsContable
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            
            If numComp = "" And numAsi = "" Then
                clsAsiento.NuevoAsiento "E", ff, 0, 0, FormatoD2(txtTotalHaber), "BENEFICIARIO: " & dcmbNombre.Text & vbNewLine & "BANCO: " & dcmbBanco.Text & " CTA: " & dcmbCuenta.Text & " CH.No: " & txtCheque.Text & vbNewLine & UCase(txtDescripcion)
            Else
                clsAsiento.NumAsiento = numAsi
                clsAsiento.ModificarAsiento FormatoD2(txtTotalHaber), FormatoD2(txtTotalHaber), ff, , , "BENEFICIARIO: " & dcmbNombre.Text & vbNewLine & "BANCO: " & dcmbBanco.Text & " CTA: " & dcmbCuenta.Text & " CH.No: " & txtCheque.Text & vbNewLine & UCase(txtDescripcion)
                'clsAsiento.NuevoAsiento "E", ffch, 0, 0, Formatod2(txtTotalDebe), "PAGO INCOMPLETO", False
                clsAsiento.EliminarAsiento False, True
            End If
            
            'Ingreso de datos en comp_egreso
            strSQL = " INSERT INTO comp_egreso (com_egr_codigo, emp_codigo,asi_numasiento, cta_ban_numero, ban_codigo, per_codigo, " & _
                 " com_egr_fecha, com_egr_descripcion, com_egr_ch_fecha,com_egr_ch_num, com_egr_ch_estado,com_egr_ch_valor,com_egr_conciliado, com_egr_nombre2, com_egr_fechamod, com_egr_usumod) " & _
                 " VALUES ('" & txtCodigo & "','" & strEmpresa & "','" & clsAsiento.NumAsiento & "','" & dcmbCuenta.BoundText & "','" & dcmbBanco.BoundText & "','" & dcmbBeneficiario.BoundText & "', " & _
                 " '" & ff & "','" & UCase(txtDescripcion) & "','" & ffch & "', '" & UCase(txtCheque) & "','GIRADO','" & txtValor & "'," & _
                 " 0," & Nombre2 & ",CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsSql.Ejecutar strSQL, "M"
            
            'ingreso de datos en el la tabla det_comp_egreso
            With VSFG
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                        clsAsiento.NuevoDetAsiento .TextMatrix(i, 1), .TextMatrix(i, 5), FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                    End If
                Next i
            End With
            
            'Actualiza los valores de los saldos
            Dim strChUlt As String
            If booCambiar = True Then
                strChUlt = " '" & txtCheque.Text & "'"
            Else
                strChUlt = " cta_ban_ch_ultimo"
            End If
            
            strSQL = " UPDATE cta_banco " & _
                     " SET cta_ban_ch_ultimo = " & strChUlt & ", cta_ban_saldoreal= '" & txtSaldoReal & "',cta_ban_saldoprevisto= '" & txtPrevisto & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                     " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
            clsSql.Ejecutar strSQL, "M"
            
            MsgBox " Los datos han sido ingresado", vbInformation, "Comprobantes"
            Dim CompEgr As New frmReporte
            CompEgr.strReporte = "rptComprobanteEgreso"
            CompEgr.strNumero = txtCodigo
            CompEgr.Show
            Dim Cheque As New frmReporte
            Cheque.strReporte = "rptCheque"
            Cheque.strNumero = txtCodigo
            Cheque.Show
            Set clsAsiento = Nothing
        End If
    End If
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdpersona_Click()
    If p = 0 Then
        frmPersona.Tag = "C"
        frmPersona.txtCodigo.Text = Me.dcmbBeneficiario.Text
        frmPersona.txtDireccion.Text = Me.txtDireccion
        frmPersona.txtRuc.Text = Me.txtRuc
        frmPersona.txtEmail.Text = Me.txtEmail
        frmPersona.Show
    ElseIf p = 1 Then
        frmPersona.Tag = "P"
        frmPersona.txtCodigo.Text = Me.dcmbBeneficiario.Text
        frmPersona.txtDireccion.Text = Me.txtDireccion
        frmPersona.txtRuc.Text = Me.txtRuc
        frmPersona.txtEmail.Text = Me.txtEmail
        frmPersona.Show
    End If
    
End Sub

Private Sub dcmbBanco_Change()
dcmbCuenta = ""
    dcmbBanco.Tag = dcmbBanco.BoundText
    
    strSQL = " SELECT cta_ban_numero, cta_ban_ch_ultimo as ban, cta_ban_ctaconta,cta_ban_saldoreal,cta_ban_saldodisponible,cta_ban_saldoprevisto" & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSQL
    If clsCtb.adorec_Def.EOF = False Then
        Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
        dcmbCuenta.ListField = ("cta_ban_numero")
    Else
        dcmbCuenta = ""
    End If
    
End Sub

Public Sub dcmbBeneficiario_Change()
Dim strComparar As String
    clsPer.Actualizar
    If clsPer.adorec_Def.RecordCount > 0 Then
        clsPer.adorec_Def.MoveFirst
    End If
    strComparar = "per_codigo =  '" & dcmbBeneficiario.Text & " '"
    clsPer.adorec_Def.Find strComparar
    dcmbBeneficiario.Tag = "A"
    If clsPer.adorec_Def.EOF = False Then
        dcmbNombre.Text = clsPer.adorec_Def("nombre")
        dcmbNombre.BoundText = dcmbBeneficiario.Text
        txtDireccion.Text = clsPer.adorec_Def("per_direccion")
        txtTelefono.Text = clsPer.adorec_Def("telefono")
        txtEmail.Text = clsPer.adorec_Def("per_email")
        txtRuc.Text = clsPer.adorec_Def("per_ruc")
        cmdpersona.Enabled = True
        Persona = clsPer.adorec_Def("per_codigo")
    Else
        dcmbNombre.Text = ""
        dcmbNombre.BoundText = ""
        txtDireccion.Text = ""
        txtTelefono.Text = ""
        txtEmail.Text = ""
        txtRuc.Text = ""
        cmdpersona.Enabled = False
    End If
    dcmbBeneficiario.Tag = ""
End Sub

Private Sub dcmbCodigo_Change()
 Dim strComparar As String
 Dim Fecha As Variant
    txtValor.Text = 0
    On Error GoTo errhandler
        If clsEgr.adorec_Def.RecordCount > 0 Then
            clsEgr.adorec_Def.MoveFirst
        End If
        strComparar = "egr_com_codigo = '" & dcmbCodigo.Text & "'"
        clsEgr.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsEgr.adorec_Def.EOF = False Then
            'pone valores en los combos
            dcmbDescripcion.Text = clsEgr.adorec_Def("descripcion")
            dcmbDescripcion.BoundText = clsEgr.adorec_Def("egr_com_codigo")
            dcmbBanco.BoundText = clsEgr.adorec_Def("ban_codigo")
            dcmbBanco.Text = clsEgr.adorec_Def("ban_nombre")
            dcmbCuenta.Text = clsEgr.adorec_Def("cta_ban_numero")
            txtDescripcion.Text = clsEgr.adorec_Def("egr_com_descripcion")
            Fecha = Format(clsEgr.adorec_Def("egr_com_fecha"), "yyyy-MM-dd")
            dd = Mid(Fecha, 9, 2)
            cmbDia.Text = dd
           'For i = 1 To VSFG.Rows - 1
            'pone valores en el grid
                strSQL = " SELECT distinct det_egreso_comun.cta_codigo,ctaconta.cta_nombre ,det_egr_com_debe, det_egr_com_haber " & _
                         " FROM ((( egreso_comun INNER JOIN det_egreso_comun ON egreso_comun.egr_com_codigo=det_egreso_comun.egr_com_codigo " & _
                         "                                                   AND egreso_comun.emp_codigo=det_egreso_comun.emp_codigo) " & _
                         "                       INNER JOIN ctaconta ON det_egreso_comun.cta_codigo= ctaconta.cta_codigo " & _
                         "                                           AND det_egreso_comun.emp_codigo= ctaconta.emp_codigo) " & _
                         "                       INNER JOIN cta_banco ON egreso_comun.cta_ban_numero=cta_banco.cta_ban_numero " & _
                         "                                            AND egreso_comun.ban_codigo=cta_banco.ban_codigo " & _
                         "                                            AND egreso_comun.emp_codigo=cta_banco.emp_codigo) " & _
                         " WHERE egreso_comun.egr_com_codigo = '" & dcmbCodigo & "' AND egreso_comun.emp_codigo = '" & strEmpresa & "'" & _
                         " ORDER BY if(det_egreso_comun.cta_codigo=cta_banco.cta_ban_ctaconta,0,1)"
                         
                clsDet.Ejecutar strSQL
               
            Set VSFG.DataSource = clsDet.adorec_Def.DataSource
            txtValor = VSFG.TextMatrix(1, 4)
            PonerBotones
            'Next i
            CalcuTotal
         Else
            dcmbDescripcion.Text = ""
            dcmbDescripcion.BoundText = ""
            dcmbCuenta.Text = ""
            dcmbBanco.Text = ""
            dcmbBanco.Tag = ""
            txtDescripcion.Text = ""
            txtTotalDebe.Text = 0
         
            txtTotalHaber.Text = 0
            TxtTotal = 0
            Set VSFG.DataSource = Nothing
            VSFG.Clear flexClearScrollable
            VSFG.Rows = 2
        End If
        dcmbCodigo.Tag = ""
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

Private Sub dcmbCodigo_LostFocus()
    If dcmbCodigo = "" Then
        dcmbCodigo.Text = "Hacer comprobante manualmente"
    End If
End Sub

Private Sub dcmbDescripcion_Change()
    
    If dcmbCodigo.Tag <> "A" Then
        If dcmbDescripcion.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbDescripcion.BoundText
        End If
        If dcmbDescripcion = "" Then
            dcmbCodigo.Text = dcmbDescripcion.Text
        End If
    End If
End Sub

Private Sub dcmbDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbDescripcion.BoundText
    End If
End Sub

Private Sub dcmbdescripcion_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    dcmbCodigo.Text = dcmbDescripcion.BoundText
End Sub

Private Sub dcmbNombre_Change()
    If dcmbBeneficiario.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbBeneficiario.Text = dcmbNombre.BoundText
        End If
        If dcmbNombre = "" Then
            dcmbBeneficiario.Text = dcmbNombre.Text
        End If
    End If
End Sub

Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbBeneficiario.Text = dcmbNombre.BoundText
    End If

End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  dcmbBeneficiario.Text = dcmbNombre.BoundText
End Sub

Private Sub Form_Activate()
 Dim strComparar As String
    
    On Error GoTo errhandler
  
    clsEgr.Actualizar
    If clsEgr.adorec_Def.RecordCount > 0 Then
        clsEgr.adorec_Def.MoveFirst
    End If
    strComparar = "egr_com_codigo = '" & dcmbCodigo.Text & "'"
    clsEgr.adorec_Def.Find strComparar
    dcmbCodigo.Tag = "A"
    
    If clsEgr.adorec_Def.EOF = False Then
    
    Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
    dcmbCodigo.ListField = "egr_com_codigo"
    Set dcmbDescripcion.RowSource = clsEgr.adorec_Def.DataSource
    dcmbDescripcion.ListField = "descripcion"
    dcmbDescripcion.BoundColumn = "egr_com_codigo"
    
'     consulta para saber los  bancos existentes
    strSQL = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL
    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        dcmbBanco = ""
    End If
        For i = 1 To VSFG.Rows - 1
        'pone valores en el grid
             strSQL = " SELECT det_egreso_comun.cta_codigo,ctaconta.cta_nombre ,det_egr_com_debe, det_egr_com_haber  " & _
                      " FROM det_egreso_comun INNER JOIN ctaconta ON det_egreso_comun.cta_codigo= ctaconta.cta_codigo " & _
                      "                                           AND det_egreso_comun.emp_codigo= ctaconta.emp_codigo" & _
                      " WHERE egr_com_codigo = '" & dcmbCodigo & "' " & _
                      " ORDER BY cta_codigo"
            clsDet.Ejecutar strSQL
            Set VSFG.DataSource = clsDet.adorec_Def.DataSource
            PonerBotones
        Next i
        CalcuTotal
    End If
    dcmbCodigo.Tag = ""
    Frame1.Caption = "Proveedor"
    clsPer.Actualizar
    If clsPer.adorec_Def.RecordCount > 0 Then
        clsPer.adorec_Def.MoveFirst
    End If
    strComparar = "per_codigo = '" & Persona & "'"
    clsPer.adorec_Def.Find strComparar
    dcmbBeneficiario.Tag = "A"
    If Not clsPer.adorec_Def.EOF Then
        Set dcmbNombre.RowSource = clsPer.adorec_Def.DataSource
        dcmbNombre.ListField = "nombre"
        dcmbNombre.BoundColumn = "per_codigo"
        dcmbNombre = clsPer.adorec_Def("nombre")
        txtDireccion.Text = clsPer.adorec_Def("per_direccion")
        txtTelefono.Text = clsPer.adorec_Def("telefono")
        txtEmail.Text = clsPer.adorec_Def("per_email")
        txtRuc.Text = clsPer.adorec_Def("per_ruc")
    End If
    dcmbBeneficiario.Tag = ""
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
'
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
    
   ' On Error GoTo errhandler
    'Inicializa las clases para hacer distintas consultas
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsCom.Inicializar AdoConn, AdoConnMaster
    clsPer.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster
    
    'Pone la fecha actual en los combos
    
    dtpFecha.Value = HoyDia
    dtpFechaCh.Value = HoyDia
    'cmbdiach.Text = d
    'Cmbañoch.Text = Y
    
    txtValor.Text = 0
    txtValor = FormatoD2(txtValor)
    strSQL = " SELECT COALESCE(max(com_egr_codigo),0) as num " & _
             " FROM comp_egreso" & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " GROUP BY emp_codigo"
    clsCom.Ejecutar strSQL
    If clsCom.adorec_Def.EOF Then
        txtCodigo.Text = 1
    Else
        txtCodigo.Text = clsCom.adorec_Def("num") + 1
    End If
    'Realiza la consulta para saber los códigos de los egresos comunes

    strSQL = " SELECT egr_com_codigo, egr_com_descripcion, CONCAT(SUBSTRING(egr_com_descripcion,1,20),'...') as descripcion,egr_com_fecha,egreso_comun.ban_codigo, cta_ban_numero,ban_nombre " & _
             " FROM (egreso_comun INNER JOIN banco ON egreso_comun.ban_codigo=banco.ban_codigo)" & _
             " WHERE egreso_comun.emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY egr_com_codigo"
    clsEgr.Ejecutar strSQL
    If clsEgr.adorec_Def.EOF = False Then
        Set dcmbCodigo.RowSource = clsEgr.adorec_Def.DataSource
        dcmbCodigo.ListField = "egr_com_codigo"
        Set dcmbDescripcion.RowSource = clsEgr.adorec_Def.DataSource
        dcmbDescripcion.ListField = "descripcion"
        dcmbDescripcion.BoundColumn = "egr_com_codigo"
    End If

'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cen_cos_codigo, cen_cos_nombre" & _
                 " FROM centro_costo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cen_cos_nombre"
     clsCta.Ejecutar strSQL
    
     VSFG.ColComboList(5) = VSFG.BuildComboList(clsCta.adorec_Def, "cen_cos_codigo, *cen_cos_nombre", "cen_cos_codigo")
'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSQL
    
     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre", "cta_codigo")

    
'    Consulta para sacar los bancos existentes en el combo
    strSQL = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL
    If clsBan.adorec_Def.EOF = False Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    End If
    'Seleccionamos el proveedor de la tabla persona (P), que esta por defecto
    Frame1.Caption = "Proveedor"
    strSQL = " SELECT per_codigo, COALESCE(per_direccion,'') as per_direccion, CONCAT(per_apellido,' ',per_nombre) as nombre,CONCAT(COALESCE(per_telf,''),'/',COALESCE(per_fax,'')) as telefono, COALESCE(per_email,'') as per_email, COALESCE(per_ruc,'') as per_ruc " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "per_codigo"
        Set dcmbNombre.RowSource = clsPer.adorec_Def.DataSource
        dcmbNombre.ListField = "nombre"
        dcmbNombre.BoundColumn = "per_codigo"
        Persona = ""
        p = 1
    End If
    txtp = 0
    txtD = 0
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

Private Sub OptCliente_Click()
    p = 0
    Frame1.Caption = "Cliente"
    dcmbBeneficiario.Text = ""
    dcmbNombre.Text = ""
    strSQL = " SELECT per_codigo, COALESCE(per_direccion,'') as per_direccion, CONCAT(per_apellido,' ',per_nombre) as nombre,CONCAT(COALESCE(per_telf,''),'/',COALESCE(per_fax,'')) as telefono, COALESCE(per_email,'') as per_email, COALESCE(per_ruc,'') as per_ruc " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'C' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "per_codigo"
        Set dcmbNombre.RowSource = clsPer.adorec_Def.DataSource
        dcmbNombre.ListField = "nombre"
        dcmbNombre.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub optproveedor_Click()
    p = 1
    Frame1.Caption = "Proveedor"
    dcmbBeneficiario.Text = ""
    dcmbNombre.Text = ""
    strSQL = " SELECT per_codigo, COALESCE(per_direccion,'') as per_direccion, CONCAT(per_apellido,' ',per_nombre) as nombre,CONCAT(COALESCE(per_telf,''),'/',COALESCE(per_fax,'')) as telefono, COALESCE(per_email,'') as per_email, COALESCE(per_ruc,'') as per_ruc " & _
             " FROM persona " & _
             " WHERE emp_codigo= '" & strEmpresa & "' AND cat_p_tipo = 'P' " & _
             " ORDER BY per_apellido,per_nombre"
    clsPer.Ejecutar strSQL
    If clsPer.adorec_Def.EOF = False Then
        Set dcmbBeneficiario.RowSource = clsPer.adorec_Def.DataSource
        dcmbBeneficiario.ListField = "per_codigo"
        Set dcmbNombre.RowSource = clsPer.adorec_Def.DataSource
        dcmbNombre.ListField = "nombre"
        dcmbNombre.BoundColumn = "per_codigo"
        p = 1
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtValor_LostFocus()
    
    VSFG.TextMatrix(1, 4) = txtValor.Text
    txtValor.Text = FormatoD2(txtValor.Text)
    CalcuTotal
    
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
' Verifica si el dato uçingresado es numérico
    If IsNumeric(txtValor.Text) = False Then
        MsgBox "Solo se permiten valores numéricos", vbOKOnly + vbInformation, "Comprobante de Egreso Común"
        txtValor.Text = 0
        txtValor.Text = FormatoD2(txtValor.Text)
        Cancel = True
    Else
        ' Pone dos decimales al valor
        txtValor.Text = FormatoD2(txtValor.Text)
        Cancel = False
    End If

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

    If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub

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
        If r > 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Egresos Comunes"   ' Define el título.
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

Private Sub dcmbCuenta_Change()
    If Trim(dcmbCuenta) <> "" Then
        strSQL = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre,cta_ban_ch_ultimo as cheque,cta_ban_saldoreal,cta_ban_saldoprevisto " & _
                 " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo " & _
                 "                                    AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
        clsctc.Ejecutar strSQL
        If Not clsctc.adorec_Def.EOF Then
            If Trim(clsctc.adorec_Def("cheque")) = "" Or IsNull(clsctc.adorec_Def("cheque")) Then
                txtCheque.Text = 1
            Else
                txtCheque.Text = clsctc.adorec_Def("cheque") + 1
            End If
            Dim strNCH As String
            Dim booPasar As Boolean
            booPasar = False
            strNCH = txtCheque.Text
            numComp = ""
            numAsi = ""
            While booPasar = False
                txtCheque.Text = InputBox("No. de cheque", "Comprobante de Egreso", strNCH)
                strSQL = " SELECT count(*) as Num FROM comp_egreso " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND ban_codigo='" & dcmbBanco.BoundText & "'" & _
                         " AND cta_ban_numero = '" & dcmbCuenta & "'" & _
                         " AND com_egr_ch_num='" & txtCheque.Text & "'"
                clsSql.Ejecutar strSQL
                If clsSql.adorec_Def("Num") <> 0 Then
                    strSQL = " SELECT com_egr_codigo,asi_numasiento,com_egr_ch_estado " & _
                             " FROM comp_egreso " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND ban_codigo='" & dcmbBanco.BoundText & "'" & _
                             " AND cta_ban_numero = '" & dcmbCuenta & "'" & _
                             " AND com_egr_ch_num='" & txtCheque.Text & "'"
                    clsSql.Ejecutar strSQL
                    If clsSql.adorec_Def("com_egr_ch_estado") = "ANULADO" Then
                        If MsgBox("El cheque tiene estado Anulado." & vbNewLine & "Desea reutilizar el cheque y el compobante?", vbYesNo + vbQuestion, "Comprobante de Egreso") = vbYes Then
                            numComp = clsSql.adorec_Def("com_egr_codigo")
                            numAsi = clsSql.adorec_Def("asi_numasiento")
                            booPasar = True
                        End If
                    Else
                        MsgBox "Ese cheque ya ha sido emitido", vbCritical, "Comprobante de Egreso"
                        txtCheque.Text = strNCH
                        numComp = ""
                        numAsi = ""
                    End If
                Else
                    booPasar = True
                End If
            Wend
            If Format(txtCheque.Text, "0000000000") >= Format(strNCH, "0000000000") Then
                booCambiar = True
            Else
                booCambiar = False
            End If
            txtSaldoReal.Text = clsctc.adorec_Def("cta_ban_saldoreal")
            txtPrevisto.Text = clsctc.adorec_Def("cta_ban_saldoprevisto")
            txtp = clsctc.adorec_Def("cta_ban_saldoprevisto")
            If clsctc.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(1, 1) = clsctc.adorec_Def("cta_ban_ctaconta")
                VSFG.TextMatrix(1, 2) = clsctc.adorec_Def("cta_nombre")
            End If
            saldodisponible
        End If
        txtValor.Enabled = True
    Else
        txtSaldoReal = 0
        txtPrevisto = 0
        txtDisponible = 0
        txtp = 0
        txtD = 0
        txtCheque = ""
        txtValor = 0
        n = 4
        a = VSFG.Rows - 1
        
        VSFG.Clear 1
        
'        For i = 1 To n
'            VSFG.TextMatrix(1, i) = ""
'        Next i
      
    End If
    'txtSaldoReal.Text = Formatod2(txtSaldoReal.Text)
    'txtDisponible.Text = Formatod2(txtDisponible.Text)
    'txtPrevisto.Text = Formatod2(txtPrevisto.Text)
    'txtValor.Text = Formatod2(txtValor.Text)
End Sub

Private Sub VSFG_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <> OldRow Then
        If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
            If VSFG.Row < VSFG.Rows Then
                If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes
    
    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            
            If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
                If Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "4" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "5" And Left(VSFG.TextMatrix(VSFG.Row, 1), 1) <> "6" Then
                    VSFG.AddItem ""
                ElseIf VSFG.TextMatrix(VSFG.Row, 5) <> "" Then
                    VSFG.AddItem ""
                End If
            Else
                VSFG.AddItem ""
            End If
            
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 1 Then
        If Col = 1 Then
            Cancel = True
        End If
        If Col = 2 Then
            Cancel = True
        End If
        If Col = 3 Then
            Cancel = True
        End If
        If Col = 4 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
            MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
            VSFG.TextMatrix(Row, 3) = ""
            VSFG.TextMatrix(Row, 4) = ""
            
           
        ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe

        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = intDato
        End If

        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = intDato
        End If
    CalcuTotal
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)

' filtra el nombre y codigo de cuenta para los combos del greed
If Row > 1 Then
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                     .TextMatrix(Row, 2) = .TextMatrix(Row, 1)
             End If

             If Col = 2 Then
                     .TextMatrix(Row, 1) = .TextMatrix(Row, 2)
             End If
         End If
    End With
End If
End Sub

Private Sub txtDisponible_Change()
    txtDisponible = FormatoD2(txtDisponible)
End Sub

Private Sub txtPrevisto_Change()
    txtPrevisto = FormatoD2(txtPrevisto)
End Sub

Private Sub txtsaldoReal_Change()
    txtSaldoReal = FormatoD2(txtSaldoReal)
End Sub

Private Sub VSFG_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            Set frmSelecCtaConta.objEscribir = VSFG
            frmSelecCtaConta.Show vbModal
        End If
    End If
End Sub

Private Sub VSFG_Validate(Cancel As Boolean)
    If Year(dtpFecha.Value) >= 2018 And VSFG.Rows > 1 Then
        If (Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "4" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "5" Or Left(VSFG.TextMatrix(VSFG.Row, 1), 1) = "6") And VSFG.TextMatrix(VSFG.Row, 5) = "" Then
            Cancel = True
        End If
    End If
End Sub
