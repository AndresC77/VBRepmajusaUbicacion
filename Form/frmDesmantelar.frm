VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDesmantelar 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos a Desmantelar"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmDesmantelar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10695
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14540253
      TabCaption(0)   =   "Desmantelar"
      TabPicture(0)   =   "frmDesmantelar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAceptar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Consulta Datos de Pedido A Desmantelar"
      TabPicture(1)   =   "frmDesmantelar.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkDesmantelaHoy"
      Tab(1).Control(1)=   "VSFGN1"
      Tab(1).Control(2)=   "VSFGLider"
      Tab(1).Control(3)=   "cmdEnviarCorreos"
      Tab(1).Control(4)=   "cmdConsultaAutomatica"
      Tab(1).Control(5)=   "cmdMas"
      Tab(1).Control(6)=   "txtPedidos"
      Tab(1).Control(7)=   "uctrVSFG1"
      Tab(1).Control(8)=   "cmbNegocio2"
      Tab(1).Control(9)=   "VSFGCliente"
      Tab(1).Control(10)=   "VSFGPeds"
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(12)=   "Label1"
      Tab(1).ControlCount=   13
      Begin VB.CheckBox chkDesmantelaHoy 
         Caption         =   "Desmantelar HOY"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -66360
         TabIndex        =   59
         Top             =   960
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGN1 
         Height          =   2415
         Left            =   -72000
         TabIndex        =   58
         Top             =   4800
         Visible         =   0   'False
         Width           =   4140
         _cx             =   7302
         _cy             =   4260
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesmantelar.frx":0342
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGLider 
         Height          =   2415
         Left            =   -73320
         TabIndex        =   57
         Top             =   4800
         Visible         =   0   'False
         Width           =   4740
         _cx             =   8361
         _cy             =   4260
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesmantelar.frx":0490
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
      Begin VB.CommandButton cmdEnviarCorreos 
         Caption         =   "Enviar Correos"
         Height          =   375
         Left            =   -68640
         TabIndex        =   55
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdConsultaAutomatica 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   -65835
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame5 
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
         Height          =   1335
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   9975
         Begin VB.TextBox txtAutorizacion 
            Height          =   285
            Left            =   4800
            TabIndex        =   7
            Top             =   990
            Width           =   1815
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Left            =   4800
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   4800
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtCaduca 
            Height          =   285
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   960
            Width           =   1060
         End
         Begin VB.TextBox txtDcto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   1335
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
            Format          =   66125827
            CurrentDate     =   37463
         End
         Begin MSDataListLib.DataCombo CmbFpago 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
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
            TabIndex        =   9
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
         Begin NEED2.dtpFecha dtpFecha 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   1815
            _extentx        =   2990
            _extenty        =   450
            value           =   41810.5463773148
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00BAA892&
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
            Left            =   720
            TabIndex        =   50
            Top             =   255
            Width           =   495
         End
         Begin VB.Label Label10 
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
            TabIndex        =   49
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label12 
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
            TabIndex        =   48
            Top             =   990
            Width           =   915
         End
         Begin VB.Label Label13 
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
            TabIndex        =   47
            Top             =   600
            Width           =   555
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
            TabIndex        =   46
            Top             =   1005
            Width           =   1080
         End
         Begin VB.Label Label14 
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
            TabIndex        =   45
            Top             =   990
            Width           =   1350
         End
         Begin VB.Label Label15 
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
            TabIndex        =   44
            Top             =   630
            Width           =   780
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
            TabIndex        =   43
            Top             =   240
            Width           =   720
         End
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
         Height          =   360
         Left            =   -66480
         TabIndex        =   25
         Top             =   487
         Width           =   375
      End
      Begin VB.TextBox txtPedidos 
         Height          =   285
         Left            =   -68640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   525
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   9975
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   255
            Left            =   7200
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSDataListLib.DataCombo cmbFactura 
            Height          =   315
            Left            =   7200
            TabIndex        =   2
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
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
         Begin MSDataListLib.DataCombo cmbCliente 
            Height          =   315
            Left            =   840
            TabIndex        =   1
            Top             =   600
            Width           =   5055
            _ExtentX        =   8916
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
         Begin MSDataListLib.DataCombo cmbNegocio 
            Height          =   315
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
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
            Left            =   120
            TabIndex        =   51
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lblCliente 
            AutoSize        =   -1  'True
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
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
            TabIndex        =   21
            Top             =   645
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Factura:"
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
            Left            =   6240
            TabIndex        =   20
            Top             =   645
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   7080
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5070
         TabIndex        =   18
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Frame Frame3 
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
         Height          =   4215
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   9975
         Begin VB.Frame Frame2 
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
            Height          =   2055
            Left            =   480
            TabIndex        =   27
            Top             =   2040
            Width           =   8895
            Begin VB.TextBox TxtObserv 
               Height          =   555
               Left            =   1080
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   13
               Top             =   1440
               Width           =   3375
            End
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
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   240
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
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   32
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
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   31
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
               Left            =   7560
               TabIndex        =   30
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
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   29
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
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   1320
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFGReca 
               Height          =   1095
               Left            =   240
               TabIndex        =   12
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
               FormatString    =   $"frmDesmantelar.frx":05DE
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
            Begin MSDataListLib.DataCombo dcmbIVA 
               Height          =   315
               Left            =   5160
               TabIndex        =   60
               Top             =   1080
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
               Left            =   4560
               TabIndex        =   61
               Top             =   1110
               Width           =   510
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observ:"
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
               TabIndex        =   40
               Top             =   1515
               Width           =   585
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
               Left            =   6660
               TabIndex        =   39
               Top             =   270
               Width           =   675
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
               Left            =   6660
               TabIndex        =   38
               Top             =   1710
               Width           =   450
            End
            Begin VB.Label Label3 
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
               Left            =   6660
               TabIndex        =   37
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
               Left            =   6660
               TabIndex        =   36
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
               Left            =   6660
               TabIndex        =   35
               Top             =   1110
               Width           =   570
            End
            Begin VB.Label Label6 
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
               Left            =   6660
               TabIndex        =   34
               Top             =   630
               Width           =   630
            End
         End
         Begin VSFlex8LCtl.VSFlexGrid vsfgDetalle 
            Height          =   1725
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   9735
            _cx             =   17171
            _cy             =   3043
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   275
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmDesmantelar.frx":065E
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
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   7080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin NEED2.uctrVSFG uctrVSFG1 
         Height          =   375
         Left            =   -74880
         TabIndex        =   23
         Top             =   960
         Width           =   4815
         _extentx        =   8493
         _extenty        =   661
      End
      Begin MSDataListLib.DataCombo cmbNegocio2 
         Height          =   315
         Left            =   -74040
         TabIndex        =   53
         Top             =   510
         Width           =   4095
         _ExtentX        =   7223
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGCliente 
         Height          =   2415
         Left            =   -74640
         TabIndex        =   56
         Top             =   4800
         Visible         =   0   'False
         Width           =   4740
         _cx             =   8361
         _cy             =   4260
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesmantelar.frx":07B3
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   22
         Top             =   1320
         Width           =   10140
         _cx             =   17886
         _cy             =   10821
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesmantelar.frx":0901
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
      Begin VB.Label Label9 
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
         Left            =   -74760
         TabIndex        =   54
         Top             =   562
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carga Pedidos:"
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
         Left            =   -69720
         TabIndex        =   26
         Top             =   562
         Width           =   1095
      End
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   240
      Picture         =   "frmDesmantelar.frx":0B01
      Top             =   6120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   0
      Picture         =   "frmDesmantelar.frx":0C2D
      Top             =   6120
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmDesmantelar"
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
Private clsCon_Prd2 As New clsConsulta
Private strSql As String
Private ValDias As Long
Private booTodaFactura As Boolean

Private Sub cmbCliente_Change()
    ValDias = 30
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    cmbFactura = ""
    If ValDias = 0 Then ValDias = 30
    
    If booTodaFactura = False Then
        strSql = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo as fac , cue_p_c_valor," & _
                 " cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d," & _
                 " persona.ven_codigo,persona.for_pag_codigo " & _
                 " FROM  (cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C'" & _
                 " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo)" & _
                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE cuenta_p_c.per_codigo='" & cmbCliente.BoundText & "' AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                 " AND tip_doc_cue_codigo=1 AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_egr_codigo,cue_p_c_valor,com_ret_total,persona.ven_codigo,persona.for_pag_codigo " & _
                 " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)=round(cue_p_c_valor,2) " & _
                 " ORDER BY cue_p_c_egr_codigo "
    Else
        strSql = " SELECT egr_codigo as fac , persona.ven_codigo,persona.for_pag_codigo " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C'" & _
                 " WHERE egreso.per_codigo='" & cmbCliente.BoundText & "' AND egreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' AND egr_anulado = 0 " & _
                 " ORDER BY egr_codigo, persona.ven_codigo,persona.for_pag_codigo "
    
    End If
    
'    strSql = " SELECT cue_p_c_egr_codigo as fac " & _
'             " FROM cuenta_p_c " & _
'             " Where emp_codigo='" & strEmpresa & "' And tip_egr_codigo='FAC' " & _
'             " AND egr_anulado=0 " & _
'             " AND egr_fecha>='" & DateAdd("d", -1 * ValDias, HoyDia) & "' AND per_codigo='" & cmbCliente.BoundText & "'" & _
'             " UNION " & _
'             " SELECT concat('R',egr_codigo) as fac " & _
'             " FROM factura_ryb " & _
'             " Where emp_codigo='" & strEmpresa & "' And tip_egr_codigo='FAC' " & _
'             " AND egr_anulado=0 " & _
'             " AND egr_fecha>='" & DateAdd("d", -1 * ValDias, HoyDia) & "' AND per_codigo='" & cmbCliente.BoundText & "'" & _
'             " ORDER BY fac "
    clsCon_Def.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbVendedor.BoundText = clsCon_Def.adorec_Def("ven_codigo")
        CmbFpago.BoundText = clsCon_Def.adorec_Def("for_pag_codigo")
    End If
    Set cmbFactura.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbFactura.ListField = "fac"
        cmbFactura.BoundColumn = "fac"
    Else
        cmbFactura = "No hay facturas del cliente "
    End If
End Sub

Private Sub RevisarDatos()
    Dim clsCon_TipDoc As New clsConsulta
    clsCon_TipDoc.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT tip_ing_codigo, tip_ing_nombre,tip_ing_impuesto,tip_ing_persona,tip_ing_cx_p_c,tip_ing_recargo,tip_ing_numsri,tip_ing_cos_pre,tip_ing_retencion " & _
             " FROM tipo_ingreso WHERE emp_codigo = '" & strEmpresa & "' AND tip_ing_codigo='DCL'"
    clsCon_TipDoc.Ejecutar strSql
    
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
    strSql = " SELECT TOP 1 COALESCE(ing_serie,'') as ing_serie,COALESCE(ing_numero,'0')+1 as ing_numero,COALESCE(ing_autorizacion,'') as ing_autorizacion,COALESCE(ing_caduca,'00/0000') as ing_caduca " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND ing_anulado=0" & _
             " AND tip_ing_codigo='DCL' " & _
             " AND ing_codigo LIKE '" & FormatoD0(strPtoFactura) & "%' AND LEN(ing_codigo)>10 " & _
             " ORDER BY ing_fecha DESC,ing_numero DESC,ing_codigo DESC "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        txtSerie.Text = clsCon_Def.adorec_Def("ing_serie")
        txtDocumento.Text = clsCon_Def.adorec_Def("ing_numero")
        txtAutorizacion.Text = clsCon_Def.adorec_Def("ing_autorizacion")
        If clsCon_Def.adorec_Def("ing_caduca") <> "00/0000" Then
            If clsCon_Def.adorec_Def("ing_caduca") <> "" Then
                dtpCaduca.Value = clsCon_Def.adorec_Def("ing_caduca")
            End If
            txtCaduca.Text = clsCon_Def.adorec_Def("ing_caduca")
        Else
            dtpCaduca.Value = Format(HoyDia, "mm\/yyyy")
            txtCaduca.Text = ""
        End If
    Else
        txtSerie.Text = ""
        txtCaduca.Text = ""
        txtDocumento.Text = ""
        txtAutorizacion.Text = ""
        dtpCaduca.Value = Format(HoyDia, "mm\/yyyy")
    End If
    txtSerie.Locked = False
    txtDocumento.Locked = False
    txtAutorizacion.Locked = False
    dtpCaduca.Enabled = True
    cmbVendedor.Visible = True
    lblV.Visible = True

End Sub

Private Sub cmbFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If MsgBox("Desea revisar todas las facturas del cliente?", vbQuestion + vbYesNo + vbDefaultButton2, "Ajustes") = vbYes Then
            frmClave.strClaveMAESTRA = "sapo"
            frmClave.Show vbModal
            If frmClave.Ret = False Then
                booTodaFactura = False
            Else
                booTodaFactura = True
            End If
        Else
            booTodaFactura = False
        End If
        cmbCliente_Validate False
    End If
End Sub

Private Sub cmbFactura_Validate(Cancel As Boolean)
    CargaProductos
    RevisarDatos
End Sub

Private Sub cmbNegocio_Change()
    If cmbNegocio.BoundText <> "" Then
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            If strPtoFactura <> clsCon_Def.adorec_Def(0) Then
                LimpiarTodo
            End If
            strPtoFactura = clsCon_Def.adorec_Def(0)
        End If
    Else
        Exit Sub
    End If
    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
End Sub

Private Sub LimpiarTodo()
    cmbCliente.BoundText = ""
End Sub

Private Sub cmdAceptar_Click()
    IngresaNotaCredito
    booTodaFactura = False
End Sub
Private Sub IngresaNotaCredito()

    Dim clsIngreso As New clsInventario
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
    
    If txtSerie.Locked = False And txtDocumento.Locked = False And txtAutorizacion.Locked = False Then
        If Trim(txtSerie.Text) = "" Or Trim(txtDocumento.Text) = "" Or Trim(txtAutorizacion.Text) = "" Then
            MsgBox "Llene los campos del Documento", vbInformation, "Documento"
            Exit Sub
        End If
    End If
    If CmbFpago.Text = "" Then
        MsgBox "Llene Forma de Pago", vbInformation, "Documento"
        Exit Sub
    End If
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    'NOTA DE CREDITO
    
    
    strSql = " SELECT tip_ped_ptofac " & _
             " FROM tipo_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "
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
    
    TxtObserv.Text = "DESMANTELADO FACTURA " & cmbFactura.BoundText & vbNewLine & TxtObserv.Text
    booGuardar = clsIngreso.NuevoIng(True, "DCL", True, Left(txtSerie.Text, 3), Right(txtSerie.Text, 3), txtDocumento.Text, CmbFpago.BoundText, Me.cmbCliente.BoundText, dtpFecha.Value, cmbFactura.BoundText, , UCase(TxtObserv.Text), , txtAutorizacion.Text, txtCaduca.Text, FormatoD2(TxtSubTotal.Text), FormatoD2(TxtRecargo.Text), FormatoD2(TxtDesc.Text), FormatoD2(TxtIva.Text), FormatoD2(TxtTotal.Text), , , , dcmbIVA.BoundText)
    If booGuardar = True Then
        strTipCompAsiento = "A"
        strObserv = UCase("NOTA DE CREDITO" & clsIngreso.strDoc & vbNewLine & "PERSONA: " & cmbCliente.Text & vbNewLine & "DOCUMENTO: " & txtSerie.Text & Format(txtDocumento.Text, "0000000") & vbNewLine & TxtObserv.Text)
        IngAsi = True
        If IngAsi = True Then
            clsAsiento.Inicializar AdoConn, AdoConnMaster
            clsAsiento.NuevoAsiento strTipCompAsiento, dtpFecha.Value, 0, 0, TxtTotal.Text, strObserv
            clsIngreso.ModificaIng , , , , , , clsAsiento.NumAsiento
            NumeroAsiento = clsAsiento.NumAsiento
        End If
        
        With vsfgDetalle
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 10) = "" Then
                    clsIngreso.IngresaAutoAContenedor = False
                    InicializarContenedorRecurrente
                Else
                    booUnContenedor = True
                    strContenedorRecurrente = .TextMatrix(i, 10)
                End If
                clsIngreso.NuevoDetIng .TextMatrix(i, 2), .TextMatrix(i, 1), FormatoD4(.TextMatrix(i, 4)), FormatoD8(.TextMatrix(i, 5)), FormatoD4(.TextMatrix(i, 8)), FormatoD4(.TextMatrix(i, 6)), Abs(FormatoD0(.TextMatrix(i, 9)))
                InicializarContenedorRecurrente
            Next i
            
        End With
        With VSFGReca
            For i = 1 To .Rows - 1
                clsIngreso.NuevoDetIngRecargo .TextMatrix(i, 1), FormatoD2(.TextMatrix(i, 3))
            Next i
        End With
        'clsIngreso.DetRetenciones
        
        strSql = " UPDATE pedido SET ped_estado=2, " & _
                 " ped_usumod='" & strUsuario & "', " & _
                 " ped_fechamod=CURRENT_TIMESTAMP" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_tip_egr_codigo='FAC' " & _
                 " AND ped_egr_codigo=" & cmbFactura.BoundText & " "
        clsCon_Def.Ejecutar strSql, "M"
        If IngAsi = True And CmbFpago.Visible = True Then
            'clsFPago.adorec_Def.MoveFirst
            'strComparar = "for_pag_codigo = '" & CmbFpago.BoundText & "'"
            'clsFPago.adorec_Def.Find strComparar
            'Inserta un nuevo registro de la cuenta por cobrar*/
            clsCta.Inicializar AdoConn, AdoConnMaster
            
            clsCta.IngAsientoIng clsAsiento, clsIngreso
            
            'Aplica la Nota de Credito
            clsCta.strPersona = clsIngreso.strPersona
            clsCta.strTipoCta = "C"
            clsCta.AplicaNC clsIngreso.strDoc, , dtpFecha.Value, cmbFactura.BoundText
            DocElectronico "04", (clsIngreso.strDoc)
            MsgBox " Los datos han sido ingresado", vbInformation, "Ingresos"
            Set clsCta = Nothing
            Set clsAsiento = Nothing
        End If
        
        
        Dim rpTNC As New frmReporte
        rpTNC.strNumero = clsIngreso.strDoc
        rpTNC.strReporte = "rptNotaCredito"
        rpTNC.Show
        
        
        Dim rpTNC2 As New frmReporte
        rpTNC2.strNumero = clsIngreso.strDoc
        rpTNC2.strReporte = "rptNotaCreditoUbicacion"
        rpTNC2.Show
        
'        Dim rpMov1 As New frmReporte
'        rpMov1.strNumero = clsIngreso.strDoc
'        rpMov1.strTipo = clsIngreso.strTipo
'        rpMov1.strReporte = "rptDetalleAdjunto"
'        rpMov1.Show

        Unload Me
    End If


End Sub



Private Sub cmdConsultaAutomatica_Click()
    CargaGrid "A"
End Sub

Private Sub cmdEnviarCorreos_Click()
    Dim i As Long
    Dim booEnviarCliente As Boolean
    Dim booEnviarLider As Boolean
    Dim booEnviarN1 As Boolean
    'CLIENTES
    booEnviarCliente = False
    booEnviarLider = False
    booEnviarN1 = False
    VSFGCliente.Rows = 1
    VSFGCliente.Clear flexClearScrollable
    VSFGLider.Rows = 1
    VSFGLider.Clear flexClearScrollable
    VSFGN1.Rows = 1
    VSFGN1.Clear flexClearScrollable
    For i = 1 To VSFGPeds.Rows - 1
        booEnviarCliente = False
        booEnviarLider = False
        booEnviarN1 = False
        'CLIENTE
        If Trim(VSFGPeds.TextMatrix(14, i)) <> "" Then
            VSFGCliente.AddItem VSFGPeds.TextMatrix(i, 0) & vbTab & _
                                VSFGPeds.TextMatrix(i, 2) & vbTab & _
                                VSFGPeds.TextMatrix(i, 3) & vbTab & _
                                VSFGPeds.TextMatrix(i, 4) & vbTab & _
                                VSFGPeds.TextMatrix(i, 5) & vbTab & _
                                VSFGPeds.TextMatrix(i, 6) & vbTab & _
                                VSFGPeds.TextMatrix(i, 7) & vbTab & _
                                VSFGPeds.TextMatrix(i, 8) & vbTab & _
                                VSFGPeds.TextMatrix(i, 9) & vbTab & _
                                VSFGPeds.TextMatrix(i, 10) & vbTab & _
                                VSFGPeds.TextMatrix(i, 11)
            If i + 1 < VSFGPeds.Rows Then
                booEnviarCliente = True
            End If
            If booEnviarCliente = False Then
                If Trim(VSFGPeds.TextMatrix(14, i)) <> Trim(VSFGPeds.TextMatrix(14, i + 1)) Then
                    booEnviarCliente = True
                End If
            End If
            If booEnviarCliente = True Then
                VSFGCliente.SaveGrid "PedidosPendientes.xls", flexFileExcel, flexXLSaveFixedCells

                EnviarMail NombreComercial & " Ajustes", CorreoServicioAlCliente, VSFGPeds.TextMatrix(3, i), Trim(VSFGPeds.TextMatrix(14, i)), "", "Reporte de pedidos pendientes de pago ", _
                        "Estimad@" & vbNewLine & _
                        VSFGPeds.TextMatrix(3, i) & vbNewLine & _
                        "Adjunto encontrará el listado de pedidos pendientes de pago." & vbNewLine & _
                        "Recuerde realizar el pago en el banco a través  del código identificado." & vbNewLine & _
                        "Saludos Cordiales" & vbNewLine & _
                        "Departamento de Ajustes" & vbNewLine & _
                        NombreComercial, "PedidosPendientes.xls"
                Kill "PedidosPendientes.xls"
                VSFGCliente.Rows = 1
                VSFGCliente.Clear flexClearScrollable
            End If
        End If
        'LIDER
        If Trim(VSFGPeds.TextMatrix(16, i)) <> "" Then
            VSFGLider.AddItem VSFGPeds.TextMatrix(i, 0) & vbTab & _
                VSFGPeds.TextMatrix(i, 2) & vbTab & _
                VSFGPeds.TextMatrix(i, 3) & vbTab & _
                VSFGPeds.TextMatrix(i, 4) & vbTab & _
                VSFGPeds.TextMatrix(i, 5) & vbTab & _
                VSFGPeds.TextMatrix(i, 6) & vbTab & _
                VSFGPeds.TextMatrix(i, 7) & vbTab & _
                VSFGPeds.TextMatrix(i, 8) & vbTab & _
                VSFGPeds.TextMatrix(i, 9) & vbTab & _
                VSFGPeds.TextMatrix(i, 10) & vbTab & _
                VSFGPeds.TextMatrix(i, 11)
            If i + 1 < VSFGPeds.Rows Then
                booEnviarLider = True
            End If
            If booEnviarLider = False Then
                If Trim(VSFGPeds.TextMatrix(16, i)) <> Trim(VSFGPeds.TextMatrix(16, i + 1)) Then
                    booEnviarLider = True
                End If
            End If
            If booEnviarLider = True Then
                VSFGLider.SaveGrid "PedidosPendientesLi.xls", flexFileExcel, flexXLSaveFixedCells

                EnviarMail NombreComercial & " Ajustes", CorreoServicioAlCliente, VSFGPeds.TextMatrix(7, i), Trim(VSFGPeds.TextMatrix(16, i)), "", "Reporte de pedidos pendientes de pago ", _
                        "Estimad@" & vbNewLine & _
                        VSFGPeds.TextMatrix(7, i) & vbNewLine & _
                        "Adjunto encontrará el listado de pedidos pendientes de pago." & vbNewLine & _
                        "Recuerde realizar el pago en el banco a través  del código identificado." & vbNewLine & _
                        "Saludos Cordiales" & vbNewLine & _
                        "Departamento de Ajustes" & vbNewLine & _
                        NombreComercial, "PedidosPendientesLi.xls"
                Kill "PedidosPendientesLi.xls"
                VSFGLider.Rows = 1
                VSFGLider.Clear flexClearScrollable
            End If
        End If
        'N1
        If Trim(VSFGPeds.TextMatrix(15, i)) <> "" Then
            VSFGN1.AddItem VSFGPeds.TextMatrix(i, 0) & vbTab & _
                VSFGPeds.TextMatrix(i, 2) & vbTab & _
                VSFGPeds.TextMatrix(i, 3) & vbTab & _
                VSFGPeds.TextMatrix(i, 4) & vbTab & _
                VSFGPeds.TextMatrix(i, 5) & vbTab & _
                VSFGPeds.TextMatrix(i, 6) & vbTab & _
                VSFGPeds.TextMatrix(i, 7) & vbTab & _
                VSFGPeds.TextMatrix(i, 8) & vbTab & _
                VSFGPeds.TextMatrix(i, 9) & vbTab & _
                VSFGPeds.TextMatrix(i, 10) & vbTab & _
                VSFGPeds.TextMatrix(i, 11)
            If i + 1 < VSFGPeds.Rows Then
                booEnviarN1 = True
            End If
            If booEnviarN1 = False Then
                If Trim(VSFGPeds.TextMatrix(15, i)) <> Trim(VSFGPeds.TextMatrix(15, i + 1)) Then
                    booEnviarN1 = True
                End If
            End If
            If booEnviarN1 = True Then
                VSFGN1.SaveGrid "PedidosPendientesN1.xls", flexFileExcel, flexXLSaveFixedCells

                EnviarMail NombreComercial & " Ajustes", CorreoServicioAlCliente, VSFGPeds.TextMatrix(4, i), Trim(VSFGPeds.TextMatrix(15, i)), "", "Reporte de pedidos pendientes de pago ", _
                        "Estimad@" & vbNewLine & _
                        VSFGPeds.TextMatrix(4, i) & vbNewLine & _
                        "Adjunto encontrará el listado de pedidos pendientes de pago." & vbNewLine & _
                        "Recuerde realizar el pago en el banco a través  del código identificado." & vbNewLine & _
                        "Saludos Cordiales" & vbNewLine & _
                        "Departamento de Ajustes" & vbNewLine & _
                        NombreComercial, "PedidosPendientesN1.xls"
                Kill "PedidosPendientesN1.xls"
                VSFGN1.Rows = 1
                VSFGN1.Clear flexClearScrollable
            End If
        End If

    Next i
    
End Sub

Private Sub cmdMas_Click()
    If cmdMas.Caption = "+" Then
        txtPedidos.Height = txtPedidos.Height * 10
        cmdMas.Caption = "-"
        txtPedidos.Locked = False
    Else
        txtPedidos.Height = txtPedidos.Height / 10
        cmdMas.Caption = "+"
        txtPedidos.Locked = True
        CargaGrid "M"
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim rpMov1 As New frmReporte
    rpMov1.strNumero = "10020000352"
    rpMov1.strReporte = "rptTckAjuste"
    rpMov1.Show
    Dim rpMov2 As New frmReporte
    rpMov2.strNumero = "10020000352"
    rpMov2.strReporte = "rptAjuste"
    rpMov2.Show
End Sub

Private Sub Command2_Click()
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT per_codigo,egreso.egr_codigo " & _
             " FROM no inner join egreso ON no.emp_codigo=egreso.emp_codigo " & _
             " and no.tip_egr_codigo=egreso.tip_egr_codigo " & _
             " and no.egr_codigo=egreso.egr_codigo "
    clsAux.Ejecutar strSql
    While Not clsAux.adorec_Def.EOF
        cmbCliente.BoundText = clsAux.adorec_Def("per_codigo")
        cmbCliente_Validate False
        cmbFactura.BoundText = clsAux.adorec_Def("egr_codigo")
        cmbFactura_Validate False
        cmdAceptar_Click
        clsAux.adorec_Def.MoveNext
    Wend
End Sub

Private Sub dtpCaduca_Change()
    txtCaduca.Text = Format(dtpCaduca.Value, "mm\/yyyy")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    strSql = ""
    Set clsCon_Def = Nothing
    Set clsCon_Prd = Nothing
    Set clsCon_Prd2 = Nothing
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    ValDias = 30
    booTodaFactura = False
    Set uctrVSFG1.VSFGControl = VSFGPeds
    uctrVSFG1.Inicializar False, False, False
    'Me.Label4.BackColor
    'uctrVSFG1.BackColor = Label4.BackColor
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd.Inicializar AdoConn, AdoConnMaster
    clsCon_Prd2.Inicializar AdoConn, AdoConnMaster
    
    Dim au As Integer
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " WHERE cod_iva_enuso=1 "
     clsCon_Def.Ejecutar (strSql)
     au = clsCon_Def.adorec_Def("cod_iva_codigo")
    TxtIva.Tag = clsCon_Def.adorec_Def("cod_iva_porcentaje")
    strSql = " SELECT cod_iva_codigo, cod_iva_porcentaje" & _
                 " FROM codigo_iva " & _
                 " ORDER BY cod_iva_porcentaje"
     clsCon_Def.Ejecutar (strSql)
     Set dcmbIVA.RowSource = clsCon_Def.adorec_Def.DataSource
     dcmbIVA.ListField = "cod_iva_porcentaje"
     dcmbIVA.BoundColumn = "cod_iva_codigo"
     dcmbIVA.BoundText = au
    
    dtpFecha.Value = HoyDia
    dtpFecha.Enabled = False
    
    cargarTipoPedido
    
    
    'Obtiene los tipos de formas de pago de una empresa y las muestra en un combo
    strSql = " SELECT for_pag_codigo, for_pag_nombre,for_pag_tiempo,for_pag_periodo " & _
             " FROM forma_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY for_pag_nombre "
    clsCon_Def.Ejecutar (strSql)
    Set CmbFpago.RowSource = clsCon_Def.adorec_Def.DataSource
    CmbFpago.ListField = "for_pag_nombre"
    CmbFpago.BoundColumn = "for_pag_codigo"
    
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
    
    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombC, COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as nombV, " & _
             " cat_p_nombre, lis_pre_codigo, per_codigo, COALESCE(vendedor.ven_codigo,'') as ven_codigo,per_ruc,per_direccion, " & _
             " COALESCE(CONCAT(per_telf,'/',per_fax),'') as per_tf,per_observacion,cat_p_dcto,per_dcto,per_credito,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,per_codigo_ref,per_codigo_ref2 " & _
             " FROM (persona LEFT JOIN vendedor ON (vendedor.ven_codigo = persona.ven_codigo) " & _
             " AND (vendedor.emp_codigo = persona.emp_codigo)) INNER JOIN categoria_p " & _
             " ON (persona.cat_p_tipo = categoria_p.cat_p_tipo) AND (persona.cat_p_codigo = categoria_p.cat_p_codigo) " & _
             " AND (persona.emp_codigo = categoria_p.emp_codigo) " & _
             " Where persona.emp_codigo='" & strEmpresa & "' And categoria_p.cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
             " AND persona.per_inactivo=0 " & _
             " ORDER BY nombC "
    clsCon_Def.Ejecutar (strSql)
    'Coloca los datos del primer cliente de la lista
    Set cmbCliente.RowSource = clsCon_Def.adorec_Def.DataSource
    If Not clsCon_Def.adorec_Def.EOF Then
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    Else
        cmbCliente = "No hay clientes en la empresa: " & strEmpresa
    End If
    
    'PonerBotones
End Sub

Private Sub vsfgDetalleIng_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 5 Or Col = 8 Then Cancel = True
End Sub

Private Sub VsfgDetalleIng_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = vsfgDetalleIng.MouseRow
    c = vsfgDetalleIng.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    If (c <> 0 Or r = (vsfgDetalleIng.Rows - 1)) Then Exit Sub
     
    ' make sure the click was on a cell with a button
    If vsfgDetalleIng.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsfgDetalleIng.Cell(flexcpLeft, r, c) + vsfgDetalleIng.Cell(flexcpWidth, r, c) - x
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnDn
    Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
    Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
    Título = "SisAdmi - Pedido a Bodega"   ' Define el título.
    respuesta = MsgBox(Mensaje, Estilo, Título)
        
    'Recorro el FlexGrid para poner números a las filas
        
    If respuesta = vbYes Then
         Dim i As Integer
         vsfgDetalleIng.RemoveItem (r)
         PonerBotones
         CalculaTotal
    Else
        vsfgDetalleIng.Cell(flexcpPicture, r, c) = imgBtnUp
    End If
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub


Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (vsfgDetalleIng.Rows - 1)
        vsfgDetalleIng.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            vsfgDetalleIng.Cell(flexcpPicture, i, 0) = imgBtnUp
            vsfgDetalleIng.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub




Private Sub cargarTipoPedido()
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    Set cmbNegocio2.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio2.ListField = "tip_ped_nombre"
    cmbNegocio2.BoundColumn = "tip_ped_codigo"
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsAux.adorec_Def(0)
        cmbNegocio2.BoundText = clsAux.adorec_Def(0)
    End If
End Sub


Private Sub CargaProductos()
    Dim i As Long
    vsfgDetalle.Clear 1
    vsfgDetalle.Rows = 2
    txtCantidad.Text = 0
    If cmbFactura.MatchedWithList = True Then
        'Consulto los productos de la factura
        If Left(cmbFactura.Text, 1) = "R" Then
            strSql = " SELECT dep_codigo,det_factura_ryb.prd_codigo,prd_nombre,det_egr_cantidad,det_egr_precio,det_egr_dcto,det_egr_costo,prd_iva " & _
                     " FROM det_factura_ryb INNER JOIN producto " & _
                     " ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
                     " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
                     " WHERE det_factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_factura_ryb.tip_egr_codigo = 'FAC' " & _
                     " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
                     " ORDER BY det_factura_ryb.prd_codigo "
        Else
            strSql = " SELECT det_egreso.dep_codigo,det_egreso.prd_codigo,prd_nombre,COALESCE(det_con_mer_cantidad,det_egr_cantidad) as det_egr_cantidad,det_egr_precio,det_egr_dcto/det_egr_cantidad*COALESCE(det_con_mer_cantidad,det_egr_cantidad) as det_egr_dcto,det_egr_costo,prd_iva,COALESCE(CAST(con_mer_codigo as varchar),'') as con_mer_codigo " & _
                     " FROM det_egreso INNER JOIN producto " & _
                     " ON det_egreso.emp_codigo=producto.emp_codigo " & _
                     " AND det_egreso.prd_codigo=producto.prd_codigo"
            strSql = strSql & " LEFT JOIN (SELECT det_contenedor_mercaderia.emp_codigo,tip_mov_codigo,mov_codigo,dep_codigo,det_contenedor_mercaderia.prd_codigo,contenedor_mercaderia.con_mer_codigo,det_con_mer_cantidad" & _
                     " FROM det_contenedor_mercaderia INNER JOIN contenedor_mercaderia ON det_contenedor_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                     " AND det_contenedor_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                     " WHERE det_contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_mov_codigo='FAC' " & _
                     " AND det_con_mer_cantidad!=0 " & _
                     " AND mov_codigo ='" & cmbFactura.BoundText & "') ubi ON det_egreso.emp_codigo=ubi.emp_codigo " & _
                     " AND det_egreso.tip_egr_codigo=ubi.tip_mov_codigo " & _
                     " AND det_egreso.egr_codigo=ubi.mov_codigo " & _
                     " AND det_egreso.dep_codigo=ubi.dep_codigo " & _
                     " AND det_egreso.prd_codigo=ubi.prd_codigo " & _
                     " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_egreso.tip_egr_codigo = 'FAC' " & _
                     " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                     " ORDER BY det_egreso.prd_codigo "
        End If
        clsCon_Prd.Ejecutar strSql
        i = 1
        While Not clsCon_Prd.adorec_Def.EOF
            vsfgDetalle.TextMatrix(i, 0) = i
            vsfgDetalle.TextMatrix(i, 1) = clsCon_Prd.adorec_Def("dep_codigo")
            vsfgDetalle.TextMatrix(i, 2) = clsCon_Prd.adorec_Def("prd_codigo")
            vsfgDetalle.TextMatrix(i, 3) = clsCon_Prd.adorec_Def("prd_nombre")
            vsfgDetalle.TextMatrix(i, 4) = clsCon_Prd.adorec_Def("det_egr_cantidad")
            txtCantidad.Text = txtCantidad.Text + clsCon_Prd.adorec_Def("det_egr_cantidad")
            vsfgDetalle.TextMatrix(i, 5) = clsCon_Prd.adorec_Def("det_egr_precio")
            vsfgDetalle.TextMatrix(i, 6) = clsCon_Prd.adorec_Def("det_egr_dcto")
            vsfgDetalle.TextMatrix(i, 7) = FormatoD2(FormatoD2(clsCon_Prd.adorec_Def("det_egr_cantidad") * clsCon_Prd.adorec_Def("det_egr_precio")) - clsCon_Prd.adorec_Def("det_egr_dcto"))
            vsfgDetalle.TextMatrix(i, 8) = clsCon_Prd.adorec_Def("det_egr_costo")
            vsfgDetalle.TextMatrix(i, 9) = clsCon_Prd.adorec_Def("prd_iva")
            vsfgDetalle.TextMatrix(i, 10) = clsCon_Prd.adorec_Def("con_mer_codigo")
            i = i + 1
            vsfgDetalle.AddItem ""
            clsCon_Prd.adorec_Def.MoveNext
        Wend
        VSFGReca.Clear 1
        VSFGReca.Rows = 2
        
        If Not Left(cmbFactura.Text, 1) = "R" Then
            strSql = " SELECT det_egreso_c.oca_codigo,oca_nombre,det_egr_c_precio " & _
                     " FROM det_egreso_c INNER JOIN ocargos " & _
                     " ON det_egreso_c.emp_codigo=ocargos.emp_codigo " & _
                     " AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                     " WHERE det_egreso_c.emp_codigo = '" & strEmpresa & "' " & _
                     " AND det_egreso_c.tip_egr_codigo = 'FAC' " & _
                     " AND det_egreso_c.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                     " ORDER BY det_egreso_c.oca_codigo "
            
            clsCon_Prd.Ejecutar strSql
            i = 1
            While Not clsCon_Prd.adorec_Def.EOF
                VSFGReca.TextMatrix(i, 0) = i
                VSFGReca.TextMatrix(i, 1) = clsCon_Prd.adorec_Def("oca_codigo")
                VSFGReca.TextMatrix(i, 2) = clsCon_Prd.adorec_Def("oca_nombre")
                VSFGReca.TextMatrix(i, 3) = clsCon_Prd.adorec_Def("det_egr_c_precio")
                i = i + 1
                VSFGReca.AddItem ""
                clsCon_Prd.adorec_Def.MoveNext
            Wend
        End If
        
        If Left(cmbFactura.Text, 1) = "R" Then
            strSql = " SELECT egr_subtotal,egr_subtotal_o,egr_dcto,egr_impuesto,egr_total,2 as cod_iva_codigo " & _
                     " FROM factura_ryb  " & _
                     " WHERE factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
                     " AND factura_ryb.tip_egr_codigo = 'FAC' " & _
                     " AND CONCAT('R',factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' "
        Else
            strSql = " SELECT egr_subtotal,egr_subtotal_o,egr_dcto,egr_impuesto,egr_total,cod_iva_codigo " & _
                     " FROM egreso  " & _
                     " WHERE egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND egreso.tip_egr_codigo = 'FAC' " & _
                     " AND egreso.egr_codigo = '" & cmbFactura.BoundText & "' "
        End If
        clsCon_Prd.Ejecutar strSql
        TxtSubTotal.Text = clsCon_Prd.adorec_Def("egr_subtotal")
        TxtRecargo.Text = clsCon_Prd.adorec_Def("egr_subtotal_o")
        TxtDesc.Text = clsCon_Prd.adorec_Def("egr_dcto")
        TxtIva.Text = clsCon_Prd.adorec_Def("egr_impuesto")
        dcmbIVA.BoundText = clsCon_Prd.adorec_Def("cod_iva_codigo")
        TxtTotal.Text = clsCon_Prd.adorec_Def("egr_total")
        
    End If
    

End Sub

Private Sub vsfgDetalleIng_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        If vsfgDetalleIng.TextMatrix(Row, Col) <> "" And Not IsNumeric(vsfgDetalleIng.TextMatrix(Row, Col)) Then
            MsgBox "Ingrese valores numéricos en Cantidad", vbInformation, "Detalle"
            vsfgDetalleIng.TextMatrix(Row, Col) = 0
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldCol = 4 Then
        If vsfgDetalleIng.TextMatrix(OldRow, 4) > vsfgDetalleIng.TextMatrix(OldRow, 9) Then
            MsgBox "No puede pedir el cambio de mas de " & vsfgDetalleIng.TextMatrix(OldRow, 9) & " prendas", vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(OldRow, 4) = 0
            Cancel = True
        End If
    End If
End Sub

Private Sub vsfgDetalleIng_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Or Col = 3 Then
        If vsfgDetalleIng.TextMatrix(Row, 1) = "" Then
            MsgBox "Seleccione primero un motivo", vbInformation, "Motivos"
            vsfgDetalleIng.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
    End If
    If Col = 2 Then
        vsfgDetalleIng.TextMatrix(Row, 3) = vsfgDetalleIng.TextMatrix(Row, 2)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_precio")
        vsfgDetalleIng.TextMatrix(Row, 9) = ConsutaCantidadMaxima(vsfgDetalleIng.TextMatrix(Row, 2))
    ElseIf Col = 3 Then
        vsfgDetalleIng.TextMatrix(Row, 2) = vsfgDetalleIng.TextMatrix(Row, 3)
        clsCon_Prd.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 2) & "'"
        vsfgDetalleIng.TextMatrix(Row, 4) = 0
        vsfgDetalleIng.TextMatrix(Row, 5) = clsCon_Prd.adorec_Def("prd_precio")
        vsfgDetalleIng.TextMatrix(Row, 9) = ConsutaCantidadMaxima(vsfgDetalleIng.TextMatrix(Row, 2))
    ElseIf Col = 6 Then
        vsfgDetalleIng.TextMatrix(Row, 7) = vsfgDetalleIng.TextMatrix(Row, 6)
        clsCon_Prd2.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 6) & "'"
        vsfgDetalleIng.TextMatrix(Row, 8) = clsCon_Prd2.adorec_Def("prd_precio")
    ElseIf Col = 7 Then
        vsfgDetalleIng.TextMatrix(Row, 6) = vsfgDetalleIng.TextMatrix(Row, 7)
        clsCon_Prd2.Filtrar "prd_codigo='" & vsfgDetalleIng.TextMatrix(Row, 7) & "'"
        vsfgDetalleIng.TextMatrix(Row, 8) = clsCon_Prd2.adorec_Def("prd_precio")
    End If
    If vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 1) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 2) <> "" And vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 3) <> "" And Val(vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 4)) <> 0 Then
        vsfgDetalleIng.AddItem ""
        vsfgDetalleIng.TextMatrix(vsfgDetalleIng.Rows - 1, 0) = vsfgDetalleIng.Rows - 1
        vsfgDetalleIng.Cell(flexcpPicture, vsfgDetalleIng.Rows - 1, 0) = imgBtnUp
        vsfgDetalleIng.Cell(flexcpPictureAlignment, vsfgDetalleIng.Rows - 1, 0) = flexAlignRightCenter
    End If
    CalculaTotal
End Sub

Private Function ConsutaCantidadMaxima(strProducto As String) As Long
    
    If Left(cmbFactura.Text, 1) = "R" Then
        strSql = " SELECT MAX(det_egr_cantidad) as cantidad " & _
                 " FROM det_factura_ryb INNER JOIN producto ON det_factura_ryb.emp_codigo=producto.emp_codigo " & _
                 " AND det_factura_ryb.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE det_factura_ryb.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_factura_ryb.tip_egr_codigo = 'FAC' " & _
                 " AND det_factura_ryb.prd_codigo='" & strProducto & "' " & _
                 " AND CONCAT('R',det_factura_ryb.egr_codigo) = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_factura_ryb.emp_codigo,det_factura_ryb.egr_codigo,det_factura_ryb.prd_codigo " & _
                 " ORDER BY cantidad DESC "
    Else
        strSql = " SELECT MAX(det_egr_cantidad) as cantidad " & _
                 " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
                 " AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE det_egreso.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_egreso.tip_egr_codigo = 'FAC' " & _
                 " AND det_egreso.prd_codigo='" & strProducto & "' " & _
                 " AND det_egreso.egr_codigo = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_egreso.emp_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo "
        strSql = strSql & " UNION " & _
                 " SELECT MAX(det_cam_cantidad) " & _
                 " FROM cambio INNER JOIN det_cambio ON cambio.emp_codigo=det_cambio.emp_codigo " & _
                 " AND cambio.cam_codigo=det_cambio.cam_codigo " & _
                 " INNER JOIN producto ON det_cambio.emp_codigo=producto.emp_codigo " & _
                 " AND det_cambio.prd_codigo_ped=producto.prd_codigo " & _
                 " INNER JOIN lista_precio_p ON producto.emp_codigo=lista_precio_p.emp_codigo " & _
                 " AND producto.prd_codigo=lista_precio_p.prd_codigo AND lista_precio_p.lis_pre_codigo='" & cmbCliente.Tag & "' " & _
                 " WHERE cambio.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_cambio.tip_ing_codigo = 'ICA' " & _
                 " AND det_cambio.prd_codigo_ped='" & strProducto & "' " & _
                 " AND cambio.cam_factura = '" & cmbFactura.BoundText & "' " & _
                 " AND prd_baja=0 " & _
                 " GROUP BY det_cambio.emp_codigo,det_cambio.cam_codigo,det_cambio.prd_codigo_ped " & _
                 " ORDER BY cantidad DESC "
    End If
    clsCon_Def.Ejecutar strSql
    ConsutaCantidadMaxima = FormatoD0(clsCon_Def.adorec_Def("cantidad"))
End Function

Private Sub CalculaTotal()
    Dim i As Long
    Dim totalIng As Double
    Dim CantIng As Double
    totalIng = 0
    CantIng = 0
    
    For i = 1 To vsfgDetalleIng.Rows - 1
        If FormatoD4(vsfgDetalleIng.TextMatrix(i, 5)) <> FormatoD4(vsfgDetalleIng.TextMatrix(i, 8)) And vsfgDetalleIng.TextMatrix(i, 6) <> "" Then
            MsgBox "no coinciden los precios del producto de la fila " & i, vbInformation, "Cambios"
            vsfgDetalleIng.TextMatrix(i, 6) = ""
            vsfgDetalleIng.TextMatrix(i, 7) = ""
            vsfgDetalleIng.TextMatrix(i, 8) = ""
        End If
        CantIng = CantIng + FormatoD4(vsfgDetalleIng.TextMatrix(i, 4))
    Next i
    
    txtCantIng.Text = FormatoD4(CantIng)
    
End Sub

Private Sub CargaGrid(strTipo As String)
    Dim LstPedidos As String
    
    VSFGPeds.Clear flexClearScrollable
    VSFGPeds = 1
    LstPedidos = ""
    'DESDE LA LISTA DE PEDIDOS (Manual)
    If strTipo = "M" Then
        LstPedidos = " ('" & Replace(Me.txtPedidos.Text, vbNewLine, "','") & "')"
    'DESDE pedidos no despachados (Automatico)
    ElseIf strTipo = "A" Then
        strSql = " SELECT pedido.emp_codigo, COUNT(pedido.ped_codigo) as n, GROUP_CONCAT(CAST(pedido.ped_codigo AS CHAR(16))) as ped " & _
                 " FROM pedido " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
                 " AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
                 " AND forma_pago.for_pag_revisiondespacho=1 " & _
                 " INNER JOIN egreso ON pedido.emp_codigo=egreso.emp_codigo " & _
                 " AND pedido.ped_egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo='FAC' AND egreso.egr_anulado=0 "
        If chkDesmantelaHoy.Value = 0 Then
            'MAÑANA DESMANTELA
            strSql = strSql & " AND (IIF(persona.for_pag_codigo_imp NOT IN ('EFE','CONT'),1=1, " & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP)) IN (7,6,5),LEFT(egr_fecha,10) = LEFT(DATEADD(d,-3,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP))=4,LEFT(egr_fecha,10) BETWEEN LEFT(DATEADD(d,-5,CURRENT_TIMESTAMP),10) AND LEFT(DATEADD(d,-3,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP)) IN (3,2),LEFT(egr_fecha,10) = LEFT(DATEADD(d,-5,CURRENT_TIMESTAMP),10)," & _
                     " LEFT(egr_fecha,10) = LEFT(DATEADD(d,-4,CURRENT_TIMESTAMP),10))))) "
            'PASADO MAÑANA DESMANTELA
            strSql = strSql & " OR IIF(persona.for_pag_codigo_imp NOT IN ('EFE','CONT'),1=1, " & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP)) IN (7,6,5,4),LEFT(egr_fecha,10) = LEFT(DATEADD(d,-2,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP))=3,LEFT(egr_fecha,10) BETWEEN LEFT(DATEADD(d,-4,CURRENT_TIMESTAMP),10) AND LEFT(DATEADD(d,-2,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP))=2,LEFT(egr_fecha,10) = LEFT(DATEADD(d,-4,CURRENT_TIMESTAMP),10)," & _
                     " LEFT(egr_fecha,10) = LEFT(DATEADD(d,-3,CURRENT_TIMESTAMP),10)))))) "
        Else
            'MAÑANA HOY
            strSql = strSql & " AND (IIF(persona.for_pag_codigo_imp NOT IN ('EFE','CONT'),1=1, " & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP)) IN (7,6),LEFT(egr_fecha,10) = LEFT(DATEADD(d,-4,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP))=5,LEFT(egr_fecha,10) BETWEEN LEFT(DATEADD(d,-6,CURRENT_TIMESTAMP),10) AND LEFT(DATEADD(d,-4,CURRENT_TIMESTAMP),10)," & _
                     " IIF(DATEPART(dw,DATEADD(d,1,CURRENT_TIMESTAMP)) IN (4,3,2),LEFT(egr_fecha,10) = LEFT(DATEADD(d,-6,CURRENT_TIMESTAMP),10)," & _
                     " LEFT(egr_fecha,10) = LEFT(DATEADD(d,-5,CURRENT_TIMESTAMP),10)))))) "
        
        End If
        strSql = strSql & " LEFT JOIN det_contenedor ON egreso.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND egreso.egr_codigo=det_contenedor.egr_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND persona.tip_ped_codigo = '" & cmbNegocio2.BoundText & "' " & _
                 " AND pedido.ped_estado in (2,4,8) " & _
                 " AND det_contenedor.emp_codigo is null " & _
                 " GROUP BY pedido.emp_codigo "
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
        LstPedidos = " (" & clsCon_Def.adorec_Def("ped") & ") "
        Else
            MsgBox "No hay pedidos para el reporte"
            LstPedidos = " ('') "
        End If
    End If
    
    strSql = " SELECT pedido.ped_codigo, LEFT(CURRENT_TIMESTAMP,10) as hoy,for_pag_nombre,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per," & _
             " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1,persona.per_direccion2,ciu_nombre, " & _
             " IIF(N9.per_codigo IS NOT NULL,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
             " IIF(N8.per_codigo IS NOT NULL,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
             " IIF(N7.per_codigo IS NOT NULL,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
             " IIF(N6.per_codigo IS NOT NULL,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
             " IIF(N5.per_codigo IS NOT NULL,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
             " IIF(N4.per_codigo IS NOT NULL,CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,''))," & _
             " IIF(N3.per_codigo IS NOT NULL,CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,''))," & _
             " IIF(N2.per_codigo IS NOT NULL,CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,''))," & _
             " IIF(N1.per_codigo IS NOT NULL,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')),''))))))))) as papa,egreso.egr_codigo," & _
             " egreso.egr_fecha,egreso.egr_total," & _
             " DATEADD(d,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha)) =1,4,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha))=2,3,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha)) in (3,4,5,6),5,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha))=7,4,0)))),egreso.egr_fecha) as fvencimiento,  " & _
             " DATEADD(d,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha))=1,5,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha))=2,4,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha)) in (3,4,5,6),6,IIF(DATEPART(dw,DATEADD(d,1,egreso.egr_fecha))=7,5,0)))),egreso.egr_fecha) as fdesmantelado,pedido.ped_usumod,  "
    strSql = strSql & " persona.per_email,N1.per_email," & _
             " IIF(N9.per_codigo IS NOT NULL,COALESCE(N9.per_email,'')," & _
             " IIF(N8.per_codigo IS NOT NULL,COALESCE(N8.per_email,'')," & _
             " IIF(N7.per_codigo IS NOT NULL,COALESCE(N7.per_email,'')," & _
             " IIF(N6.per_codigo IS NOT NULL,COALESCE(N6.per_email,'')," & _
             " IIF(N5.per_codigo IS NOT NULL,COALESCE(N5.per_email,'')," & _
             " IIF(N4.per_codigo IS NOT NULL,COALESCE(N4.per_email,'')," & _
             " IIF(N3.per_codigo IS NOT NULL,COALESCE(N3.per_email,'')," & _
             " IIF(N2.per_codigo IS NOT NULL,COALESCE(N2.per_email,'')," & _
             " IIF(N1.per_codigo IS NOT NULL,COALESCE(N1.per_email,''),''))))))))) as emailpapa " & _
             " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
             " INNER JOIN persona ON pedido.emp_codigo = persona.emp_codigo AND pedido.per_codigo = persona.per_codigo INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
             " AND persona.tip_ped_codigo LIKE '" & cmbNegocio2.BoundText & "' " & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo "
    strSql = strSql & " INNER JOIN egreso ON pedido.emp_codigo=egreso.emp_codigo AND pedido.ped_egr_codigo=egreso.egr_codigo AND pedido.ped_tip_egr_codigo=egreso.tip_egr_codigo " & _
             " AND pedido.per_codigo=persona.per_codigo AND egreso.egr_anulado=0 "
    strSql = strSql & " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo  " & _
             " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo  AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1" & _
             " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo  AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1" & _
             " LEFT JOIN persona N3 ON persona.emp_codigo = N3.emp_codigo  AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1" & _
             " LEFT JOIN persona N4 ON persona.emp_codigo = N4.emp_codigo  AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1" & _
             " LEFT JOIN persona N5 ON persona.emp_codigo = N5.emp_codigo  AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1" & _
             " LEFT JOIN persona N6 ON persona.emp_codigo = N6.emp_codigo  AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1" & _
             " LEFT JOIN persona N7 ON persona.emp_codigo = N7.emp_codigo  AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1" & _
             " LEFT JOIN persona N8 ON persona.emp_codigo = N8.emp_codigo  AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1" & _
             " LEFT JOIN persona N9 ON persona.emp_codigo = N9.emp_codigo  AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1" & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C'" & _
             " AND pedido.ped_codigo in " & LstPedidos & " " & _
             " ORDER BY persona.tip_ped_codigo,fdesmantelado,nn1,papa,pedido.ped_codigo "
    clsCon_Def.Ejecutar (strSql)
    Set VSFGPeds.DataSource = clsCon_Def.adorec_Def.DataSource
End Sub
