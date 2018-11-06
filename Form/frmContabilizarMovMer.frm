VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilizarMovMer 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilizar Movimientos de Mercadería"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "frmContabilizarMovMer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10830
   Begin VB.CommandButton cmdRecosteo 
      Caption         =   "Recostear"
      Height          =   375
      Left            =   9480
      TabIndex        =   32
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Mercadería"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   128
      TabIndex        =   16
      Top             =   0
      Width           =   10575
      Begin VB.Frame frmSucursal 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Sucursal"
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
         Left            =   7440
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   2895
         Begin MSDataListLib.DataCombo dcmbSucursal 
            Height          =   330
            Left            =   840
            TabIndex        =   6
            Top             =   360
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   435
            Width           =   615
         End
      End
      Begin VB.Frame frmCtaConta 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Importacion"
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
         Height          =   735
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
         Begin MSDataListLib.DataCombo dcmbCtaConta 
            Height          =   330
            Left            =   1320
            TabIndex        =   3
            Top             =   300
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.Contable"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.TextBox txtDcto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtSubTotal_s 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtRec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtTotalDebe 
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
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtTotalHaber 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   6120
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Rango"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpFechaI 
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
            Left            =   1320
            TabIndex        =   4
            Top             =   263
            Width           =   2055
            _ExtentX        =   3625
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66387971
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpFechaF 
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
            Left            =   1320
            TabIndex        =   5
            Top             =   623
            Width           =   2055
            _ExtentX        =   3625
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   66387971
            CurrentDate     =   37463
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Fin:"
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
            TabIndex        =   22
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inicio:"
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
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Asiento"
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
         Left            =   4080
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
         Begin NEED2.dtpFecha dtpFechaA 
            Height          =   285
            Left            =   720
            TabIndex        =   33
            Top             =   360
            Width           =   2055
            _extentx        =   3625
            _extenty        =   503
         End
         Begin VB.Label Label3 
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
            Left            =   90
            TabIndex        =   19
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Movimientos"
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
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4575
         Begin VB.OptionButton optImportacion 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Importaciones"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optCompraLocal 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Compras Locales"
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
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optFactura 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Facturas"
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
            Height          =   375
            Left            =   1920
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGdb 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   7695
         _cx             =   13573
         _cy             =   3201
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmContabilizarMovMer.frx":030A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFGAdquisicion 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   10335
         _cx             =   18230
         _cy             =   2990
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmContabilizarMovMer.frx":03B7
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4320
         TabIndex        =   24
         Top             =   3990
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4800
         TabIndex        =   23
         Top             =   6150
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3675
      TabIndex        =   8
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   6600
      Width           =   1575
   End
End
Attribute VB_Name = "frmContabilizarMovMer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para Contabilizar de Adquisiciones                                    #
'#  frmContabilizarAdq V1.0                                                     #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para ingresar el asiento adquisición                                #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  adquisicion: Esta tabla contiene los datos de la adquisicion                #
'#  aisento: tabla donde se almace los asientoa                                 #
'#  tipo de aisento: donde se guardan los datos tipos de asiento                #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'


Private clsAdq As New clsConsulta
Private clsAsi As New clsConsulta
Private clsPro As New clsConsulta
Private clsAct As New clsConsulta
Private clsSum As New clsConsulta
Private clsMaxAsi As New clsConsulta
Private clsSql As New clsConsulta
Dim strSQL As String
Private cta_pagar As String
Private tip_egr_ctaconta As String
Private tip_egr_ctaconta2 As String
Private iva_compra As String
Private cta_pagar_n As String
Private tip_egr_ctaconta_n As String
Private tip_egr_ctaconta_n2 As String
Private iva_compra_n As String

Private cta_cobrar As String
Private tip_ing_ctaconta As String
Private tip_ing_ctaconta2 As String
Private tip_ing_ctaconta3 As String
Private iva_venta As String
Private cta_cobrar_n As String
Private tip_ing_ctaconta_n As String
Private tip_ing_ctaconta_n2 As String
Private tip_ing_ctaconta_n3 As String
Private iva_venta_n As String

Dim j As Integer
Dim ban As Variant

Private Sub cmdRecosteo_Click()
    frmRecosteo.Show
End Sub

Private Sub dcmbCtaConta_Change()
    Rango_Fecha
End Sub

Private Sub dcmbCtaConta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        frmSelecCtaConta.Tag = "UN"
        frmSelecCtaConta.Show
        Set frmSelecCtaConta.objEscribir = dcmbCtaConta
    End If
End Sub

Private Sub dcmbSucursal_LostFocus()
    Call Rango_Fecha
End Sub

Private Sub dtpFechaI_Validate(Cancel As Boolean)
    Rango_Fecha
End Sub

Private Sub dtpFechaF_Validate(Cancel As Boolean)
    Rango_Fecha
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsAdq = Nothing
    Set clsAsi = Nothing
    Set clsPro = Nothing
    Set clsAct = Nothing
    Set clsSum = Nothing
    Set clsMaxAsi = Nothing
    Set clsSql = Nothing
End Sub

Private Sub PonerNumeros(Optional conBot As Boolean = True)
    For i = 1 To (VSFGAdquisicion.Rows - 1)
        VSFGAdquisicion.TextMatrix(i, 0) = i
    Next i
End Sub
Private Sub DebeHaber()
    Dim count As Integer
    count = 0
    
    VSFGdb.Clear 1
    VSFGdb.Rows = 1
    For k = 1 To (VSFGAdquisicion.Rows - 1)
        If VSFGAdquisicion.TextMatrix(k, 1) = -1 Then
               
        strSQL = " SELECT par_texto,par_nombre,'0.00' as debe,adq_subtotal " & _
                 " FROM ((adquisicion " & _
                 " INNER JOIN  persona " & _
                 " ON adquisicion.per_codigo = persona.per_codigo " & _
                 " AND adquisicion.emp_codigo = persona.emp_codigo ) " & _
                 " INNER JOIN parametro " & _
                 " ON persona.emp_codigo = parametro.emp_codigo) " & _
                 " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                 " AND cat_p_tipo = 'P' " & _
                 "  AND par_codigo = 'CXP' " & _
                 " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                 " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "

        clsPro.Ejecutar strSQL
        j = clsPro.adorec_Def.RecordCount
            If (clsPro.adorec_Def.RecordCount > 0) Then
            clsPro.adorec_Def.MoveFirst
            'Set VSFGdb.DataSource = clsAct.adorec_Def.DataSource
             ' For i = 1 To j
                    VSFGdb.Rows = VSFGdb.Rows + 1
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsPro.adorec_Def("par_texto")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsPro.adorec_Def("par_nombre")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsPro.adorec_Def("debe")
                    VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsPro.adorec_Def("adq_subtotal")
            ' Next i
            End If
        
        strSQL = " SELECT tip_act_ctaconta,cta_nombre,act_fij_valor,'0.00'as ha " & _
                " FROM ((((adquisicion " & _
                " INNER JOIN  det_adquisicion_af " & _
                " ON adquisicion.adq_codigo = det_adquisicion_af.adq_codigo " & _
                " AND adquisicion.emp_codigo = det_adquisicion_af.emp_codigo ) " & _
                " INNER JOIN activo_fijo " & _
                " ON det_adquisicion_af.emp_codigo = activo_fijo.emp_codigo " & _
                " AND activo_fijo.act_fij_codigo = det_adquisicion_af.act_fij_codigo ) " & _
                " INNER JOIN tipo_activo " & _
                " ON tipo_activo.tip_act_codigo = activo_fijo.tip_act_codigo " & _
                " AND tipo_activo.emp_codigo = activo_fijo.emp_codigo ) " & _
                " INNER JOIN ctaconta " & _
                " ON tipo_activo.tip_act_ctaconta = ctaconta.cta_codigo " & _
                " AND tipo_activo.emp_codigo = ctaconta.emp_codigo ) " & _
                " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "
        clsAct.Ejecutar strSQL

        If (clsAct.adorec_Def.RecordCount > 0) Then
            
            For m = 1 To clsAct.adorec_Def.RecordCount
                VSFGdb.Rows = VSFGdb.Rows + 1
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsAct.adorec_Def("tip_act_ctaconta")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsAct.adorec_Def("cta_nombre")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsAct.adorec_Def("act_fij_valor")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsAct.adorec_Def("ha")
                clsAct.adorec_Def.MoveNext
            Next m
        End If

       strSQL = " SELECT tip_sum_ctaconta,cta_nombre,(det_adq_su_cantidad*sum_ultimo_precio) as deb,'0.00'as hab " & _
                " FROM ((((adquisicion " & _
                " INNER JOIN  det_adquisicion_su " & _
                " ON adquisicion.adq_codigo = det_adquisicion_su.adq_codigo " & _
                " AND adquisicion.emp_codigo = det_adquisicion_su.emp_codigo ) " & _
                " INNER JOIN suministro " & _
                " ON det_adquisicion_su.emp_codigo = suministro.emp_codigo " & _
                " AND suministro.sum_codigo = det_adquisicion_su.sum_codigo ) " & _
                " INNER JOIN tipo_suministro " & _
                " ON tipo_suministro.tip_sum_codigo = suministro.tip_sum_codigo " & _
                " AND tipo_suministro.emp_codigo = suministro.emp_codigo ) " & _
                " INNER JOIN ctaconta " & _
                " ON tipo_suministro.tip_sum_ctaconta = ctaconta.cta_codigo " & _
                " AND tipo_suministro.emp_codigo = ctaconta.emp_codigo ) " & _
                " WHERE adquisicion.emp_codigo = '" & strEmpresa & "' " & _
                " AND adquisicion.adq_codigo = '" & VSFGAdquisicion.TextMatrix(k, 2) & "' " & _
                " AND adquisicion.adq_numdoc = '" & VSFGAdquisicion.TextMatrix(k, 3) & "' "
        clsSum.Ejecutar strSQL

        If (clsSum.adorec_Def.RecordCount > 0) Then
            For n = 1 To clsSum.adorec_Def.RecordCount
                VSFGdb.Rows = VSFGdb.Rows + 1
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 1) = clsSum.adorec_Def("tip_sum_ctaconta")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 2) = clsSum.adorec_Def("cta_nombre")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 3) = clsSum.adorec_Def("deb")
                VSFGdb.TextMatrix(VSFGdb.Rows - 1, 4) = clsSum.adorec_Def("hab")
                clsSum.adorec_Def.MoveNext
            Next n
        End If
    End If
            If VSFGAdquisicion.TextMatrix(k, 1) = "-1" Then
                count = count + 1
            End If
   Next k
   'abilita el boton aceptar
        If count <= 0 Then
                cmdAceptar.Enabled = False
            Else
                cmdAceptar.Enabled = True
        End If
End Sub
Private Sub CalTotalDebeHaber()

   'Calcula totales
    Dim SumaDebe As Double
    Dim SumaHaber As Double
    'Calcula total debe
    For i = 1 To VSFGdb.Rows - 1
        SumaDebe = SumaDebe + Val(VSFGdb.TextMatrix(i, 3))
    Next i
    txtTotalDebe = Format(SumaDebe, "##0.00")
    'Calcula total haber
    For i = 1 To VSFGdb.Rows - 1
        SumaHaber = SumaHaber + Val(VSFGdb.TextMatrix(i, 4))
    Next i
    txtTotalHaber = Format(SumaHaber, "##0.00")

End Sub
Private Sub Rango_Fecha()
    If Me.optFactura.Value = True And (dcmbSucursal.MatchedWithList = False) Then
        Exit Sub
    End If
   'Ejectua el Selct con el rango de fechas deseado
    fi = Format(dtpFechaI.Value, "yyyy-mm-dd")
    ff = Format(dtpFechaF.Value, "yyyy-mm-dd")
    'If cmbDiaA.Tag <> "A" Then
        'Verifican si las fecha ingresadas son correctas
        If (IsDate(fi)) = False Then
            MsgBox "La fecha de Inicio no es correcta", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
            'cmbAñoI.SetFocus
            Exit Sub
        End If
        If (IsDate(ff)) = False Then
            MsgBox "La fecha de Fin no es correcta", vbExclamation, "SisAdmi - Contabilizar Adquisiciones "
            'cmbAñoF.SetFocus
            Exit Sub
        End If
'    Else
'        Exit Sub
'    End If

    If ff >= fi Then
    'llenar flexgrid
        If Me.optCompraLocal.Value = True Then
            strSQL = " SELECT '1'as sel,ing_factura,ing_codigo, ing_fecha, concat(per_apellido,' ',per_nombre)as persona,ing_subtotal as ing_subtotal_m,'0' as ing_subtotal_s,ing_subtotal_o,ing_dcto,ing_impuesto,ing_total,concat(substring(ing_observacion,1,10),'...') as obs " & _
                     " FROM ingreso  INNER JOIN  persona  ON ingreso.per_codigo = persona.per_codigo AND ingreso.emp_codigo = persona.emp_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & strEmpresa & "' AND ingreso.ing_anulado=0 " & _
                     " AND ing_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND tip_ing_codigo = 'COM' " & _
                     " AND (ing_numasiento='' OR ing_numasiento IS NULL)" & _
                     " UNION " & _
                     "  " & _
                     " SELECT '1'as sel,egr_factura,egr_codigo, egr_fecha, concat(per_apellido,' ',per_nombre)as persona,(egr_subtotal * -1) as ing_subtotal_m,'0' as ing_subtotal_s,(egr_subtotal_o * -1) as ing_subtotal_o,egr_dcto as ing_dcto,(egr_impuesto * -1) as ing_impuesto,(egr_total * -1) as ing_total,concat(substring(egr_observacion,1,10),'...') as obs " & _
                     " FROM egreso  INNER JOIN  persona  ON egreso.per_codigo = persona.per_codigo AND egreso.emp_codigo = persona.emp_codigo " & _
                     " WHERE egreso.emp_codigo = '" & strEmpresa & "' AND egreso.egr_anulado=0 " & _
                     " AND egr_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND tip_egr_codigo = 'DPV' " & _
                     " AND (egr_numasiento='' OR egr_numasiento IS NULL)" & _
                     " ORDER BY ing_fecha,ing_codigo "
        ElseIf Me.optFactura.Value = True Then
            strSQL = " SELECT '1'as sel,'' as a,egreso.egr_codigo, egr_fecha, concat(per_apellido,' ',per_nombre)as persona,(egr_subtotal - (sum(if(left(prd_codigo,3)!='PR-' AND left(prd_codigo,3)!='MDO',0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0))))) as egr_subtotal_m,(sum(if(left(prd_codigo,3)!='PR-' AND left(prd_codigo,3)!='MDO',0,COALESCE(det_egr_cantidad,0)*COALESCE(det_egr_precio,0)))) as egr_subtotal_s,egr_subtotal_o,egr_dcto,egr_impuesto,egr_total,concat(substring(egr_observacion,1,10),'...') as obs " & _
                     " FROM (egreso  INNER JOIN  persona  ON egreso.per_codigo = persona.per_codigo AND egreso.emp_codigo = persona.emp_codigo)INNER JOIN det_egreso ON egreso.egr_codigo=det_egreso.egr_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.emp_codigo=det_egreso.emp_codigo " & _
                     " WHERE egreso.emp_codigo = '" & strEmpresa & "' AND egreso.egr_anulado=0 " & _
                     " AND egr_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND egreso.tip_egr_codigo = 'FAC' " & _
                     " AND (egr_numasiento='' OR egr_numasiento IS NULL)" & _
                     " AND egr_observacion NOT LIKE 'FACTURA ANULA%' " & _
                     " AND egreso.egr_codigo LIKE '" & Format(dcmbSucursal.BoundText, "###0") & "%' " & _
                     " GROUP BY sel,a,egreso.egr_codigo, egr_fecha, persona,egr_subtotal,egr_subtotal_o,egr_impuesto,egr_total,obs,egr_subtotal" & _
                     " UNION " & _
                     " SELECT '1'as sel,ing_factura,ing_codigo, ing_fecha, concat(per_apellido,' ',per_nombre)as persona,(ing_subtotal * -1) as egr_subtotal_m,0 as egr_subtotal_s,(ing_subtotal_o * -1) as egr_subtotal_o,ing_dcto as egr_dcto,(ing_impuesto * -1) as egr_impuesto,(ing_total * -1) as egr_total,concat(substring(ing_observacion,1,10),'...') as obs " & _
                     " FROM ingreso  INNER JOIN  persona  ON ingreso.per_codigo = persona.per_codigo AND ingreso.emp_codigo = persona.emp_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & strEmpresa & "' AND ingreso.ing_anulado=0 " & _
                     " AND ing_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND tip_ing_codigo = 'DCL' " & _
                     " AND (ing_numasiento='' OR ing_numasiento IS NULL)" & _
                     " AND ingreso.ing_codigo LIKE '" & Format(dcmbSucursal.BoundText, "###0") & "%' " & _
                     " ORDER BY egr_fecha,egreso.egr_codigo "
        Else
            strSQL = " SELECT '1'as sel,ing_factura,ingreso.ing_codigo, ing_fecha, concat(per_apellido,' ',per_nombre)as persona,sum(COALESCE(det_ing_cantidad,0)*COALESCE(det_ing_precio,0)) as ing_subtotal_m,'0' as ing_subtotal_s,ing_subtotal_o,ing_dcto,ing_impuesto,ing_total,concat(substring(ing_observacion,1,10),'...') as obs " & _
                     " FROM ingreso INNER JOIN persona  ON ingreso.per_codigo = persona.per_codigo AND ingreso.emp_codigo = persona.emp_codigo " & _
                     " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo AND  ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & strEmpresa & "' AND ingreso.ing_anulado=0 " & _
                     " AND ing_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND ingreso.tip_ing_codigo = 'IIM' " & _
                     " AND (ing_numasiento='' OR ing_numasiento IS NULL) GROUP BY ing_factura,ingreso.ing_codigo, ing_fecha, concat(per_apellido,' ',per_nombre),ing_subtotal_o,ing_dcto,ing_impuesto,ing_total,concat(substring(ing_observacion,1,10),'...')" & _
                     " ORDER BY ing_fecha,ing_codigo "
        End If
        clsAdq.Ejecutar strSQL
        
        If Not (clsAdq.adorec_Def.EOF) Then
            TxtSubTotal.Text = 0
            txtSubTotal_s.Text = 0
            txtDcto.Text = 0
            TxtIva.Text = 0
            TxtTotal.Text = 0
            txtRec.Text = 0
            ban = 0
            Set VSFGAdquisicion.DataSource = clsAdq.adorec_Def.DataSource
            VSFGAdquisicion.ColDataType(1) = flexDTBoolean
            ban = 1
            Call PonerNumeros
            If dcmbCtaConta.Text <> "" Then
                Call Cal_Total
            End If
        Else
            'MsgBox "No hay Compras Ingresadas en este rango de fechas", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
            Call limpiarFxGD
            TxtSubTotal.Text = ""
            txtSubTotal_s.Text = ""
            txtDcto.Text = ""
            TxtIva.Text = ""
            TxtTotal.Text = ""
        End If
    Else
        MsgBox "La Fecha Fin es mayor que la Fecha Inicio, Verifiquela por Favor!", vbExclamation, "SisAdmi - Contabilizar Adquisiciones"
        Exit Sub
    End If
      
End Sub
Private Sub Cal_Total()
   'Calcula totales del grid de adquisicion
    Dim Subtotal As Double
    Dim SubTotal_s As Double
    Dim Dcto As Double
    Dim IVA As Double
    Dim rec As Double
    Dim Total As Double
    Dim ff As Date
    Dim fi As Date
    Subtotal = 0
    SubTotal_s = 0
    Dcto = 0
    IVA = 0
    Total = 0
    fi = Format(dtpFechaI.Value, "yyyy-mm-dd")
    ff = Format(dtpFechaF.Value, "yyyy-mm-dd")
    For i = 1 To VSFGAdquisicion.Rows - 1
        If Abs(VSFGAdquisicion.TextMatrix(i, 1)) = 1 Then
            VSFGAdquisicion.Select i, 1, i, 12
            VSFGAdquisicion.FillStyle = flexFillRepeat
            VSFGAdquisicion.CellBackColor = &HC0FFFF
            VSFGAdquisicion.Select i, 12
            Subtotal = Subtotal + (Val(Format(VSFGAdquisicion.TextMatrix(i, 6), "###0.00")))
            SubTotal_s = SubTotal_s + (Val(VSFGAdquisicion.TextMatrix(i, 7)))
            rec = rec + (Val(VSFGAdquisicion.TextMatrix(i, 8)))
            Dcto = Dcto + (Val(VSFGAdquisicion.TextMatrix(i, 9)))
            IVA = IVA + (Val(VSFGAdquisicion.TextMatrix(i, 10)))
            Total = Total + (Val(VSFGAdquisicion.TextMatrix(i, 11)))
         End If
    Next i
    TxtSubTotal.Text = Format(Subtotal, "##0.00")
    txtSubTotal_s.Text = Format(SubTotal_s, "##0.00")
    txtDcto.Text = Format(Dcto, "##0.00")
    TxtIva.Text = Format(Val(IVA), "##0.00")
    TxtTotal.Text = Format(Val(Total), "##0.00")
    txtRec.Text = Format(Val(rec), "##0.00")
    VSFGdb.Clear 1
    VSFGdb.Rows = 6
    If Me.optCompraLocal.Value = True Then
            strSQL = " SELECT oca_ctaconta,cta_nombre,sum(COALESCE(det_ing_c_cantidad,0)*COALESCE(det_ing_c_precio,0)) as val " & _
                     " FROM ingreso  INNER JOIN  det_ingreso_c ON ingreso.ing_codigo = det_ingreso_c.ing_codigo AND ingreso.emp_codigo = det_ingreso_c.emp_codigo AND ingreso.tip_ing_codigo = det_ingreso_c.tip_ing_codigo " & _
                     " INNER JOIN ocargos ON det_ingreso_c.oca_codigo= ocargos.oca_codigo AND det_ingreso_c.emp_codigo= ocargos.emp_codigo " & _
                     " INNER JOIN ctaconta ON ocargos.oca_ctaconta=ctaconta.cta_codigo AND ctaconta.emp_codigo= ocargos.emp_codigo " & _
                     " WHERE ingreso.emp_codigo = '" & strEmpresa & "' AND ingreso.ing_anulado=0 " & _
                     " AND ing_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND ingreso.tip_ing_codigo = 'COM' " & _
                     " AND (ing_numasiento='' OR ing_numasiento IS NULL)" & _
                     " GROUP BY oca_ctaconta,cta_nombre "
            clsSql.Ejecutar strSQL
    ElseIf Me.optFactura.Value = True Then
            strSQL = " SELECT oca_ctaconta,cta_nombre,sum(COALESCE(det_egr_c_cantidad,0)*COALESCE(det_egr_c_precio,0)) as val " & _
                     " FROM egreso  INNER JOIN  det_egreso_c ON egreso.egr_codigo = det_egreso_c.egr_codigo AND egreso.emp_codigo = det_egreso_c.emp_codigo AND egreso.tip_egr_codigo = det_egreso_c.tip_egr_codigo " & _
                     " INNER JOIN ocargos ON det_egreso_c.oca_codigo= ocargos.oca_codigo AND det_egreso_c.emp_codigo= ocargos.emp_codigo " & _
                     " INNER JOIN ctaconta ON ocargos.oca_ctaconta=ctaconta.cta_codigo AND ctaconta.emp_codigo= ocargos.emp_codigo " & _
                     " WHERE egreso.emp_codigo = '" & strEmpresa & "' AND egreso.egr_anulado=0 " & _
                     " AND egr_fecha BETWEEN '" & fi & "'AND '" & ff & "'AND egreso.tip_egr_codigo = 'FAC' " & _
                     " AND (egr_numasiento='' OR egr_numasiento IS NULL)" & _
                     " GROUP BY oca_ctaconta,cta_nombre "
            clsSql.Ejecutar strSQL
    End If
    If Me.optCompraLocal.Value = True Then
        VSFGdb.TextMatrix(1, 1) = cta_pagar
        VSFGdb.TextMatrix(1, 2) = cta_pagar_n
        VSFGdb.TextMatrix(1, 3) = 0
        VSFGdb.TextMatrix(1, 4) = TxtTotal.Text
        VSFGdb.TextMatrix(2, 1) = tip_ing_ctaconta
        VSFGdb.TextMatrix(2, 2) = tip_ing_ctaconta_n
        VSFGdb.TextMatrix(2, 3) = TxtSubTotal.Text
        VSFGdb.TextMatrix(2, 4) = 0
        VSFGdb.TextMatrix(3, 1) = tip_ing_ctaconta2
        VSFGdb.TextMatrix(3, 2) = tip_ing_ctaconta_n2
        VSFGdb.TextMatrix(3, 3) = txtSubTotal_s.Text
        VSFGdb.TextMatrix(3, 4) = 0
        VSFGdb.TextMatrix(4, 1) = iva_compra
        VSFGdb.TextMatrix(4, 2) = iva_compra_n
        VSFGdb.TextMatrix(4, 3) = TxtIva.Text
        VSFGdb.TextMatrix(4, 4) = 0
        
        If clsSql.adorec_Def.RecordCount > 0 Then
            While Not clsSql.adorec_Def.EOF
                VSFGdb.AddItem "" & vbTab & clsSql.adorec_Def("oca_ctaconta") & vbTab & clsSql.adorec_Def("cta_nombre") & _
                                vbTab & clsSql.adorec_Def("val") & vbTab & "0"
                clsSql.adorec_Def.MoveNext
            Wend
        End If
    ElseIf Me.optFactura.Value = True Then
        'Consulta para conocer la cuenta contable ventas de productos de la sucursal
        strSQL = " SELECT suc_ctaconta_ventas,cta_nombre " & _
                  " FROM sucursal INNER JOIN ctaconta ON sucursal.suc_ctaconta_ventas=ctaconta.cta_codigo AND sucursal.emp_codigo=ctaconta.emp_codigo " & _
                  " WHERE sucursal.suc_codigo='" & dcmbSucursal.BoundText & "' AND sucursal.emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSQL
        If (clsSql.adorec_Def.RecordCount > 0) Then
            tip_egr_ctaconta = clsSql.adorec_Def("suc_ctaconta_ventas")
            tip_egr_ctaconta_n = clsSql.adorec_Def("cta_nombre")
        End If
        'Consulta para conocer la cuenta contable ventas de servicios de la sucursal
        strSQL = " SELECT suc_ctaconta_servicios,cta_nombre " & _
                  " FROM sucursal INNER JOIN ctaconta ON sucursal.suc_ctaconta_servicios=ctaconta.cta_codigo AND sucursal.emp_codigo=ctaconta.emp_codigo " & _
                  " WHERE sucursal.suc_codigo='" & dcmbSucursal.BoundText & "' AND sucursal.emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSQL
        If (clsSql.adorec_Def.RecordCount > 0) Then
            tip_egr_ctaconta2 = clsSql.adorec_Def("suc_ctaconta_servicios")
            tip_egr_ctaconta_n2 = clsSql.adorec_Def("cta_nombre")
        End If
        VSFGdb.TextMatrix(1, 1) = cta_cobrar
        VSFGdb.TextMatrix(1, 2) = cta_cobrar_n
        VSFGdb.TextMatrix(1, 4) = 0
        VSFGdb.TextMatrix(1, 3) = TxtTotal.Text
        VSFGdb.TextMatrix(5, 1) = dcmbCtaConta.Text
        VSFGdb.TextMatrix(5, 2) = "DESCUENTO EN VENTAS"
        VSFGdb.TextMatrix(5, 4) = 0
        VSFGdb.TextMatrix(5, 3) = txtDcto.Text
        VSFGdb.TextMatrix(2, 1) = tip_egr_ctaconta
        VSFGdb.TextMatrix(2, 2) = tip_egr_ctaconta_n
        VSFGdb.TextMatrix(2, 4) = TxtSubTotal.Text
        VSFGdb.TextMatrix(2, 3) = 0
        VSFGdb.TextMatrix(3, 1) = tip_egr_ctaconta2
        VSFGdb.TextMatrix(3, 2) = tip_egr_ctaconta_n2
        VSFGdb.TextMatrix(3, 4) = txtSubTotal_s.Text
        VSFGdb.TextMatrix(3, 3) = 0
        VSFGdb.TextMatrix(4, 1) = iva_venta
        VSFGdb.TextMatrix(4, 2) = iva_venta_n
        VSFGdb.TextMatrix(4, 4) = TxtIva.Text
        VSFGdb.TextMatrix(4, 3) = 0
    Else
        VSFGdb.TextMatrix(2, 1) = dcmbCtaConta.Text
        VSFGdb.TextMatrix(2, 2) = "IMPORTACION DEL PROVEEDOR"
        VSFGdb.TextMatrix(2, 3) = 0
        VSFGdb.TextMatrix(2, 4) = TxtSubTotal.Text
        VSFGdb.TextMatrix(1, 1) = tip_ing_ctaconta3
        VSFGdb.TextMatrix(1, 2) = tip_ing_ctaconta_n3
        VSFGdb.TextMatrix(1, 3) = TxtSubTotal.Text
        VSFGdb.TextMatrix(1, 4) = 0
    End If
    'DebeHaber
    CalTotalDebeHaber
    cmdAceptar.Enabled = True
End Sub
Private Sub Limpiar()
    VSFGAdquisicion.Clear 1
    VSFGAdquisicion.Rows = 2
    VSFGdb.Clear 1
    VSFGdb.Rows = 2
    TxtSubTotal = 0
    TxtIva = 0
    TxtTotal = 0
    txtTotalHaber = 0
    txtTotalDebe = 0
End Sub

Private Sub cmdAceptar_Click()
    Dim fa As String
    Dim Des As String
    'Comprueba que todos los datos esten ingresados
    fa = Format(dtpFechaA.Value, "yyyy-mm-dd")
    If (IsDate(fa) = False) Then
        MsgBox "La fecha de Asiento no es válida", vbInformation, "SisAdmi-Comprobante Aquisición"
        Exit Sub
    End If
    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe.Text <> txtTotalHaber.Text Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "SisAdmi-Comprobante Aquisición"
    End If
   'Compacta la matriz
    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el grid debe haber
    a = VSFGdb.Rows - 1
    For i = 1 To a
        For j = i + 1 To a
            If VSFGdb.TextMatrix(i, 1) = VSFGdb.TextMatrix(j, 1) Then
                VSFGdb.TextMatrix(i, 3) = Val(VSFGdb.TextMatrix(i, 3)) + Val(VSFGdb.TextMatrix(j, 3))
                VSFGdb.TextMatrix(i, 4) = Val(VSFGdb.TextMatrix(i, 4)) + Val(VSFGdb.TextMatrix(j, 4))
                VSFGdb.RemoveItem j
                a = a - 1
                j = j - 1
            End If
            If j >= a Then
                Exit For
            End If
        Next j
    Next i
    'Verificar que todos los datos se han llenado para ingresar en la base de datos
    If VSFGdb.TextMatrix(1, 1) = "" Then
        MsgBox "No estan ingresados todos los datos", vbInformation, "SisAdmi-Comprobante Aquisición"
        Exit Sub
    Else
        Dim clsAsi As New clsContable
        clsAsi.Inicializar AdoConn, AdoConnMaster
        
        If Me.optCompraLocal.Value = True Then
            Des = "ASIENTO DE COMPRAS LOCALES "
        ElseIf Me.optFactura.Value = True Then
            Des = "ASIENTO DE FACTURAS "
        Else
            Des = "ASIENTO DE IMPORTACION "
        End If
        clsAsi.NuevoAsiento "D", fa, 0, 0, FormatoD2(txtTotalDebe.Text), Des, True
        strMaximo = clsAsi.NumAsiento
        'Actualiza el campo adq_asentada poniendo "1"
        For i = 1 To (VSFGAdquisicion.Rows - 1)
            If Abs(VSFGAdquisicion.TextMatrix(i, 1)) = 1 Then
                If Me.optCompraLocal.Value = True Then
                    If VSFGAdquisicion.TextMatrix(i, 11) > 0 Then
                        strSQL = " UPDATE ingreso " & _
                                 " SET ing_numasiento = '" & strMaximo & "'" & _
                                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                                 " AND ing_codigo = '" & VSFGAdquisicion.TextMatrix(i, 3) & "' " & _
                                 " AND tip_ing_codigo='COM'"
                    Else
                        strSQL = " UPDATE egreso " & _
                                 " SET egr_numasiento = '" & strMaximo & "'" & _
                                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                                 " AND egr_codigo = '" & VSFGAdquisicion.TextMatrix(i, 3) & "' " & _
                                 " AND tip_egr_codigo='DPV'"
                    End If
                ElseIf Me.optFactura.Value = True Then
                    If VSFGAdquisicion.TextMatrix(i, 11) > 0 Then
                        strSQL = " UPDATE egreso " & _
                                 " SET egr_numasiento = '" & strMaximo & "'" & _
                                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                                 " AND egr_codigo = '" & VSFGAdquisicion.TextMatrix(i, 3) & "'" & _
                                 " AND tip_egr_codigo='FAC'"
                    Else
                        strSQL = " UPDATE ingreso " & _
                                 " SET ing_numasiento = '" & strMaximo & "'" & _
                                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                                 " AND ing_codigo = '" & VSFGAdquisicion.TextMatrix(i, 3) & "' " & _
                                 " AND tip_ing_codigo='DCL'"
                    End If
                Else
                    strSQL = " UPDATE ingreso " & _
                                 " SET ing_numasiento = '" & strMaximo & "'" & _
                                 " WHERE emp_codigo = '" & strEmpresa & "' " & _
                                 " AND ing_codigo = '" & VSFGAdquisicion.TextMatrix(i, 3) & "' " & _
                                 " AND tip_ing_codigo='IIM'"
                End If
                clsSql.Ejecutar strSQL, "M"
            End If
        Next i

        With VSFGdb
            For i = 1 To .Rows - 1
                'Ingresa el detalle del asiento del comprobante
                If .TextMatrix(i, 1) = "" Then
                    'Exit For
                Else
                    clsAsi.NuevoDetAsiento .TextMatrix(i, 1), "", FormatoD2(.TextMatrix(i, 3)), FormatoD2(.TextMatrix(i, 4))
                End If
            Next i
        End With
        MsgBox " Los datos han sido ingresado", vbInformation, "SisAdmi - Asientos"
    End If

    Call Limpiar
    Call Rango_Fecha
End Sub
Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Detecta cuando se ha dado un enter para enviar un tab
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
Private Sub Form_Load()
'Inicializa las clases para hacer distintas consultas

    clsAdq.Inicializar AdoConn, AdoConnMaster
    clsAsi.Inicializar AdoConn, AdoConnMaster
    clsPro.Inicializar AdoConn, AdoConnMaster
    clsSum.Inicializar AdoConn, AdoConnMaster
    clsAct.Inicializar AdoConn, AdoConnMaster
    clsMaxAsi.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    
    dtpFechaI.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaF.Value = Format(HoyDia, "yyyy-mm-dd")
    dtpFechaA.Value = Format(HoyDia, "yyyy-mm-dd")
    ' Extrae todas las cuentas de último nivel de una empresa
    strSQL = " SELECT cta_codigo FROM ctaconta " & _
             " WHERE emp_codigo = '" & strEmpresa & "' AND cta_subcta = 0 " & _
             " ORDER BY cta_codigo "
    'Ejecuta la consulta anterior
    clsSql.Ejecutar strSQL
    'Muestra los datos de los códigos de las cuentas en un datacombo
    Set dcmbCtaConta.RowSource = clsSql.adorec_Def.DataSource
    dcmbCtaConta.ListField = "cta_codigo"
    ' Extrae todas las cuentas de último nivel de una empresa
    strSQL = " SELECT suc_codigo,suc_nombre FROM sucursal " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY suc_codigo "
    'Ejecuta la consulta anterior
    clsSql.Ejecutar strSQL
    'Muestra los datos de los códigos de las cuentas en un datacombo
    Set dcmbSucursal.RowSource = clsSql.adorec_Def.DataSource
    dcmbSucursal.ListField = "suc_nombre"
    dcmbSucursal.BoundColumn = "suc_codigo"
    'Consulta para conocer la cuenta contable con la que trabajan las cuentas por cobrar
    strSQL = " SELECT par_texto,cta_nombre " & _
             " FROM parametro INNER JOIN ctaconta ON parametro.par_texto=ctaconta.cta_codigo AND parametro.emp_codigo=ctaconta.emp_codigo " & _
             " WHERE par_codigo='CXP' AND parametro.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        cta_pagar = clsSql.adorec_Def("par_texto")
        cta_pagar_n = clsSql.adorec_Def("cta_nombre")
    End If
    'Consulta para conocer la cuenta contable con la que trabajan las cuentas por cobrar
    strSQL = " SELECT tip_ing_ctaconta,cta_nombre " & _
              " FROM tipo_ingreso INNER JOIN ctaconta ON tipo_ingreso.tip_ing_ctaconta=ctaconta.cta_codigo AND tipo_ingreso.emp_codigo=ctaconta.emp_codigo " & _
              " WHERE tip_ing_codigo='COM' AND tipo_ingreso.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        tip_ing_ctaconta = clsSql.adorec_Def("tip_ing_ctaconta")
        tip_ing_ctaconta_n = clsSql.adorec_Def("cta_nombre")
    End If
    'Consulta para conocer la cuenta contable con la que trabajan las cuentas por cobrar
    strSQL = " SELECT tip_ing_ctaconta2,cta_nombre " & _
              " FROM tipo_ingreso INNER JOIN ctaconta ON tipo_ingreso.tip_ing_ctaconta2=ctaconta.cta_codigo AND tipo_ingreso.emp_codigo=ctaconta.emp_codigo " & _
              " WHERE tip_ing_codigo='COM' AND tipo_ingreso.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        tip_ing_ctaconta2 = clsSql.adorec_Def("tip_ing_ctaconta2")
        tip_ing_ctaconta_n2 = clsSql.adorec_Def("cta_nombre")
    End If
    'Consulta para conocer la cuenta contable con la que trabajan las cuentas por cobrar
    strSQL = " SELECT tip_ing_ctaconta,cta_nombre " & _
              " FROM tipo_ingreso INNER JOIN ctaconta ON tipo_ingreso.tip_ing_ctaconta=ctaconta.cta_codigo AND tipo_ingreso.emp_codigo=ctaconta.emp_codigo " & _
              " WHERE tip_ing_codigo='IIM' AND tipo_ingreso.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        tip_ing_ctaconta3 = clsSql.adorec_Def("tip_ing_ctaconta")
        tip_ing_ctaconta_n3 = clsSql.adorec_Def("cta_nombre")
    End If
    'Consulta para conocer la cuenta contable con la que trabajan el iva de compras
    strSQL = " SELECT par_texto,cta_nombre " & _
             " FROM parametro INNER JOIN ctaconta ON parametro.par_texto=ctaconta.cta_codigo AND parametro.emp_codigo=ctaconta.emp_codigo " & _
             " WHERE par_codigo='IVAC' AND parametro.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        iva_compra = clsSql.adorec_Def("par_texto")
        iva_compra_n = clsSql.adorec_Def("cta_nombre")
    End If
    
    'Consulta para conocer la cuenta contable con la que trabajan las cuentas por cobrar
    strSQL = " SELECT par_texto,cta_nombre " & _
             " FROM parametro INNER JOIN ctaconta ON parametro.par_texto=ctaconta.cta_codigo AND parametro.emp_codigo=ctaconta.emp_codigo " & _
             " WHERE par_codigo='CXC' AND parametro.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        cta_cobrar = clsSql.adorec_Def("par_texto")
        cta_cobrar_n = clsSql.adorec_Def("cta_nombre")
    End If
    'Consulta para conocer la cuenta contable con la que trabajan el iva de compras
    strSQL = " SELECT par_texto,cta_nombre " & _
             " FROM parametro INNER JOIN ctaconta ON parametro.par_texto=ctaconta.cta_codigo AND parametro.emp_codigo=ctaconta.emp_codigo " & _
             " WHERE par_codigo='IVAV' AND parametro.emp_codigo='" & strEmpresa & "'"
    clsSql.Ejecutar strSQL
    If (clsSql.adorec_Def.RecordCount > 0) Then
        iva_venta = clsSql.adorec_Def("par_texto")
        iva_venta_n = clsSql.adorec_Def("cta_nombre")
    End If
    
    Call limpiarFxGD
    Call Rango_Fecha
    'boton aceptar desactivado
    cmdAceptar.Enabled = False

End Sub

Private Sub optCompraLocal_Click()
    frmCtaConta.Visible = False
    frmSucursal.Visible = False
    Rango_Fecha
End Sub

Private Sub optFactura_Click()
    dcmbCtaConta.Text = ""
    frmCtaConta.Visible = True
    frmCtaConta.Caption = "Descuento en Ventas"
    frmSucursal.Visible = True
    Rango_Fecha
End Sub

Private Sub optImportacion_Click()
    dcmbCtaConta.Text = ""
    frmCtaConta.Visible = True
    frmCtaConta.Caption = "Importación"
    frmSucursal.Visible = False
    Rango_Fecha
End Sub

Private Sub TxtSubTotal_Change()
    TxtSubTotal = Format(Val(TxtSubTotal), "##0.00")
End Sub
Private Sub TxtIva_Change()
    TxtIva = Format(Val(TxtIva), "##0.00")
End Sub
Private Sub txtTotal_Change()
    TxtTotal = Format(Val(TxtTotal), "##0.00")
End Sub
Private Sub txtTotalDebe_Change()
    txtTotalDebe = Format(Val(txtTotalDebe), "##0.00")
End Sub
Private Sub txtTotalHaber_Change()
 txtTotalHaber = Format(Val(txtTotalHaber), "##0.00")
End Sub
Private Sub VSFGdb_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 And Col = 2 And Col = 3 And Col = 4 Then
         Cancel = True
    End If
End Sub
Private Sub VSFGAdquisicion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Or Col = 8 Or Col = 9 Then
    Cancel = True
    VSFGAdquisicion.Col = 1
  End If
End Sub
Private Sub VSFGAdquisicion_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'este es el check box
    If ban = 1 Then
        If Col = 1 And Row > 0 Then
            If Abs(VSFGAdquisicion.TextMatrix(Row, 1)) = 1 Then
                VSFGAdquisicion.Select Row, 1, Row, 11
                VSFGAdquisicion.FillStyle = flexFillRepeat
                VSFGAdquisicion.CellBackColor = &HC0FFFF
                VSFGAdquisicion.Select Row, 9
            ElseIf VSFGAdquisicion.TextMatrix(Row, 1) = "0" Then
              VSFGAdquisicion.Select Row, 1, Row, 11
              VSFGAdquisicion.FillStyle = flexFillRepeat
              VSFGAdquisicion.CellBackColor = &HFFFFFF
              VSFGAdquisicion.Select Row, 9
            End If
        Call DebeHaber
        Call Cal_Total
        'Call CalTotalDebeHaber
        End If
    End If
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim x, Y  As Integer
    VSFGAdquisicion.Rows = 1
    VSFGAdquisicion.Clear 1
End Sub
