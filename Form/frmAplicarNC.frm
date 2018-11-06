VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAplicarNC 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicar Notas de Crédito"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAplicarNC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   14850
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5858
      TabIndex        =   2
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Aplicación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   14640
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   5760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame5 
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
         Height          =   975
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   14175
         Begin MSDataListLib.DataCombo cmbNegocio 
            Height          =   315
            Left            =   960
            TabIndex        =   20
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
            Left            =   6600
            TabIndex        =   21
            Top             =   240
            Width           =   6615
            _ExtentX        =   11668
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
         Begin MSDataListLib.DataCombo cmbDirector 
            Height          =   315
            Left            =   6600
            TabIndex        =   22
            Top             =   600
            Width           =   6615
            _ExtentX        =   11668
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Negocio:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gerente Zona:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5445
            TabIndex        =   24
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Director:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5880
            TabIndex        =   23
            Top             =   705
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Facturas Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   2655
         Begin VB.CheckBox chkIncluirPedidos 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Incluir Pedidos"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton cmdMostrarAplicacion 
            Caption         =   "&Mostrar Posible Aplicación"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   1560
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpFechaFacFin 
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
            Left            =   960
            TabIndex        =   16
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69599235
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpFechaFacIni 
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
            Left            =   960
            TabIndex        =   17
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69599235
            CurrentDate     =   37463
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   13
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   12
            Top             =   645
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Notas de Crédito Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
         Begin VB.CommandButton cmdMostrarNC 
            Caption         =   "&Mostrar Notas de Credito"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox chkSelTodo 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Seleccionar TODO"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpFechaNCIni 
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
            Left            =   960
            TabIndex        =   14
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69599235
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker dtpFechaNCFin 
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
            Left            =   960
            TabIndex        =   15
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   69599235
            CurrentDate     =   37463
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   6
            Top             =   645
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   360
            TabIndex        =   5
            Top             =   285
            Width           =   510
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1695
         Left            =   3120
         TabIndex        =   3
         Top             =   1320
         Width           =   11295
         _cx             =   94981587
         _cy             =   94964654
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
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarNC.frx":030A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   3735
         Left            =   3120
         TabIndex        =   9
         Top             =   3480
         Width           =   11295
         _cx             =   93933011
         _cy             =   93919676
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
         Rows            =   1
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAplicarNC.frx":042C
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
      Begin NEED2.uctrVSFG ucrtVSFG2 
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   3120
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7538
      TabIndex        =   0
      Top             =   7560
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   120
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
End
Attribute VB_Name = "frmAplicarNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de la Lista de Precio poder modificar,              #
'#  crear o eliminar las listas                                                 #
'#  frmSelListaPrecio V1.0                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las listas que al momento estan                      #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  lista modificarla o eliminar las listas ya creadas.                         #
'#  Desde esta ventana se llama a la ventana frmListaPrecio en la que se crea   #
'#  y modifica las listas                                                       #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las nuevas listas, se               #
'#               modifican los datos de las listas y se eliminan.               #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As New clsConsulta
Private strSql As String

Private Sub cmbNegocio_Validate(Cancel As Boolean)
    cargarGZDir
End Sub

Public Sub cmdAplicar_Click()
    Dim i As Long
    Dim clsCta As New clsCtaXx
    Dim strMaxPago As String
    Dim clsIngreso As New clsInventario
    
    Me.MousePointer = 11
    
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    
    For i = 1 To VSFG2.Rows - 1
        'Aplica la Nota de Credito
        If Abs(VSFG2.TextMatrix(i, 0)) = 1 Then
            VSFG2.Cell(flexcpBackColor, i, 0, i, VSFG2.Cols - 1) = vbYellow
            VSFG2.ShowCell i, 1
            VSFG2.Refresh
            If chkIncluirPedidos.Value = 0 Then
                clsCta.strPersona = VSFG2.TextMatrix(i, 2)
                clsCta.strTipoCta = "C"
                clsCta.AplicaNC2 VSFG2.TextMatrix(i, 9), VSFG2.TextMatrix(i, 1), VSFG2.TextMatrix(i, 10)
            Else
                If VSFG2.TextMatrix(i, 1) <> "P" Then
                    clsCta.strPersona = VSFG2.TextMatrix(i, 2)
                    clsCta.strTipoCta = "C"
                    clsCta.AplicaNC2 VSFG2.TextMatrix(i, 9), VSFG2.TextMatrix(i, 1), VSFG2.TextMatrix(i, 10)
                Else
                    
                    strSql = " BEGIN TRAN "
                    clsCon_Def.Ejecutar strSql, "M"
                    strSql = " SELECT MAX(doc_pag_ped_codigo) as Num FROM doc_pago_pedido WITH (TABLOCKX) " & _
                             " WHERE emp_codigo='" & strEmpresa & "'" & _
                             " AND doc_pag_ped_codigo LIKE 'DCL%'"
                    clsCon_Def.Ejecutar strSql, "M"
                    strMaxPago = "DCL" & Format(FormatoD0(Mid(clsCon_Def.adorec_Def("Num"), 4)) + 1, "0000000000")
                
                    ffch = Format(HoyDia, "yyyy-mm-dd")
                    Fecha = Format(HoyDia, "yyyy-mm-dd")
                    
                    fechac = Format(HoyDia, "yyyy-mm-dd") '*********
                    fechac = ffch
                    ValorPago = FormatoD2(VSFG2.TextMatrix(i, 10))
                    ElAsiento = "NULL"
                
                    intAnticipo = 0
                    
                    'txtDescripcion = "NOTA DE CREDITO: " & VSFG2.TextMatrix(i, 9)
                    strSql = " INSERT INTO doc_pago_pedido (emp_codigo, doc_pag_ped_tipo,doc_pag_ped_codigo,doc_pag_codigo, tip_doc_pag_codigo, " & _
                             " ban_codigo, doc_pag_ped_numero, doc_pag_ped_fecha_recepcion, doc_pag_ped_fecha_doc ," & _
                             " per_codigo, doc_pag_ped_valor, doc_pag_ped_observacion, doc_pag_ped_estado," & _
                             " doc_pag_ped_pendiente,doc_pag_ped_anticipo, doc_pag_ped_fechamod, doc_pag_ped_usumod," & _
                             " per_codigo_ch,ped_codigo)" & _
                             " VALUES ('" & strEmpresa & "','N','" & strMaxPago & "','" & strMaxPago & "', 'DCL', " & _
                             " '', '" & VSFG2.TextMatrix(i, 9) & "','" & ffch & "', '" & Fecha & "'," & _
                             " '" & VSFG2.TextMatrix(i, 2) & "', '" & ValorPago & "', '" & UCase("NOTA DE CREDITO: " & VSFG2.TextMatrix(i, 9)) & "', 'GIRADO'," & _
                             " 0,0, CURRENT_TIMESTAMP, '" & strUsuario & "'," & _
                             " '','" & VSFG2.TextMatrix(i, 4) & "') "
    
                    clsCon_Def.Ejecutar strSql, "M"
                    strSql = " COMMIT TRAN "
                    clsCon_Def.Ejecutar strSql, "M"
                    strSql = " UPDATE ingreso SET ing_saldo=ing_saldo+'" & ValorPago & "' WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='DCL' AND ing_codigo='" & VSFG2.TextMatrix(i, 9) & "'"
                    clsCon_Def.Ejecutar strSql, "M"
                End If
            End If
        End If
    Next i
    
''GUARDAR EN EXCEL
    If VSFG2.Rows > 1 Then
    Dim sDir As String
    sDir = CurDir
    cdArchivo.FileName = ""
    While cdArchivo.FileName = ""
        cdArchivo.ShowSave
    Wend
    ChDir sDir
    If (cdArchivo.FileName <> "") Then
        VSFG2.SaveGrid cdArchivo.FileName, flexFileExcel, flexXLSaveFixedCells
    End If
    End If
    cmdAplicar.Enabled = False
    Me.MousePointer = 0
    MsgBox "Notas de Crédito Aplicadas", vbInformation, "Facturas"
    'Unload Me
End Sub

Public Sub cmdMostrarAplicacion_Click()
    Dim i As Long
    VSFG2.Rows = 1
    If VSFG.Rows > 1 Then
        VSFG.Cell(flexcpBackColor, 1, 0, VSFG.Rows - 1, VSFG.Cols - 1) = vbWhite
        For i = 1 To VSFG.Rows - 1
            VSFG.ShowCell i, 0
            If Abs(VSFG.TextMatrix(i, 0)) = 1 Then
    '            If VSFG.TextMatrix(i, 1) = "20040039947" Then
    '                MsgBox "AA"
    '            End If
                RevisarAplicacion i
            End If
        Next i
    End If
    cmdAplicar.Enabled = True
End Sub

Private Sub RevisarAplicacion(Linea As Long)
    Dim SaldoNC As Double
    SaldoNC = VSFG.TextMatrix(Linea, 9)
    VSFG.ShowCell Linea, 1
    VSFG.Refresh
    strSql = " SELECT CAST(cuenta_p_c.cue_p_c_codigo as varchar) as cue_p_c_codigo, cuenta_p_c.per_codigo, CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as cli, " & _
             " persona.per_direccion2,COALESCE(fp.for_pag_nombre,'') as fp1,COALESCE(fpi.for_pag_nombre,'') as fp2, " & _
             " CONCAT(gez.per_apellido,' ',gez.per_nombre,' (',gez.per_ruc,')') as gz, " & _
             " CONCAT(dir.per_apellido,' ',dir.per_nombre,' (',dir.per_ruc,')') as di, " & _
             " CONCAT(emp.per_apellido,' ',emp.per_nombre,' (',emp.per_ruc,')') as em,cue_p_c_egr_codigo, " & _
             " cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor," & _
             " cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as saldo " & _
             " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo " & _
             " LEFT JOIN forma_pago fp ON persona.emp_codigo=fp.emp_codigo AND persona.for_pag_codigo=fp.for_pag_codigo" & _
             " LEFT JOIN forma_pago fpi ON persona.emp_codigo=fpi.emp_codigo AND persona.for_pag_codigo_imp=fpi.for_pag_codigo" & _
             " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
             " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
             " LEFT JOIN persona gez ON persona.emp_codigo=gez.emp_codigo AND persona.per_codigo_ref=gez.per_codigo AND gez.per_es_gz=1 " & _
             " LEFT JOIN persona dir ON persona.emp_codigo=dir.emp_codigo AND persona.per_codigo_ref2=dir.per_codigo AND dir.per_es_di=1 " & _
             " LEFT JOIN persona emp ON persona.emp_codigo=emp.emp_codigo AND persona.per_codigo_ref3=emp.per_codigo AND emp.per_es_em=1 " & _
             " WHERE cuenta_p_c.per_codigo IN ('" & VSFG.TextMatrix(Linea, 3) & "') AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
             " AND tip_doc_cue_codigo=1 AND cue_p_c_fechaemision BETWEEN '" & dtpFechaFacIni.Value & "' AND '" & dtpFechaFacFin.Value & "' AND cue_p_c_fechaemision>='" & VSFG.TextMatrix(Linea, 2) & "'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo, cuenta_p_c.per_codigo, persona.per_apellido, persona.per_nombre, persona.per_ruc, persona.per_direccion2, fp.for_pag_nombre, fpi.for_pag_nombre, gez.per_apellido, gez.per_nombre, gez.per_ruc,dir.per_apellido, dir.per_nombre, dir.per_ruc, emp.per_apellido, emp.per_nombre, emp.per_ruc, cue_p_c_egr_codigo,cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor, com_ret_total HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)!=0  "
    If chkIncluirPedidos.Value = 1 Then
        strSql = strSql & " UNION " & _
                 " SELECT 'P', pedido.per_codigo, CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')') as cli, " & _
                 " persona.per_direccion2,COALESCE(fp.for_pag_nombre,'') as fp1,COALESCE(fpi.for_pag_nombre,'') as fp2, " & _
                 " CONCAT(gez.per_apellido,' ',gez.per_nombre,' (',gez.per_ruc,')') as gz, " & _
                 " CONCAT(dir.per_apellido,' ',dir.per_nombre,' (',dir.per_ruc,')') as di, " & _
                 " CONCAT(emp.per_apellido,' ',emp.per_nombre,' (',emp.per_ruc,')') as em,pedido.ped_codigo, " & _
                 " ped_fecha, ped_fecha, " & _
                 " ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(persona.per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(persona.per_dcto,0))/100.00,2))) * (100.00+par_numero)/100.00,2)," & _
                 " ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(persona.per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(persona.per_dcto,0))/100.00,2))) * (100.00+par_numero)/100.00,2) - COALESCE(doc_pag_valor,0.00) as saldo "
        strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                 " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                 " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " LEFT JOIN forma_pago fp ON persona.emp_codigo=fp.emp_codigo AND persona.for_pag_codigo=fp.for_pag_codigo" & _
                 " LEFT JOIN forma_pago fpi ON persona.emp_codigo=fpi.emp_codigo AND persona.for_pag_codigo_imp=fpi.for_pag_codigo" & _
                 " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                 " FROM doc_pago_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                 " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                 " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                 " AND pedido.per_codigo=pag.per_codigo " & _
                 " LEFT JOIN persona gez ON persona.emp_codigo=gez.emp_codigo AND persona.per_codigo_ref=gez.per_codigo AND gez.per_es_gz=1 " & _
                 " LEFT JOIN persona dir ON persona.emp_codigo=dir.emp_codigo AND persona.per_codigo_ref2=dir.per_codigo AND dir.per_es_di=1 " & _
                 " LEFT JOIN persona emp ON persona.emp_codigo=emp.emp_codigo AND persona.per_codigo_ref3=emp.per_codigo AND emp.per_es_em=1 " & _
                 " WHERE pedido.per_codigo IN ('" & VSFG.TextMatrix(Linea, 3) & "') AND pedido.emp_codigo = '" & strEmpresa & "' AND pedido.ped_estado = 0 " & _
                 " AND LEFT(ped_fechamod,10) BETWEEN '" & dtpFechaFacIni.Value & "' AND '" & dtpFechaFacFin.Value & "' AND LEFT(ped_fechamod,10)>=LEFT('" & VSFG.TextMatrix(Linea, 2) & "',10) " & _
                 " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
                 " GROUP BY pedido.per_codigo, persona.per_apellido, persona.per_nombre, persona.per_ruc, persona.per_direccion2,fp.for_pag_nombre, fpi.for_pag_nombre, gez.per_apellido, gez.per_nombre, gez.per_ruc, dir.per_apellido, dir.per_nombre, dir.per_ruc,emp.per_apellido, emp.per_nombre, emp.per_ruc, pedido.ped_codigo,ped_fecha,par_numero,doc_pag_valor HAVING round(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(persona.per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(persona.per_dcto,0))/100.00,2))) * (100.00+par_numero)/100.00,2) - COALESCE(doc_pag_valor,0.00),2)!=0 "
    End If
    strSql = strSql & " ORDER BY cue_p_c_fechaemision,cue_p_c_egr_codigo,cue_p_c_codigo "
    clsCon_Def.Ejecutar strSql
    While SaldoNC > 0 And Not clsCon_Def.adorec_Def.EOF
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            SaldoFactura = clsCon_Def.adorec_Def("saldo")
            If Left(clsCon_Def.adorec_Def("cue_p_c_codigo"), 1) <> "P" Then
                If clsCon_Def.adorec_Def("cue_p_c_codigo") = VSFG2.TextMatrix(VSFG2.Rows - 1, 1) Then
                    SaldoFactura = FormatoD2(FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 8)) - FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 10)))
                End If
            Else
                If clsCon_Def.adorec_Def("cue_p_c_codigo") = VSFG2.TextMatrix(VSFG2.Rows - 1, 4) Then
                    SaldoFactura = FormatoD2(FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 8)) - FormatoD2(VSFG2.TextMatrix(VSFG2.Rows - 1, 10)))
                End If

            End If
            If SaldoFactura > 0 Then
            VSFG2.AddItem "1" & vbTab & _
                          clsCon_Def.adorec_Def("cue_p_c_codigo") & vbTab & _
                          clsCon_Def.adorec_Def("per_codigo") & vbTab & _
                          clsCon_Def.adorec_Def("cli") & vbTab & _
                          clsCon_Def.adorec_Def("cue_p_c_egr_codigo") & vbTab & _
                          clsCon_Def.adorec_Def("cue_p_c_fechaemision") & vbTab & _
                          clsCon_Def.adorec_Def("cue_p_c_fechapropuesta") & vbTab & _
                          clsCon_Def.adorec_Def("cue_p_c_valor") & vbTab & _
                          SaldoFactura & vbTab & _
                          VSFG.TextMatrix(Linea, 1) & vbTab & "0" & vbTab & VSFG.TextMatrix(Linea, 2) & vbTab & _
                          clsCon_Def.adorec_Def("gz") & vbTab & _
                          clsCon_Def.adorec_Def("di") & vbTab & _
                          clsCon_Def.adorec_Def("em") & vbTab & _
                          clsCon_Def.adorec_Def("per_direccion2") & vbTab & _
                          clsCon_Def.adorec_Def("fp1") & vbTab & _
                          clsCon_Def.adorec_Def("fp2")
            VSFG2.ShowCell VSFG2.Rows - 1, 1
            VSFG2.Refresh
            VSFG2.TextMatrix(VSFG2.Rows - 1, 9) = VSFG.TextMatrix(Linea, 1)
            VSFG.Cell(flexcpBackColor, Linea, 0, Linea, VSFG.Cols - 1) = vbYellow
            If FormatoD2(SaldoFactura) >= FormatoD2(SaldoNC) Then
                VSFG2.TextMatrix(VSFG2.Rows - 1, 10) = SaldoNC
                SaldoNC = 0
            ElseIf FormatoD2(SaldoFactura) < FormatoD2(SaldoNC) Then
                VSFG2.TextMatrix(VSFG2.Rows - 1, 10) = FormatoD2(SaldoFactura)
                SaldoNC = SaldoNC - FormatoD2(SaldoFactura)
            End If
            End If
            clsCon_Def.adorec_Def.MoveNext
        Else
            SaldoNC = 0
        End If
    Wend
    
End Sub

Public Sub cmdMostrarNC_Click()
    Dim strDiGz As String
    strDiGz = ""
    If cmbGerente.BoundText <> "-1" And cmbGerente.BoundText <> "" Then
         strDiGz = strDiGz & " AND persona.per_codigo_ref='" & cmbGerente.BoundText & "'"
    End If
    If cmbDirector.BoundText <> "-1" And cmbDirector.BoundText <> "" Then
         strDiGz = strDiGz & " AND persona.per_codigo_ref2='" & cmbDirector.BoundText & "'"
    End If
    
    strSql = " SELECT '" & IIf(chkSelTodo.Value = 1, 1, 0) & "' as selec,ingreso.ing_codigo, ing_fecha,persona.per_codigo, CONCAT(persona.per_apellido,' ',persona.per_nombre,' (',persona.per_ruc,')'), " & _
             " ing_subtotal,ing_dcto,ing_impuesto,ing_total,ing_total-ing_saldo as saldo " & _
             " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo " & _
             " AND ingreso.per_codigo=persona.per_codigo " & strDiGz & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.tip_ing_codigo='DCL' " & _
             " AND ingreso.ing_fecha BETWEEN '" & dtpFechaNCIni.Value & "' AND '" & dtpFechaNCFin.Value & "' " & _
             " AND ing_anulado=0 AND ROUND(ing_total-ing_saldo,2)!=0 " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "' AND ingreso.per_codigo NOT IN ('C00294','C00346','C01511','C01538','C02218','C45560C','C45571C','C103291','C125486','C128394') " & _
             " ORDER BY CONCAT(persona.per_apellido,' ',persona.per_nombre),ing_fecha,ing_codigo"
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
             
    
End Sub


Private Sub Command1_Click()
    LiberarYBajarPedidos False, cmbNegocio.BoundText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Public Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpFechaNCIni.Value = HoyDia
    dtpFechaNCFin.Value = HoyDia
    dtpFechaFacIni.Value = HoyDia
    dtpFechaFacFin.Value = HoyDia
    Set ucrtVSFG2.VSFGControl = VSFG2
    ucrtVSFG2.Inicializar False, False, False
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
        cargarGZDir
        
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

Private Sub cargarGZDir()

        
        strSql = " SELECT '-1' as codigo,' Todos los Gerentes de Zona' as nombre " & _
                 " UNION " & _
                 " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
                 " FROM persona p1 " & _
                 " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
                 " AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND p1.per_es_gz=1 " & _
                 " ORDER BY 2 "
        clsCon_Def.Ejecutar strSql
        Set cmbGerente.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbGerente.ListField = "nombre"
        cmbGerente.BoundColumn = "codigo"
        
        strSql = " SELECT '-1' as codigo,' Todos los Directores' as nombre " & _
                 " UNION " & _
                 " SELECT DISTINCT p1.per_codigo as codigo,CONCAT(p1.per_apellido,' ',p1.per_nombre,' (',p1.per_ruc,')') as nombre " & _
                 " FROM persona p1 " & _
                 " WHERE p1.emp_codigo= '" & strEmpresa & "' AND p1.cat_p_tipo = 'C' " & _
                 " AND p1.tip_ped_codigo='" & cmbNegocio.BoundText & "' " & _
                 " AND p1.per_es_di=1 " & _
                 " ORDER BY 2 "
        clsCon_Def.Ejecutar strSql
        Set cmbDirector.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbDirector.ListField = "nombre"
        cmbDirector.BoundColumn = "codigo"
        
        cmbGerente.BoundText = "-1"
        cmbDirector.BoundText = "-1"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

