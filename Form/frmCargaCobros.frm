VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCargaCobros 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Cobros"
   ClientHeight    =   8700
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
   Icon            =   "frmCargaCobros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   14850
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Facturas"
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
      TabIndex        =   1
      Top             =   0
      Width           =   14640
      Begin VB.CommandButton cmdImprimirAnticipos 
         Caption         =   "&Imp.Anticipos"
         Height          =   375
         Left            =   13080
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   11880
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtTotalProcesar 
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
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   7800
         Width           =   1215
      End
      Begin VB.TextBox txtTotalArchivo 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   4320
         Begin MSDataListLib.DataCombo dcmbDocumento 
            Height          =   315
            Left            =   945
            TabIndex        =   27
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   1005
         Left            =   3720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   6960
         Width           =   3135
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
         Left            =   3000
         TabIndex        =   15
         Top             =   5160
         Width           =   3855
         Begin VB.TextBox txtNumero 
            Height          =   285
            Left            =   1530
            TabIndex        =   16
            Top             =   1080
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo dcmbCuenta 
            Height          =   315
            Left            =   1530
            TabIndex        =   17
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcmbBancoE 
            Height          =   315
            Left            =   1530
            TabIndex        =   18
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Bancaria:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   21
            Top             =   765
            Width           =   1245
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº. Documento:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   20
            Top             =   1110
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   405
            Width           =   510
         End
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
         Left            =   120
         TabIndex        =   11
         Top             =   5160
         Width           =   2535
         Begin VB.OptionButton optBanco 
            BackColor       =   &H00DDDDDD&
            Caption         =   "En Banco o Cajas"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   120
            TabIndex        =   12
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
            TabIndex        =   13
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
            Format          =   65798147
            CurrentDate     =   37463
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C3DBD1&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   780
            Width           =   495
         End
      End
      Begin VB.TextBox txtDescripciont 
         Enabled         =   0   'False
         Height          =   765
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   7320
         Width           =   2415
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   300
         Width           =   4275
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   10440
         TabIndex        =   5
         Top             =   300
         Width           =   375
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   14415
         _cx             =   59597650
         _cy             =   59577542
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaCobros.frx":030A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   3255
         Left            =   7320
         TabIndex        =   4
         Top             =   4560
         Width           =   7095
         _cx             =   97988835
         _cy             =   97982061
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
         Cols            =   33
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCargaCobros.frx":0398
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
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   10320
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin NEED2.dtpFecha dtpFecha 
         Height          =   315
         Left            =   3690
         TabIndex        =   8
         Top             =   300
         Width           =   1335
         _extentx        =   2355
         _extenty        =   556
         value           =   41821.4661111111
      End
      Begin MSDataListLib.DataCombo dcmbTipo 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   7320
         TabIndex        =   29
         Top             =   4200
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
      End
      Begin NEED2.uctrVSFG uctrVSFG1 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Procesar:"
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
         Left            =   11640
         TabIndex        =   34
         Top             =   7830
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Archivo:"
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
         TabIndex        =   32
         Top             =   4230
         Width           =   1125
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de nota:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   7005
         Width           =   930
      End
      Begin VB.Label lbldescripcion1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   7365
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3120
         TabIndex        =   9
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5400
         TabIndex        =   6
         Top             =   345
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7538
      TabIndex        =   0
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   8280
      Width           =   5685
   End
End
Attribute VB_Name = "frmCargaCobros"
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
Private FacYaReg As Long
Private NoHayCli As Long
Private BloqueadoCli As Long
Private ProductoProblema As Long
Private errorCarga As Boolean


Private Sub PonerDescripcion2(Linea As Long)
    Dim Cadena1 As String
    Dim Cadena2 As String
    Dim Cadena3 As String
    Dim Cadena4 As String
    Dim Cadena5 As String
    
    If dcmbDocumento <> "" Then
        Cadena1 = dcmbDocumento & " "
    End If
    'transaccion
    If txtNumero <> "" Then
        Cadena2 = VSFG2.TextMatrix(Linea, 10) & " "
    End If
    If dcmbBancoE <> "" Then
        Cadena3 = dcmbBancoE & " "
    End If
    'cliente
    Descripcion = "CLI: " & VSFG2.TextMatrix(Linea, 3) & " DOC: " & dcmbDocumento.Text & "(" & dcmbBancoE.Text & ") No:" & Cadena2 & " ABONO: " & VSFG2.TextMatrix(Linea, 8)
    If Descripcion <> "" Then
        Cadena4 = Descripcion & " - "
    End If
    If dcmbBeneficiario <> "" Then
        Cadena5 = VSFG2.TextMatrix(Linea, 3) & " - "
    End If
    txtDescripcion = Cadena5 & Cadena4 & Cadena1 & Cadena2 & Cadena3
End Sub


Private Sub cmdAplicar_Click()
    Dim ElAsiento As String
    Dim maxpago As String, maxpagoAux As String, booPasar As Boolean, respuesta As Integer, booGrabar As Boolean
    Dim intAnticipo As Integer
    Dim ValorPago As Double
    Dim NumAsiAnticipo() As String
    Dim CantidadAnticipo As Long
    CantidadAnticipo = 0
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
    Dim clsdoc As New clsConsulta
    Dim clscdo As New clsConsulta
    clsdoc.Inicializar AdoConn, AdoConnMaster
    clscdo.Inicializar AdoConn, AdoConnMaster
'Comprueba que todos los datos esten ingresados
    ReDim NumAsiAnticipo(VSFG2.Rows)
    For i = 1 To VSFG2.Rows - 1
        Descripcion = UCase("Depósito Banco: " + " " + dcmbBancoE.Text + " " + " CASH - Archivo:" & txtNumero.Text)
        Pendiente = 0
        EstadoCH = "COBRADO"
        campoAsiento = "asi_numasiento"
        If VSFG2.TextMatrix(i, 0) <> "P" Then
            '*************************************************
            booPasar = True
            booGrabar = True
         
            If booGrabar = True Then
            
                strSql = " BEGIN TRAN "
                clsdoc.Ejecutar strSql, "M"
                strSql = " SELECT MAX(doc_pag_codigo) as Num FROM doc_pago WITH (TABLOCKX) " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND doc_pag_codigo LIKE 'CSH%'"
                clsdoc.Ejecutar strSql, "M"
                maxpago = "CSH" & Format(FormatoD0(Mid(clsdoc.adorec_Def("Num"), 4)) + 1, "0000000000")
            
                ffch = Format(VSFG2.TextMatrix(i, 9), "yyyy-mm-dd")
                Fecha = Format(VSFG2.TextMatrix(i, 16), "yyyy-mm-dd")
                
                fechac = Format(VSFG2.TextMatrix(i, 9), "yyyy-mm-dd") '*********
                fechac = ffch
                ValorPago = FormatoD2(VSFG2.TextMatrix(i, 8))
                ElAsiento = "NULL"
                If Left(VSFG2.TextMatrix(i, 4), 8) = "ANTICIPO" And Left(VSFG2.TextMatrix(i, 5), 8) = "ANTICIPO" Then
                    intAnticipo = 1
                    intPend = 0
    ''''''''''''''''revisar descripcion de anticipo
                    txtDescripcion.Text = Descripcion & " " & VSFG2.TextMatrix(i, 3) & " ANTICIPO RECIBIDO " & VSFG2.TextMatrix(i, 5)
                    CantidadAnticipo = CantidadAnticipo + 1
                strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, tip_doc_pag_codigo, ban_codigo, doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_doc ," & _
                         " per_codigo, doc_pag_valor, doc_pag_observacion, doc_pag_estado,doc_pag_pendiente,doc_pag_anticipo, doc_pag_fechamod, doc_pag_usumod,ped_codigo,per_codigo_ch)" & _
                         " VALUES ('" & maxpago & "', '" & strEmpresa & "', '" & dcmbDocumento.BoundText & "', '" & VSFG2.TextMatrix(i, 14) & "', '" & VSFG2.TextMatrix(i, 10) & "','" & ffch & "', '" & Fecha & "'," & _
                         " '" & VSFG2.TextMatrix(i, 2) & "', '" & VSFG2.TextMatrix(i, 8) & "', '" & UCase(txtDescripcion) & "', 'GIRADO','" & intPend & "','" & intAnticipo & "', CURRENT_TIMESTAMP, '" & strUsuario & "','" & FormatoD0(Right(VSFG2.TextMatrix(i, 5), 11)) & "','" & VSFG2.TextMatrix(i, 31) & "') "
                Else
                    intAnticipo = 0
                    txtDescripcion.Text = Descripcion & " " & VSFG2.TextMatrix(i, 3) & " FAC: " & VSFG2.TextMatrix(i, 4)
                    strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, tip_doc_pag_codigo, ban_codigo, doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_doc ," & _
                             " per_codigo, doc_pag_valor, doc_pag_observacion, doc_pag_estado,doc_pag_pendiente,doc_pag_anticipo, doc_pag_fechamod, doc_pag_usumod,per_codigo_ch)" & _
                             " VALUES ('" & maxpago & "', '" & strEmpresa & "', '" & dcmbDocumento.BoundText & "', '" & VSFG2.TextMatrix(i, 14) & "', '" & VSFG2.TextMatrix(i, 10) & "','" & ffch & "', '" & Fecha & "'," & _
                             " '" & VSFG2.TextMatrix(i, 2) & "', '" & VSFG2.TextMatrix(i, 8) & "', '" & UCase(txtDescripcion) & "', 'GIRADO','" & intPend & "','" & intAnticipo & "', CURRENT_TIMESTAMP, '" & strUsuario & "','" & VSFG2.TextMatrix(i, 31) & "') "
                End If
                clsdoc.Ejecutar strSql, "M"
                strSql = " COMMIT TRAN "
                clsdoc.Ejecutar strSql, "M"
                VSFG2.TextMatrix(i, 12) = maxpago
                'Calcula el máximo codigo de pago para la cuenta
                If intAnticipo = 0 Then
                     strSql = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                              " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                              "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                              "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                              " WHERE cuenta_p_c.cue_p_c_codigo= '" & VSFG2.TextMatrix(i, 1) & "' " & _
                              " AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'C'" & _
                              " GROUP BY pago.emp_codigo"
                    clsdoc.Ejecutar strSql
                    If clsdoc.adorec_Def.EOF Then
                        maxpag = 1
                    Else
                        maxpag = clsdoc.adorec_Def("pag") + 1
                    End If
                    
                    strSql = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, " & _
                             " pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & VSFG2.TextMatrix(i, 1) & "', 'C', '" & Val(maxpag) & "', '" & ffch & "', '" & ValorPago & "', " & _
                             " '" & VSFG2.TextMatrix(i, 10) & "', '" & UCase(txtDescripcion) & "', " & _
                             " '" & maxpago & "'," & ElAsiento & ",CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsdoc.Ejecutar strSql, "M"
               End If
                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "','" & maxpago & "',0, '*', '','" & FormatoD2(ValorPago) & "', 0 , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsdoc.Ejecutar strSql, "M"
                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "','" & maxpago & "',0, '" & VSFG2.TextMatrix(i, 11) & "', '', 0,'" & FormatoD2(ValorPago) & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsdoc.Ejecutar strSql, "M"
                
                '**************************************************************
                Dim clsAsientoE As New clsContable
                clsAsientoE.Inicializar AdoConn, AdoConnMaster
                clsAsientoE.NuevoAsiento "I", fechac, 0, 0, FormatoD2(ValorPago), Descripcion
                Descripcion = Descripcion & vbNewLine & "CLI: " & VSFG2.TextMatrix(i, 3) & " DOC: " & Me.dcmbDocumento.Text & "(" & VSFG2.TextMatrix(i, 15) & ") No:" & VSFG2.TextMatrix(i, 10) & " ABONO: " & ValorPago & Replace(" OBS:" & UCase(txtDescripcion.Text), vbNewLine, " ")
                'Actualiza asientos en pagos
                If intAnticipo = 0 Then
                    strSql = " UPDATE pago " & _
                             " SET asi_numasiento='" & clsAsientoE.NumAsiento & _
                             "' , pag_fechamod= CURRENT_TIMESTAMP, pag_usumod='" & strUsuario & "' " & _
                             " WHERE doc_pag_codigo= '" & maxpago & "' AND emp_codigo = '" & strEmpresa & "' " & _
                             " AND cue_p_c_tipo='C' "
                    clsdoc.Ejecutar strSql, "M"
                End If
                VSFG2.TextMatrix(i, 13) = clsAsientoE.NumAsiento
                'Actualiza la tabla doc_pago
                strSql = " UPDATE doc_pago " & _
                         " SET doc_pag_fecha_efec='" & fechac & _
                         "'," & campoAsiento & "='" & clsAsientoE.NumAsiento & _
                         "',doc_pag_pendiente='" & Pendiente & "', doc_pag_estado = '" & EstadoCH & "' , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                         " WHERE doc_pag_codigo= '" & maxpago & "' AND emp_codigo = '" & strEmpresa & "' "
                clsdoc.Ejecutar strSql, "M"
                If intAnticipo = 0 Then
                    strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo," & _
                             " max(doc_pago.doc_pag_fecha_doc) as fecha,cuenta_p_c.cue_p_c_valor,COALESCE(sum(p2.pag_monto),0),COALESCE(com_ret_total,0)," & _
                             " cuenta_p_c.cue_p_c_valor-COALESCE(sum(p2.pag_monto),0)-COALESCE(com_ret_total,0) as saldo " & _
                             " FROM cuenta_p_c INNER JOIN pago as p1 ON cuenta_p_c.cue_p_c_codigo=p1.cue_p_c_codigo " & _
                             " AND cuenta_p_c.cue_p_c_tipo=p1.cue_p_c_tipo " & _
                             " AND cuenta_p_c.emp_codigo=p1.emp_codigo " & _
                             " AND p1.doc_pag_codigo='" & maxpago & "' " & _
                             " INNER JOIN pago as p2 ON cuenta_p_c.cue_p_c_codigo=p2.cue_p_c_codigo " & _
                             " AND cuenta_p_c.cue_p_c_tipo=p2.cue_p_c_tipo " & _
                             " AND cuenta_p_c.emp_codigo=p2.emp_codigo " & _
                             " INNER JOIN doc_pago ON p2.doc_pag_codigo=doc_pago.doc_pag_codigo " & _
                             " AND p2.emp_codigo=doc_pago.emp_codigo " & _
                             " AND doc_pago.doc_pag_pendiente=0 AND doc_pago.doc_pag_estado!='ANULADO' " & _
                             " LEFT JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                             " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                             " AND cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                             " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                             " AND cuenta_p_c.cue_p_c_tipo='C' " & _
                             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_valor,com_ret_total "
                    clscdo.Ejecutar strSql, "M"
                    While Not clscdo.adorec_Def.EOF
                        If (FormatoD2(clscdo.adorec_Def("saldo")) <= 0) Then
                            strSql = " UPDATE cuenta_p_c " & _
                                     " SET cue_p_c_fechapago='" & clscdo.adorec_Def("fecha") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                                     " WHERE cue_p_c_tipo= 'C' " & _
                                     " AND cue_p_c_codigo= '" & clscdo.adorec_Def("cue_p_c_codigo") & _
                                     "' AND cue_p_c_egr_codigo = '" & clscdo.adorec_Def("cue_p_c_egr_codigo") & _
                                     "' AND emp_codigo = '" & strEmpresa & "' "
                            clsdoc.Ejecutar strSql, "M"
                        End If
                        clscdo.adorec_Def.MoveNext
                    Wend
                End If
                clsAsientoE.ModificarAsiento FormatoD2(ValorPago), FormatoD2(ValorPago), , , , Descripcion
                'ingreso del detalle del asiento
                'banco
                clsAsientoE.NuevoDetAsiento Me.dcmbCuenta.BoundText, "", FormatoD2(ValorPago), FormatoD2("0")
                'cliente
                clsAsientoE.NuevoDetAsiento VSFG2.TextMatrix(i, 11), "", FormatoD2("0"), FormatoD2(ValorPago)
                                        
                If optBanco.Value = True Then
                    'GENERACION DE LA NOTA DE CREDITO
                    'Calcula el código de la Nota de Crédito
                    strSql = " SELECT cta_ban_saldoreal, cta_ban_saldoprevisto " & _
                              " FROM cta_banco " & _
                              " WHERE cta_ban_numero = '" & dcmbCuenta.Text & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsdoc.Ejecutar strSql
                    If Not clsdoc.adorec_Def.EOF Then
                        saldoreal = clsdoc.adorec_Def("cta_ban_saldoreal") + txtValor
                        saldoPrevisto = clsdoc.adorec_Def("cta_ban_saldoprevisto") + txtValor
                    Else
                        saldoreal = txtValor
                        saldoPrevisto = txtValor
                    End If
                    'Guarda los datos de la Nota de Crédito
                    strSql = " BEGIN TRAN "
                    clsdoc.Ejecutar strSql, "M"
                    strSql = " SELECT COALESCE(max(not_d_c_codigo),0) as n FROM nota_d_c WITH (TABLOCKX) where emp_codigo='" & strEmpresa & "' AND tip_not_d_c='C' GROUP BY emp_codigo"
                    clsdoc.Ejecutar strSql, "M"
                    
                    strSql = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto,asi_numasiento,not_d_c_conciliado , not_d_c_fechamod, not_d_c_usumod) " & _
                             " VALUES ('C','" & clsdoc.adorec_Def("n") + 1 & "', '" & dcmbCuenta.Text & "', '" & dcmbBancoE.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & VSFG2.TextMatrix(i, 10) & "','" & fechac & "','" & Descripcion & "','" & ValorPago & "','" & clsAsientoE.NumAsiento & "',0, CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsdoc.Ejecutar strSql, "M"
                    strSql = " COMMIT TRAN "
                    clsdoc.Ejecutar strSql, "M"
                    'Actualiza los valores de los saldos
                    strSql = " UPDATE cta_banco " & _
                             " SET cta_ban_saldoreal= '" & saldoreal & "',cta_ban_saldoprevisto= '" & saldoPrevisto & "', cta_ban_fechamod = CURRENT_TIMESTAMP, cta_ban_usumod= '" & strUsuario & "'" & _
                             " WHERE cta_ban_numero = '" & dcmbCuenta.Text & " ' AND ban_codigo = '" & dcmbBancoE.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
                    clsdoc.Ejecutar strSql, "M"
    
                End If
                
                '****************************************************
                If intAnticipo = 1 Then
                    NumAsiAnticipo(CantidadAnticipo) = clsAsientoE.NumAsiento
                End If
                '******************************************************
                
                
                Set clsAsiento = Nothing
            End If
        Else
        
'''''''******PAGO PEDIDO
            '*************************************************
            booPasar = True
            booGrabar = True
         
            If booGrabar = True Then
            
                strSql = " BEGIN TRAN "
                clsdoc.Ejecutar strSql, "M"
                strSql = " SELECT MAX(doc_pag_ped_codigo) as Num FROM doc_pago_pedido WITH (TABLOCKX) " & _
                         " WHERE emp_codigo='" & strEmpresa & "'" & _
                         " AND doc_pag_ped_codigo LIKE 'CSP%'"
                clsdoc.Ejecutar strSql, "M"
                maxpago = "CSP" & Format(FormatoD0(Mid(clsdoc.adorec_Def("Num"), 4)) + 1, "0000000000")
            
                ffch = Format(VSFG2.TextMatrix(i, 9), "yyyy-mm-dd")
                Fecha = Format(VSFG2.TextMatrix(i, 16), "yyyy-mm-dd")
                
                fechac = Format(VSFG2.TextMatrix(i, 9), "yyyy-mm-dd") '*********
                fechac = ffch
                ValorPago = FormatoD2(VSFG2.TextMatrix(i, 8))
                ElAsiento = "NULL"
            
                intAnticipo = 0
                If IsEmpty(intPend) Then
                    intPend = 0
                End If
                txtDescripcion.Text = Descripcion & " " & VSFG2.TextMatrix(i, 3) & " PED: " & VSFG2.TextMatrix(i, 4)
                strSql = " INSERT INTO doc_pago_pedido (emp_codigo, doc_pag_ped_tipo,doc_pag_ped_codigo, tip_doc_pag_codigo, " & _
                         " ban_codigo, doc_pag_ped_numero, doc_pag_ped_fecha_recepcion, doc_pag_ped_fecha_doc ," & _
                         " per_codigo, doc_pag_ped_valor, doc_pag_ped_observacion, doc_pag_ped_estado," & _
                         " doc_pag_ped_pendiente,doc_pag_ped_anticipo, doc_pag_ped_fechamod, doc_pag_ped_usumod," & _
                         " per_codigo_ch,ped_codigo)" & _
                         " VALUES ('" & strEmpresa & "','P','" & maxpago & "', '" & dcmbDocumento.BoundText & "', " & _
                         " '" & VSFG2.TextMatrix(i, 14) & "', '" & VSFG2.TextMatrix(i, 10) & "','" & ffch & "', '" & Fecha & "'," & _
                         " '" & VSFG2.TextMatrix(i, 2) & "', '" & VSFG2.TextMatrix(i, 8) & "', '" & UCase(txtDescripcion) & "', 'GIRADO'," & _
                         " '" & intPend & "','" & intAnticipo & "', CURRENT_TIMESTAMP, '" & strUsuario & "'," & _
                         " '" & VSFG2.TextMatrix(i, 31) & "','" & VSFG2.TextMatrix(i, 4) & "') "

                clsdoc.Ejecutar strSql, "M"
                strSql = " COMMIT TRAN "
                clsdoc.Ejecutar strSql, "M"
                
            End If
        
        End If
    Next i
    'LiberarYBajarPedidos
''GUARDAR EN EXCEL
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
    
    If CantidadAnticipo > 0 Then
        Dim rptNuevo() As New frmReporte
        ReDim rptNuevo(CantidadAnticipo)
        For i = 1 To CantidadAnticipo
            rptNuevo(i).strAsiento = NumAsiAnticipo(i)
            rptNuevo(i).strReporte = "rptAsiento"
            rptNuevo(i).Show
        Next i
    End If
'''    '************** INCENTIVOS
'''    'pbGeneral.Value = 1
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "Cargando Incentivos"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optIncentivo.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optPromoCombo.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    'pbGeneral.Value = 2
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "Premio x Monto"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optPremio.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    'pbGeneral.Value = 3
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "Promo Combo Pedido"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optPromoComboPedido.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    'pbGeneral.Value = 4
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "Dcto x Combo"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optDctoCombo.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    'pbGeneral.Value = 5
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "n Prendas a $x.xx"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optNPrendasAY.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
'''    'dtpFechaFinAplicar.Value = dtpEjecutar.Value
'''    'pbGeneral.Value = 6
'''    'pbGeneral.Refresh
'''    lblEstado.Caption = "Premio por monto Marca"
'''    lblEstado.Refresh
'''    frmIncentivos.Show
'''    frmIncentivos.cmbNegocioAplicar.BoundText = "JON"
'''    frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFecha.Value), "yyyy-mm-dd")
'''    frmIncentivos.dtpFechaFinAplicar.Value = dtpFecha.Value
'''    frmIncentivos.optPromoPremioPorMontoMarca.Value = True
'''    frmIncentivos.cmdActualizar_Click
'''    frmIncentivos.cmdAplicarAplicar_Click
''''    frmIncentivos.CmdSalir_Click
''''*****************FIN INCENTIVOS
'''
''''***************** APLICAR NC A PEDIDOS
'''
'''    lblEstado.Caption = "Aplicando NC a PEDIDOS"
'''    lblEstado.Refresh
'''    frmAplicarNC.Show
'''    frmAplicarNC.cmbNegocio.BoundText = "JON"
'''    frmAplicarNC.dtpFechaNCIni.Value = Format(DateAdd("m", -6, dtpFecha.Value), "yyyy-mm-dd")
'''    frmAplicarNC.dtpFechaNCFin.Value = dtpFecha.Value
'''    frmAplicarNC.chkSelTodo.Value = 1
'''    frmAplicarNC.cmdMostrarNC_Click
'''    frmAplicarNC.chkIncluirPedidos.Value = 1
'''    frmAplicarNC.dtpFechaFacIni.Value = Format(DateAdd("m", -3, dtpFecha.Value), "yyyy-mm-dd")
'''    frmAplicarNC.dtpFechaFacFin.Value = dtpFecha.Value
'''    frmAplicarNC.cmdMostrarAplicacion_Click
'''    frmAplicarNC.cmdAplicar_Click
'''    'guardar excel
'''    frmAplicarNC.CmdSalir_Click
'''
''''***************** FIN APLICAR NC A PEDIDOS
'''
''''***************** ENVIO DE SMS PEDIDOS
'''    lblEstado.Caption = "Envio SMS"
'''    lblEstado.Refresh
'''    frmCarteraPedidos.Show
'''    frmCarteraPedidos.cmdActualizar_Click
'''    frmCarteraPedidos.cmdEnvioCorreo_Click
'''    frmCarteraPedidos.cmdcancelar_Click
''''***************** FIN ENVIO DE SMS PEDIDOS
    MsgBox "Carga y Aplicaciones Finalizadas", vbInformation, "Cartera"
    cmdAplicar.Enabled = False
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    Dim respuesta As Integer
    Dim FechaPro As String
    Dim FechaDoc As String
    Dim strNombreBanco As String
    Dim strNegocio As String
    Dim clsSqlBan As New clsConsulta
    Dim TotalArchivo As Double
    Dim TotalAplicar As Double
    Dim TotalAnticipos As Long
    Dim strCodigoDeudor As String
    clsSqlBan.Inicializar AdoConn, AdoConnMaster
    cmdAplicar.Enabled = False
    TotalAnticipos = 0
    sDir = CurDir
    cdArchivo.FileName = ""
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    If cdArchivo.FileName = "" Then Exit Sub
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    If (txtArchivo.Text <> "") Then
        Me.MousePointer = 11
        VSFG.ClipSeparators = ";" & vbCr
        VSFG.FixedRows = 0
        VSFG.Rows = 0
        VSFG2.Rows = 1
        'VSFG.LoadGrid txtArchivo.Text, flexFileCustomText
        'VSFG.LoadGrid txtArchivo.Text, flexFileCommaText
        VSFG.LoadGrid txtArchivo.Text, flexFileTabText
        VSFG.Cols = 11
        VSFG.AddItem "", 0
        VSFG.TextMatrix(0, 0) = "Negocio"
        VSFG.TextMatrix(0, 1) = "CI/RUC"
        VSFG.TextMatrix(0, 2) = "Factura"
        VSFG.TextMatrix(0, 3) = "Cliente"
        VSFG.TextMatrix(0, 4) = "Valor"
        VSFG.TextMatrix(0, 5) = "Fecha proceso"
        VSFG.TextMatrix(0, 6) = "No. transaccion"
        VSFG.TextMatrix(0, 7) = "Banco"
        VSFG.TextMatrix(0, 8) = "Fecha Doc."
        VSFG.TextMatrix(0, 9) = "DeudorCH"
        VSFG.TextMatrix(0, 10) = "Proceso"
        VSFG.FixedRows = 1
        TotalArchivo = 0
        TotalAplicar = 0
        For i = 1 To VSFG.Rows - 1
            errorCarga = False
            VSFG.ShowCell i, 0
            If IsNumeric(VSFG.TextMatrix(i, 4)) = True Then
                TotalArchivo = TotalArchivo + VSFG.TextMatrix(i, 4)
            End If
            If VSFG.TextMatrix(i, 1) <> "" Then
                If VSFG.TextMatrix(i, 0) <> "" Then
                    strNegocio = ""
                    strSql = " SELECT tip_ped_nombre FROM tipo_pedido WHERE tip_ped_codigo='" & VSFG.TextMatrix(i, 0) & "' "
                    clsSqlBan.Ejecutar strSql
                    If clsSqlBan.adorec_Def.RecordCount > 0 Then
                        strNegocio = clsSqlBan.adorec_Def("tip_ped_nombre")
                    Else
                        errorCarga = True
                        VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN TIPO DE NEGOCIO"
                    End If
                Else
                    errorCarga = True
                    VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN TIPO DE NEGOCIO"
                End If
                
                If VSFG.TextMatrix(i, 7) <> "" Then
                    strNombreBanco = ""
                    strSql = " SELECT ban_nombre FROM banco WHERE ban_codigo='" & VSFG.TextMatrix(i, 7) & "' "
                    clsSqlBan.Ejecutar strSql
                    If clsSqlBan.adorec_Def.RecordCount > 0 Then
                        strNombreBanco = clsSqlBan.adorec_Def("ban_nombre")
                    Else
                        errorCarga = True
                        VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN CODIGO DE BANCO"
                    End If
                Else
                    'errorCarga = True
                    'VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN CODIGO DE BANCO"
                End If
                FechaPro = Mid(VSFG.TextMatrix(i, 5), 7, 4) & "/" & Mid(VSFG.TextMatrix(i, 5), 4, 2) & "/" & Mid(VSFG.TextMatrix(i, 5), 1, 2)
                If IsDate(FechaPro) = False Then
                    errorCarga = True
                    VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN FECHA DE PROCESO"
                End If
                If VSFG.TextMatrix(i, 8) <> "" Then
                    FechaDoc = Mid(VSFG.TextMatrix(i, 8), 7, 4) & "/" & Mid(VSFG.TextMatrix(i, 8), 4, 2) & "/" & Mid(VSFG.TextMatrix(i, 8), 1, 2)
                    If IsDate(FechaDoc) = False Then
                        errorCarga = True
                        VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN FECHA DE DOCUMENTO"
                    End If
                Else
                    FechaDoc = FechaPro
                End If
                strCodigoDeudor = ""
                If FechaDoc > FechaPro Then
                    If VSFG.TextMatrix(i, 9) <> "" Then
                        strSql = " SELECT per_codigo FROM persona " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND cat_p_tipo='C' " & _
                                 " AND tip_ped_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                                 " AND per_ruc like '%" & VSFG.TextMatrix(i, 9) & "'"
                        clsSqlBan.Ejecutar strSql
                        If clsSqlBan.adorec_Def.RecordCount > 0 Then
                            strCodigoDeudor = clsSqlBan.adorec_Def("per_codigo")
                        Else
                            errorCarga = True
                            VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN DEUDOR DEL CHEQUE"
                        End If
                    Else
                        VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN DEUDOR DEL CHEQUE"
                    End If
                End If
                If errorCarga = False Then
                    If Left(VSFG.TextMatrix(i, 2), 1) <> "P" And Left(VSFG.TextMatrix(i, 2), 1) <> "9" Then
                        respuesta = RevisarCXCParaAplicar(VSFG.TextMatrix(i, 0), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 2), VSFG.TextMatrix(i, 4), FechaPro, VSFG.TextMatrix(i, 6), VSFG.TextMatrix(i, 7), strNombreBanco, FechaDoc, strNegocio, strCodigoDeudor)
                        If respuesta = 1 Then
                            VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
                            VSFG.TextMatrix(i, 10) = "PROCESADO COMO PAGO"
                            TotalAplicar = TotalAplicar + VSFG.TextMatrix(i, 4)
                        ElseIf respuesta <= 0 Then
                            If Trim(UCase(VSFG.TextMatrix(i, 2))) = "ANTICIPO" Or Left(Trim(UCase(VSFG.TextMatrix(i, 2))), 1) = "P" Then
                                respuesta = RevisarAnticipo(VSFG.TextMatrix(i, 0), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 2), VSFG.TextMatrix(i, 4), FechaPro, VSFG.TextMatrix(i, 6), VSFG.TextMatrix(i, 7), strNombreBanco, FechaDoc, strNegocio, strCodigoDeudor)
                                If respuesta = 1 Then
                                    VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbCyan
                                    VSFG.TextMatrix(i, 10) = "PROCESADO COMO ANTICIPO"
                                    TotalAplicar = TotalAplicar + VSFG.TextMatrix(i, 4)
                                    TotalAnticipos = TotalAnticipos + 1
                                ElseIf respuesta = -1 Then
                                    VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbBlue
                                    VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN DATOS PARA ANTICIPO"
                                End If
                            Else
                                VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbBlue
                                VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN DATOS PARA APLICAR A FACTURA"
                            End If
                        End If
                    Else
                        respuesta = RevisarPedidoParaAplicar(VSFG.TextMatrix(i, 0), VSFG.TextMatrix(i, 1), VSFG.TextMatrix(i, 2), VSFG.TextMatrix(i, 4), FechaPro, VSFG.TextMatrix(i, 6), VSFG.TextMatrix(i, 7), strNombreBanco, FechaDoc, strNegocio, strCodigoDeudor)
                        If respuesta = 1 Then
                            VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
                            VSFG.TextMatrix(i, 10) = "PROCESADO COMO PAGO A PEDIDO"
                            TotalAplicar = TotalAplicar + VSFG.TextMatrix(i, 4)
                        Else
                            VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbBlue
                            VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN DATOS PARA APLICAR A PEDIDO"
                        End If
                    End If
                Else
                    VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbBlue
                End If
                If Not (Left(Trim(VSFG.TextMatrix(i, 10)), 19) = "PROCESADO COMO PAGO" Or Trim(VSFG.TextMatrix(i, 10)) = "PROCESADO COMO ANTICIPO") Then
                    VSFG.TextMatrix(i, 10) = "NO PROCESADO" & VSFG.TextMatrix(i, 10)
                End If
            Else
                VSFG.TextMatrix(i, 10) = VSFG.TextMatrix(i, 10) & " / " & "ERROR EN CLIENTE"
            End If
        Next i
        txtTotalArchivo.Text = TotalArchivo
        txtTotalProcesar.Text = TotalAplicar
        If FormatoD2(txtTotalArchivo.Text) = FormatoD2(txtTotalProcesar.Text) Then
            cmdAplicar.Enabled = True
        End If
        Me.MousePointer = 0
    End If
End Sub
Private Function RevisarAnticipo(strNeg As String, strRUC As String, strFac As String, strValor As String, strFecha As String, strTransac As String, strBanco As String, strNombreBanco As String, strFechaDoc As String, strNegocio As String, strCodigoDeudor As String) As Integer
    Dim clsConsultaCXC As New clsConsulta
    Dim clsConsultaDeudor As New clsConsulta
    Dim strPer_codigo As String
    Dim strCue_p_c_codigo As String
    Dim strPag_codigo As String
    Dim strCtaContable As String
    Dim strNombreCliente As String
    Dim strNumeroPed As String
    Dim strDireccion2 As String
    Dim strFP As String
    Dim strFPI As String
    Dim strN1 As String
    Dim strN2 As String
    Dim strN3 As String
    Dim strN4 As String
    Dim strN5 As String
    Dim strN6 As String
    Dim strN7 As String
    Dim strN8 As String
    Dim strN9 As String
    Dim strDeudor As String
    Dim dblValor As Double
    strNumeroPed = ""
    If strValor <> "" Then
        strValor = Format(strValor, "#0000.00")
        dblValor = FormatoD2(FormatoD0(Left(strValor, Len(strValor) - 2)) & "." & Format(FormatoD0(Right(strValor, 2)), "00"))
    Else
        dblValor = 0
    End If
    
    clsConsultaCXC.Inicializar AdoConn, AdoConnMaster
    clsConsultaDeudor.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT tip_ped_ptofac " & _
             " FROM tipo_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND tip_ped_codigo='" & strNeg & "'"
    clsConsultaCXC.Ejecutar strSql
    strNumeroPed = strSucursal & clsConsultaCXC.adorec_Def("tip_ped_ptofac") & Right(strFac, 7)
    strSql = " SELECT persona.per_codigo, CONCAT(persona.per_apellido, ' ',persona.per_nombre) as cli, " & _
             " 'ANTICIPO' AS ped, " & _
             " persona.per_direccion2,COALESCE(fp.for_pag_nombre,'') as fpp,COALESCE(fpi.for_pag_nombre,'') as fpip, " & _
             " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) AS NN1," & _
             " CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')) AS NN2," & _
             " CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')) AS NN3," & _
             " CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')) AS NN4," & _
             " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) AS NN5," & _
             " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) AS NN6," & _
             " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) AS NN7," & _
             " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) AS NN8," & _
             " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) AS NN9 " & _
             " FROM persona  " & _
             " LEFT JOIN forma_pago fp " & _
             " ON persona.emp_codigo=fp.emp_codigo " & _
             " AND persona.for_pag_codigo=fp.for_pag_codigo " & _
             " LEFT JOIN forma_pago fpi " & _
             " ON persona.emp_codigo=fpi.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=fpi.for_pag_codigo "
    strSql = strSql & " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo " & _
             " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
             " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo " & _
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
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND persona.tip_ped_codigo='" & strNeg & "' AND persona.cat_p_tipo='C' " & _
             " AND persona.per_ruc like '%" & strRUC & "'"
    clsConsultaCXC.Ejecutar strSql
    If clsConsultaCXC.adorec_Def.RecordCount > 0 Then
        strPer_codigo = clsConsultaCXC.adorec_Def("per_codigo")
        strNombreCliente = Replace(clsConsultaCXC.adorec_Def("cli"), vbTab, " ")
        strDireccion2 = Replace(clsConsultaCXC.adorec_Def("per_direccion2"), vbTab, " ")
        strNumeroPed = Replace(clsConsultaCXC.adorec_Def("ped"), vbTab, " ")
        strFP = Replace(clsConsultaCXC.adorec_Def("fpp"), vbTab, " ")
        strFPI = Replace(clsConsultaCXC.adorec_Def("fpip"), vbTab, " ")
        strN1 = Replace(clsConsultaCXC.adorec_Def("NN1"), vbTab, " ")
        strN2 = Replace(clsConsultaCXC.adorec_Def("NN2"), vbTab, " ")
        strN3 = Replace(clsConsultaCXC.adorec_Def("NN3"), vbTab, " ")
        strN4 = Replace(clsConsultaCXC.adorec_Def("NN4"), vbTab, " ")
        strN5 = Replace(clsConsultaCXC.adorec_Def("NN5"), vbTab, " ")
        strN6 = Replace(clsConsultaCXC.adorec_Def("NN6"), vbTab, " ")
        strN7 = Replace(clsConsultaCXC.adorec_Def("NN7"), vbTab, " ")
        strN8 = Replace(clsConsultaCXC.adorec_Def("NN8"), vbTab, " ")
        strN9 = Replace(clsConsultaCXC.adorec_Def("NN9"), vbTab, " ")
        strSql = " SELECT IIF(cat_p_ctaconta_ant IS NULL OR cat_p_ctaconta_ant='',par_texto,cat_p_ctaconta_ant) as par_texto " & _
                 " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                 " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                 " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                 " AND per_codigo='" & strPer_codigo & "' AND persona.cat_p_tipo='C' "
        clsConsultaCXC.Ejecutar strSql
        strCtaContable = clsConsultaCXC.adorec_Def("par_texto")
        
        strDeudor = ""
        If strCodigoDeudor <> "" Then
            strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombre " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cat_p_tipo='C' " & _
                     " AND per_codigo='" & strCodigoDeudor & "'"
            clsConsultaDeudor.Ejecutar strSql
            strDeudor = clsConsultaDeudor.adorec_Def("nombre")
        End If
        VSFG2.AddItem "-" & vbTab & _
                      "-" & vbTab & _
                      strPer_codigo & vbTab & _
                      strNombreCliente & vbTab & _
                      "ANTICIPO" & vbTab & _
                      "ANTICIPO PEDIDO " & strNumeroPed & vbTab & _
                      "0" & vbTab & _
                      "0" & vbTab & _
                      dblValor & vbTab & _
                      strFecha & vbTab & _
                      strTransac & vbTab & _
                      strCtaContable & vbTab & _
                      "" & vbTab & _
                      "" & vbTab & _
                      strBanco & vbTab & _
                      strNombreBanco & vbTab & _
                      strFechaDoc & vbTab & _
                      strNegocio & vbTab & _
                      strNumeroPed & vbTab & strFP & vbTab & strFPI & vbTab & _
                      strN1 & vbTab & strN2 & vbTab & strN3 & vbTab & _
                      strN4 & vbTab & strN5 & vbTab & strN6 & vbTab & _
                      strN7 & vbTab & strN8 & vbTab & strN9 & vbTab & _
                      strDireccion2 & vbTab & strCodigoDeudor & vbTab & strDeudor
        VSFG2.ShowCell VSFG2.Rows - 1, 0
        RevisarAnticipo = 1
    Else
        RevisarAnticipo = 0
    End If

End Function

Private Function RevisarCXCParaAplicar(strNeg As String, strRUC As String, strFac As String, strValor As String, strFecha As String, strTransac As String, strBanco As String, strNombreBanco As String, strFechaDoc As String, strNegocio As String, strCodigoDeudor As String) As Integer
    Dim clsConsultaCXC1 As New clsConsulta
    Dim clsConsultaCXC As New clsConsulta
    Dim clsConsultaDeudor As New clsConsulta
    Dim strPer_codigo As String
    Dim strCue_p_c_codigo As String
    Dim strPag_codigo As String
    Dim strCtaContable As String
    Dim strNombreCliente As String
    Dim strNumeroPed As String
    Dim strDireccion2 As String
    Dim strFP As String
    Dim strFPI As String
    Dim strN1 As String
    Dim strN2 As String
    Dim strN3 As String
    Dim strN4 As String
    Dim strN5 As String
    Dim strN6 As String
    Dim strN7 As String
    Dim strN8 As String
    Dim strN9 As String
    Dim strDeudor As String
    Dim dblValor As Double
    Dim booTermino As Boolean
    
    strFac = Right(strFac, 7)
    If strValor <> "" Then
        strValor = Format(strValor, "#0000.00")
        dblValor = FormatoD2(FormatoD0(Left(strValor, Len(strValor) - 2)) & "." & Format(FormatoD0(Right(strValor, 2)), "00"))
    Else
        dblValor = 0
    End If
    
    clsConsultaCXC1.Inicializar AdoConn, AdoConnMaster
    clsConsultaCXC.Inicializar AdoConn, AdoConnMaster
    clsConsultaDeudor.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT persona.per_codigo,cue_p_c_codigo, CONCAT(persona.per_apellido, ' ',persona.per_nombre) as cli," & _
             " COALESCE(SUBSTRING(cue_p_c_descripcion,CHARINDEX('PED.:',cue_p_c_descripcion),20),'') AS ped, " & _
             " persona.per_direccion2,COALESCE(fp.for_pag_nombre,'') as fpp,COALESCE(fpi.for_pag_nombre,'') as fpip, " & _
             " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) AS NN1," & _
             " CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')) AS NN2," & _
             " CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')) AS NN3," & _
             " CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')) AS NN4," & _
             " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) AS NN5," & _
             " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) AS NN6," & _
             " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) AS NN7," & _
             " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) AS NN8," & _
             " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) AS NN9 " & _
             " FROM persona INNER JOIN cuenta_p_c " & _
             " ON persona.emp_codigo=cuenta_p_c.emp_codigo " & _
             " AND persona.per_codigo=cuenta_p_c.per_codigo and cue_p_c_pagado=0" & _
             " LEFT JOIN forma_pago fp " & _
             " ON persona.emp_codigo=fp.emp_codigo " & _
             " AND persona.for_pag_codigo=fp.for_pag_codigo " & _
             " LEFT JOIN forma_pago fpi " & _
             " ON persona.emp_codigo=fpi.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=fpi.for_pag_codigo "
    strSql = strSql & " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo " & _
             " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
             " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo " & _
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
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND persona.tip_ped_codigo='" & strNeg & "'" & _
             " AND cuenta_p_c.cue_p_c_egr_codigo LIKE '%" & strFac & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo='C'" & _
             " AND persona.cat_p_tipo='C'" & _
             " AND persona.per_ruc like '%" & strRUC & "'"
    clsConsultaCXC1.Ejecutar strSql
    If clsConsultaCXC1.adorec_Def.RecordCount > 0 Then
        booTermino = False
        While booTermino = False And Not clsConsultaCXC1.adorec_Def.EOF
            strPer_codigo = clsConsultaCXC1.adorec_Def("per_codigo")
            strCue_p_c_codigo = clsConsultaCXC1.adorec_Def("cue_p_c_codigo")
            strNombreCliente = Replace(clsConsultaCXC1.adorec_Def("cli"), vbTab, " ")
            strDireccion2 = Replace(clsConsultaCXC1.adorec_Def("per_direccion2"), vbTab, " ")
            strNumeroPed = Replace(clsConsultaCXC1.adorec_Def("ped"), vbTab, " ")
            strFP = Replace(clsConsultaCXC1.adorec_Def("fpp"), vbTab, " ")
            strFPI = Replace(clsConsultaCXC1.adorec_Def("fpip"), vbTab, " ")
            strN1 = Replace(clsConsultaCXC1.adorec_Def("NN1"), vbTab, " ")
            strN2 = Replace(clsConsultaCXC1.adorec_Def("NN2"), vbTab, " ")
            strN3 = Replace(clsConsultaCXC1.adorec_Def("NN3"), vbTab, " ")
            strN4 = Replace(clsConsultaCXC1.adorec_Def("NN4"), vbTab, " ")
            strN5 = Replace(clsConsultaCXC1.adorec_Def("NN5"), vbTab, " ")
            strN6 = Replace(clsConsultaCXC1.adorec_Def("NN6"), vbTab, " ")
            strN7 = Replace(clsConsultaCXC1.adorec_Def("NN7"), vbTab, " ")
            strN8 = Replace(clsConsultaCXC1.adorec_Def("NN8"), vbTab, " ")
            strN9 = Replace(clsConsultaCXC1.adorec_Def("NN9"), vbTab, " ")
            strSql = " SELECT COALESCE(MAX(pag_codigo),1) as pag " & _
                     " FROM pago " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND cue_p_c_codigo = '" & strCue_p_c_codigo & "'" & _
                     " AND cue_p_c_tipo='C'"
            clsConsultaCXC.Ejecutar strSql
            strPag_codigo = clsConsultaCXC.adorec_Def("pag")
            
            strSql = " SELECT IIF(cat_p_ctaconta IS NULL OR cat_p_ctaconta='',par_texto,cat_p_ctaconta) as par_texto " & _
                     " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                     " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                     " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                     " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                     " AND per_codigo='" & strPer_codigo & "' AND persona.cat_p_tipo='C' "
            clsConsultaCXC.Ejecutar strSql
            strCtaContable = clsConsultaCXC.adorec_Def("par_texto")
            strSql = " SELECT '" & strPag_codigo & "' as pag_codigo,cuenta_p_c.cue_p_c_codigo, cuenta_p_c.per_codigo," & _
                     "'" & strNombreCliente & "' as cli, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_valor," & _
                     " cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d, '" & dblValor & "' as e," & _
                     "'" & strFecha & "' as f, '" & strTransac & "' as g,'" & strCtaContable & "' as h,'' as i,'' as j," & _
                     "'" & strBanco & "' as k,'" & strNombreBanco & "' as l , '" & strFechaDoc & "' as m,'" & strNegocio & "' as n " & _
                     " FROM  (cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo)" & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " WHERE per_codigo='" & strPer_codigo & "' AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                     " AND cuenta_p_c.cue_p_c_codigo='" & strCue_p_c_codigo & "' " & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo, cuenta_p_c.cue_p_c_tipo, cuenta_p_c.per_codigo,cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_valor,com_ret_total HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)!=0 " & _
                     " ORDER BY cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
            clsConsultaCXC.Ejecutar strSql
            If clsConsultaCXC.adorec_Def.RecordCount > 0 Then
                If FormatoD2(dblValor) <= FormatoD2(clsConsultaCXC.adorec_Def(7)) Then
                    strDeudor = ""
                    If strCodigoDeudor <> "" Then
                        strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombre " & _
                                 " FROM persona " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND cat_p_tipo='C' " & _
                                 " AND per_codigo='" & strCodigoDeudor & "'"
                        clsConsultaDeudor.Ejecutar strSql
                        strDeudor = clsConsultaDeudor.adorec_Def("nombre")
                    End If
                    VSFG2.AddItem clsConsultaCXC.adorec_Def(0) & vbTab & _
                                  clsConsultaCXC.adorec_Def(1) & vbTab & _
                                  clsConsultaCXC.adorec_Def(2) & vbTab & _
                                  clsConsultaCXC.adorec_Def(3) & vbTab & _
                                  clsConsultaCXC.adorec_Def(4) & vbTab & _
                                  clsConsultaCXC.adorec_Def(5) & vbTab & _
                                  clsConsultaCXC.adorec_Def(6) & vbTab & _
                                  clsConsultaCXC.adorec_Def(7) & vbTab & _
                                  clsConsultaCXC.adorec_Def(8) & vbTab & _
                                  clsConsultaCXC.adorec_Def(9) & vbTab & _
                                  clsConsultaCXC.adorec_Def(10) & vbTab & _
                                  clsConsultaCXC.adorec_Def(11) & vbTab & _
                                  clsConsultaCXC.adorec_Def(12) & vbTab & _
                                  clsConsultaCXC.adorec_Def(13) & vbTab & _
                                  clsConsultaCXC.adorec_Def(14) & vbTab & _
                                  clsConsultaCXC.adorec_Def(15) & vbTab & _
                                  clsConsultaCXC.adorec_Def(16) & vbTab & _
                                  clsConsultaCXC.adorec_Def(17) & vbTab & _
                                  strNumeroPed & vbTab & strFP & vbTab & strFPI & vbTab & _
                                  strN1 & vbTab & strN2 & vbTab & strN3 & vbTab & _
                                  strN4 & vbTab & strN5 & vbTab & strN6 & vbTab & _
                                  strN7 & vbTab & strN8 & vbTab & strN9 & vbTab & _
                                  strDireccion2 & vbTab & strCodigoDeudor & vbTab & strDeudor
                    VSFG2.ShowCell VSFG2.Rows - 1, 0
                    RevisarCXCParaAplicar = 1
                    booTermino = True
                Else
                    RevisarCXCParaAplicar = 0
                    booTermino = False
                End If
            Else
'                strSQL = " SELECT * FROM pago WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_tipo='C' and cue_p_c_codigo='" & strCue_p_c_codigo & "' AND pag_fecha='" & Format(strFecha, "yyyy-mm-dd") & "' AND pag_no_doc='" & strTransac & "'"
'                clsConsultaCXC.Ejecutar strSQL
                RevisarCXCParaAplicar = -1
                booTermino = False
            End If
            clsConsultaCXC1.adorec_Def.MoveNext
        Wend
        
    Else
        RevisarCXCParaAplicar = 0
    End If
End Function

Private Function RevisarPedidoParaAplicar(strNeg As String, strRUC As String, strPed As String, strValor As String, strFecha As String, strTransac As String, strBanco As String, strNombreBanco As String, strFechaDoc As String, strNegocio As String, strCodigoDeudor As String) As Integer
    Dim clsConsultaCXC As New clsConsulta
    Dim clsConsultaDeudor As New clsConsulta
    Dim strPer_codigo As String
    Dim strPed_codigo As String
    Dim strPag_codigo As String
    Dim strCtaContable As String
    Dim strNombreCliente As String
    Dim strNumeroPed As String
    Dim strDireccion2 As String
    Dim strFP As String
    Dim strFPI As String
    Dim strN1 As String
    Dim strN2 As String
    Dim strN3 As String
    Dim strN4 As String
    Dim strN5 As String
    Dim strN6 As String
    Dim strN7 As String
    Dim strN8 As String
    Dim strN9 As String
    Dim strDeudor As String
    Dim dblValor As Double
    
    strPed = Right(strPed, 7)
    If strValor <> "" Then
        strValor = Format(strValor, "#0000.00")
        dblValor = FormatoD2(FormatoD0(Left(strValor, Len(strValor) - 2)) & "." & Format(FormatoD0(Right(strValor, 2)), "00"))
    Else
        dblValor = 0
    End If
    
    clsConsultaCXC.Inicializar AdoConn, AdoConnMaster
    clsConsultaDeudor.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT persona.per_codigo,ped_codigo, CONCAT(persona.per_apellido, ' ',persona.per_nombre) as cli," & _
             " ped_codigo AS ped, " & _
             " persona.per_direccion2,COALESCE(fp.for_pag_nombre,'') as fpp,COALESCE(fpi.for_pag_nombre,'') as fpip, " & _
             " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) AS NN1," & _
             " CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')) AS NN2," & _
             " CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')) AS NN3," & _
             " CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')) AS NN4," & _
             " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) AS NN5," & _
             " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) AS NN6," & _
             " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) AS NN7," & _
             " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) AS NN8," & _
             " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) AS NN9 " & _
             " FROM persona INNER JOIN pedido " & _
             " ON persona.emp_codigo=pedido.emp_codigo " & _
             " AND persona.per_codigo=pedido.per_codigo and ped_estado=0" & _
             " LEFT JOIN forma_pago fp " & _
             " ON persona.emp_codigo=fp.emp_codigo " & _
             " AND persona.for_pag_codigo=fp.for_pag_codigo " & _
             " LEFT JOIN forma_pago fpi " & _
             " ON persona.emp_codigo=fpi.emp_codigo " & _
             " AND persona.for_pag_codigo_imp=fpi.for_pag_codigo "
    strSql = strSql & " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo " & _
             " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
             " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo " & _
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
             " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
             " AND persona.tip_ped_codigo='" & strNeg & "'" & _
             " AND pedido.ped_codigo LIKE '%" & strPed & "'" & _
             " AND persona.cat_p_tipo='C'" & _
             " AND persona.per_ruc like '%" & strRUC & "'"
    clsConsultaCXC.Ejecutar strSql
    If clsConsultaCXC.adorec_Def.RecordCount > 0 Then
        strPer_codigo = clsConsultaCXC.adorec_Def("per_codigo")
        strPed_codigo = clsConsultaCXC.adorec_Def("ped_codigo")
        strNombreCliente = Replace(clsConsultaCXC.adorec_Def("cli"), vbTab, " ")
        strDireccion2 = Replace(clsConsultaCXC.adorec_Def("per_direccion2"), vbTab, " ")
        strNumeroPed = Replace(clsConsultaCXC.adorec_Def("ped"), vbTab, " ")
        strFP = Replace(clsConsultaCXC.adorec_Def("fpp"), vbTab, " ")
        strFPI = Replace(clsConsultaCXC.adorec_Def("fpip"), vbTab, " ")
        strN1 = Replace(clsConsultaCXC.adorec_Def("NN1"), vbTab, " ")
        strN2 = Replace(clsConsultaCXC.adorec_Def("NN2"), vbTab, " ")
        strN3 = Replace(clsConsultaCXC.adorec_Def("NN3"), vbTab, " ")
        strN4 = Replace(clsConsultaCXC.adorec_Def("NN4"), vbTab, " ")
        strN5 = Replace(clsConsultaCXC.adorec_Def("NN5"), vbTab, " ")
        strN6 = Replace(clsConsultaCXC.adorec_Def("NN6"), vbTab, " ")
        strN7 = Replace(clsConsultaCXC.adorec_Def("NN7"), vbTab, " ")
        strN8 = Replace(clsConsultaCXC.adorec_Def("NN8"), vbTab, " ")
        strN9 = Replace(clsConsultaCXC.adorec_Def("NN9"), vbTab, " ")
        strPag_codigo = "P"
        strCtaContable = "P"
        'ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(COALESCE(prd_pro_porcentaje,0)>COALESCE(per_dcto,0),COALESCE(prd_pro_porcentaje,0),COALESCE(per_dcto,0))/100,2))),2) + ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(COALESCE(prd_pro_porcentaje,0)>COALESCE(per_dcto,0),COALESCE(prd_pro_porcentaje,0),COALESCE(per_dcto,0))/100,2))) * (par_numero)/100,2)
        strSql = " SELECT '" & strPag_codigo & "' as pag_codigo,pedido.ped_codigo, pedido.per_codigo," & _
                 "'" & strNombreCliente & "' as cli, pedido.ped_codigo, CONCAT('PEDIDO ',pedido.ped_codigo) as descr, " & _
                 " SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
                 " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
                 " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
                 " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) as c ," & _
                 " ROUND(" & _
                 " SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
                 " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
                 " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
                 " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) " & _
                 " - COALESCE(doc_pag_valor,0.00),2) as d," & _
                 " '" & dblValor & "' as e," & _
                 " '" & strFecha & "' as f, '" & strTransac & "' as g,'" & strCtaContable & "' as h,'' as i,'' as j," & _
                 " '" & strBanco & "' as k,'" & strNombreBanco & "' as l , '" & strFechaDoc & "' as m,'" & strNegocio & "' as n "
        strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                 " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & strNeg & "'" & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                 " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                 " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                 " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                 " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                 " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                 " FROM doc_pago_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                 " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                 " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                 " AND pedido.per_codigo=pag.per_codigo "
        strSql = strSql & " WHERE pedido.per_codigo='" & strPer_codigo & "' AND pedido.emp_codigo = '" & strEmpresa & "' " & _
                 " AND ped_estado=0 AND det_pedido.det_ped_incentivo=0 " & _
                 " AND pedido.ped_codigo='" & strPed_codigo & "' " & _
                 " GROUP BY pedido.ped_codigo, pedido.per_codigo,par_numero,doc_pag_valor " & _
                 " HAVING round(ROUND(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2))),2) " & _
                 "+ ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2))) * (par_numero)/100.00,2) - COALESCE(doc_pag_valor,0.00),2),2)!=0 " & _
                 " ORDER BY pedido.ped_codigo "
        clsConsultaCXC.Ejecutar strSql
        If clsConsultaCXC.adorec_Def.RecordCount > 0 Then
            If FormatoD2(dblValor) <= FormatoD2(clsConsultaCXC.adorec_Def(7)) Then
                strDeudor = ""
                If strCodigoDeudor <> "" Then
                    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre,' (',per_ruc,')') as nombre " & _
                             " FROM persona " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND cat_p_tipo='C' " & _
                             " AND per_codigo='" & strCodigoDeudor & "'"
                    clsConsultaDeudor.Ejecutar strSql
                    strDeudor = clsConsultaDeudor.adorec_Def("nombre")
                End If
                VSFG2.AddItem clsConsultaCXC.adorec_Def(0) & vbTab & _
                              clsConsultaCXC.adorec_Def(1) & vbTab & _
                              clsConsultaCXC.adorec_Def(2) & vbTab & _
                              clsConsultaCXC.adorec_Def(3) & vbTab & _
                              clsConsultaCXC.adorec_Def(4) & vbTab & _
                              clsConsultaCXC.adorec_Def(5) & vbTab & _
                              clsConsultaCXC.adorec_Def(6) & vbTab & _
                              clsConsultaCXC.adorec_Def(7) & vbTab & _
                              clsConsultaCXC.adorec_Def(8) & vbTab & _
                              clsConsultaCXC.adorec_Def(9) & vbTab & _
                              clsConsultaCXC.adorec_Def(10) & vbTab & _
                              clsConsultaCXC.adorec_Def(11) & vbTab & _
                              clsConsultaCXC.adorec_Def(12) & vbTab & _
                              clsConsultaCXC.adorec_Def(13) & vbTab & _
                              clsConsultaCXC.adorec_Def(14) & vbTab & _
                              clsConsultaCXC.adorec_Def(15) & vbTab & _
                              clsConsultaCXC.adorec_Def(16) & vbTab & _
                              clsConsultaCXC.adorec_Def(17) & vbTab & _
                              strNumeroPed & vbTab & strFP & vbTab & strFPI & vbTab & _
                              strN1 & vbTab & strN2 & vbTab & strN3 & vbTab & _
                              strN4 & vbTab & strN5 & vbTab & strN6 & vbTab & _
                              strN7 & vbTab & strN8 & vbTab & strN9 & vbTab & _
                              strDireccion2 & vbTab & strCodigoDeudor & vbTab & strDeudor
                VSFG2.ShowCell VSFG2.Rows - 1, 0
                RevisarPedidoParaAplicar = 1
            Else
                RevisarPedidoParaAplicar = 0
            End If
        Else
            strSql = " SELECT * FROM pago WHERE emp_codigo='" & strEmpresa & "' AND cue_p_c_tipo='C' and cue_p_c_codigo='" & strCue_p_c_codigo & "' AND pag_fecha='" & Format(strFecha, "yyyy-mm-dd") & "' AND pag_no_doc='" & strTransac & "'"
            clsConsultaCXC.Ejecutar strSql
            RevisarPedidoParaAplicar = -1
        End If
        
    Else
        RevisarPedidoParaAplicar = 0
    End If
End Function


Private Sub cmdImprimirAnticipos_Click()
    Dim strListaAsiento As String
    strSql = " SELECT asi_numasiento " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND doc_pag_codigo like 'CSP%-A' " & _
             " AND LEFT(doc_pag_fecha_recepcion,10)=LEFT('" & dtpFecha.Value & "',10)"
    clsCon_Def.Ejecutar strSql
    strListaAsiento = ""
    While Not clsCon_Def.adorec_Def.EOF
        strListaAsiento = strListaAsiento & ",'" & clsCon_Def.adorec_Def("asi_numasiento") & "'"
        clsCon_Def.adorec_Def.MoveNext
    Wend
    If Len(strListaAsiento) > 2 Then
        strListaAsiento = Right(strListaAsiento, Len(strListaAsiento) - 2)
        strListaAsiento = Left(strListaAsiento, Len(strListaAsiento) - 1)
    End If
    frmReporte.strAsiento = strListaAsiento
    frmReporte.strReporte = "rptAsiento"
    frmReporte.Show
End Sub

Private Sub Command1_Click()
    frmCarteraPedidos.Show
End Sub

Private Sub dcmbBancoE_Change()
    dcmbCuenta = ""
    'clsCon_Def.Inicializar AdoConn, AdoConnMaster
    If dcmbBancoE.Text = "" Then
        dcmbCuenta.Text = ""
        Exit Sub
    Else
        strSql = " SELECT cta_ban_numero, cta_ban_ctaconta, ban_codigo " & _
                 " FROM cta_banco " & _
                 " WHERE ban_codigo = '" & dcmbBancoE.BoundText & "' AND emp_codigo = '" & strEmpresa & "'"
        clsCon_Def.Ejecutar strSql
        
        If clsCon_Def.adorec_Def.EOF = False Then
            Set dcmbCuenta.RowSource = clsCon_Def.adorec_Def.DataSource
            dcmbCuenta.ListField = "cta_ban_numero"
            dcmbCuenta.BoundColumn = "cta_ban_ctaconta"
            'dcmbCuenta.Tag = clsCta.adorec_Def("cta_ban_ctaconta")
            'dcmbCuenta.Text = clsCta.adorec_Def("cta_ban_numero")
            
        Else
            Set dcmbCuenta.RowSource = Nothing
        End If
        
    End If
End Sub

Private Sub dcmbDocumento_Change()
    Dim TipoDocPago As String
    Dim clsSqlA As New clsConsulta
    clsSqlA.Inicializar AdoConn, AdoConnMaster
    'If optcheque.value = True Then
        strSql = " SELECT ban_codigo,cta_ban_numero " & _
                 " FROM tipo_doc_pago " & _
                 " WHERE tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' "
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

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    dtpFecha.Value = HoyDia
    dtpFechaConta.Value = HoyDia
    
    Set ucrtVSFG.VSFGControl = VSFG2
    ucrtVSFG.Inicializar False, False, False
    Set uctrVSFG1.VSFGControl = VSFG
    uctrVSFG1.Inicializar False, False, False
    
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        
    strSql = " SELECT tip_doc_pag_codigo, tip_doc_pag_nombre " & _
             " FROM tipo_doc_pago "
    clsCon_Def.Ejecutar strSql
    
    Set dcmbDocumento.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbDocumento.ListField = "tip_doc_pag_nombre"
    dcmbDocumento.BoundColumn = "tip_doc_pag_codigo"
    
    
    strSql = " SELECT tip_not_codigo, tip_not_nombre, CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
             " FROM tipo_nota " & _
             " WHERE tip_not_d_c = 'C'" & _
             " ORDER BY tip_not_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.EOF = False Then
        Set dcmbTipo.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbTipo.ListField = "tip_not_nombre"
        dcmbTipo.BoundColumn = "tip_not_codigo"
    Else
        dcmbTipo = ""
    End If
    
    strSql = " SELECT banco.ban_codigo, ban_nombre " & _
             " FROM banco INNER JOIN cta_banco ON cta_banco.ban_codigo=banco.ban_codigo" & _
             " WHERE cta_banco.emp_codigo='" & strEmpresa & "'" & _
             " GROUP BY banco.ban_codigo, ban_nombre ORDER BY ban_codigo"
    clsCon_Def.Ejecutar strSql

    If clsCon_Def.adorec_Def.EOF = False Then
        Set dcmbBancoE.RowSource = clsCon_Def.adorec_Def.DataSource
        dcmbBancoE.ListField = "ban_nombre"
        dcmbBancoE.BoundColumn = "ban_codigo"
    Else
        dcmbBancoE = ""
    End If
        
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

Private Sub dcmbTipo_Change()
 If dcmbTipo = "" Then
    txtDescripciont = ""
    Exit Sub
 End If
strSql = " SELECT CONCAT(SUBSTRING(tip_not_descripcion,1,50),'...') as descripcion " & _
         " FROM tipo_nota " & _
         " WHERE tip_not_d_c = 'C' AND  tip_not_codigo = '" & dcmbTipo.BoundText & "' "
clsCon_Def.Ejecutar strSql

If clsCon_Def.adorec_Def.EOF = False Then
    txtDescripciont = clsCon_Def.adorec_Def("descripcion")
Else
    txtDescripciont.Text = ""
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
