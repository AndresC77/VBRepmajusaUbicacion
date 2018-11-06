VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmV_VerPedBodFac 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Pedidos Facturados"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmV_VerPedBodFac.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10785
   Begin VB.CheckBox chkCIRUC 
      BackColor       =   &H00DDDDDD&
      Caption         =   "CI/RUC"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   795
      Width           =   1215
   End
   Begin VB.OptionButton optNoPedido 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Por Pedido"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   855
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optListaPedido 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Por Listado de Pedidos"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   855
      Width           =   2655
   End
   Begin VB.TextBox txtPedido 
      Height          =   285
      Left            =   8160
      TabIndex        =   22
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdImprimirEnBloque 
      Caption         =   "Imprimir en Bloque"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGSave 
      Height          =   495
      Left            =   8640
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _cx             =   82513449
      _cy             =   82510697
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
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Tipo de Negocio:"
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
      TabIndex        =   19
      Top             =   0
      Width           =   6015
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   255
         Width           =   4455
         _ExtentX        =   7858
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
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Factura"
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frmDet 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Detalle de Pedido:"
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
      Height          =   4935
      Left            =   105
      TabIndex        =   12
      Top             =   2400
      Width           =   10560
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtCantPed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   18
         Top             =   4440
         Width           =   930
      End
      Begin VB.TextBox txtCantEnt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   4440
         Width           =   930
      End
      Begin VB.TextBox txtLector 
         Height          =   285
         Left            =   7800
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3615
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   10155
         _cx             =   82527736
         _cy             =   82516200
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmV_VerPedBodFac.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   7080
         TabIndex        =   16
         Top             =   435
         Width           =   555
      End
      Begin VB.Label LblDetalle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle del Pedido Nº"
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
         Left            =   330
         TabIndex        =   15
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   60
      End
   End
   Begin VB.Frame frmPed 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Listado de Pedidos:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   10575
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "ACT"
         Height          =   855
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGPeds 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10020
         _cx             =   82527498
         _cy             =   82511332
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
         FormatString    =   $"frmV_VerPedBodFac.frx":03FA
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confirmar Pedido"
      Height          =   375
      Left            =   165
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Detalle"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Timer TmrAct 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   360
      Top             =   2520
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   8160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido a Buscar:"
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
      TabIndex        =   26
      Top             =   870
      Width           =   1245
   End
End
Attribute VB_Name = "frmV_VerPedBodFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para confirmar el stock en bodega de un pedido ya realizado con ante_ #
'#  rioridad.                                                                   #
'#  frmV_VerPedBod V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Opciones que permite:                                                       #
'#  *   En una lista se despliegan los pedidos y sus detalles para que la       #
'#      persona encargada pueda ver en mayor detalle el mismo y así poder       #
'#      confirmar la cantidad de los productos que se está pidiendo.            #
'#                                                                              #
'#  Procesos internos que maneja:                                               #
'#  *   La lista que muestra los distintos pedidos se refresca automáticamente  #
'#      cada 20 segundos para buscar un nuevo pedido generado.                  #
'#  *   Al dar un click en la lista de pedidos, automáticamente se cargan los   #
'#      detalles del mismo en un segundo grid.                                  #
'#  *   Se controla que el usuario pida como máximo la cantidad de productos    #
'#      solicitada a la bodega.                                                 #
'#  *   Una vez que el pedido ha sido confirmado su estado pasa a revisado.     #
'#  *   Se pueden ver solo los pedidos que aún no están revisados o los que ya  #
'#      se han revisado el día de hoy.                                          #
'#                                                                              #
'#  Tablas que maneja:                                                          #
'#                                                                              #
'#  persona:                                                                    #
'#  *   De esta tabla se extrae los datos del cliente al que se le adjudica el  #
'#      pedido que se está confirmando.                                         #
'#  *   También se extrae el nombre del vendedor asignado al pedido.            #
'#  pedido:                                                                     #
'#  *   Aquí se actualizan los datos de la cabecera de un pedido.               #
'#  det_pedido:                                                                 #
'#  *   Aquí se actualizan los datos de la cantidad confirmada a entregar.      #
'#                                                                              #
'################################################################################

Private clsPedidosV As New clsConsulta
Private clsPed As New clsConsulta
Private clsExiPrd As New clsConsulta
Private clsSql As New clsConsulta
Private clsPedidos As New clsConsulta
Private intDato As Variant
Private codCot As Double, numPeds As Long
Private banTm As Integer
Private strFac As String

Private Sub chkCIRUC_Click()
    cmbNegocio_Change
End Sub

Private Sub cmbNegocio_Change()
    Dim strCli As String
    cmdLimpiar_Click
    If cmbNegocio.BoundText <> "" Then
        
        strSql = " SELECT tip_ped_ptofac " & _
                 " FROM tipo_pedido " & _
                 " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
        clsSql.Ejecutar strSql
        If chkCIRUC.Value = 1 Then
            strCli = "CONCAT(per_ruc,' ',per_apellido,' ',per_nombre)"
        Else
            strCli = "CONCAT(per_apellido,' ',per_nombre)"
        End If
        If Me.optListaPedido.Value = True Then
            If clsSql.adorec_Def.RecordCount > 0 Then
                banTm = 0
                'Consulta todos los pedidos que pasan a bodega para ser revisados
                strSql = " SELECT RIGHT(ped_codigo,7)+0 as c,ped_fecha,ped_observacion," & strCli & " as nombC,ped_estado,cot_codigo,ped_estado,pedido.ven_codigo,ped_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
                         " FROM pedido INNER JOIN persona ON pedido.emp_codigo = persona.emp_codigo AND pedido.per_codigo = persona.per_codigo " & _
                         " Where pedido.emp_codigo='" & strEmpresa & "' AND cat_p_tipo='C' AND ped_estado in (8) " & _
                         "  " & _
                         " AND pedido.ped_codigo LIKE CONCAT('" & strSucursal & clsSql.adorec_Def(0) & "'+0,'%') " & _
                         " ORDER BY pedido.ped_estado,pedido.ped_codigo "
                clsPedidosV.Ejecutar (strSql)
                'Muestra los datos de los distintos pedidos en un listado
                Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
                
                strSql = " SELECT est_codigo,est_descripcion " & _
                         " FROM est_pedido " & _
                         " ORDER BY est_codigo"
                clsSql.Ejecutar strSql
                VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
            End If
        End If
    Else
        Exit Sub
    End If
    banTm = banTm + 1
    If banTm = 2 Then
        'clsExiPrd.Actualizar
        banTm = 0
    End If
    numPeds = VSFGPeds.Rows
End Sub


Private Sub cmdActualizar_Click()
''''    Dim i As Long
''''    clsPedidosV.Actualizar
''''    banTm = banTm + 1
''''    If banTm = 2 Then
''''        clsExiPrd.Actualizar
''''        banTm = 0
''''    End If
''''    'Muestra los datos de los distintos pedidos en un listado
''''    Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
''''    For i = 1 To VSFGPeds.Rows - 1
''''        If IsNumeric(LblPedido.Caption) Then
''''            If Val(LblPedido.Caption) = Val(VSFGPeds.TextMatrix(i, 0)) Then
''''                VSFGPeds.Row = i
''''                i = VSFGPeds.Rows
''''            End If
''''        End If
''''    Next i
    
    clsPedidosV.Actualizar
    banTm = banTm + 1
    If banTm = 2 Then
        'clsExiPrd.Actualizar
        banTm = 0
    End If
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
    
    strSql = " SELECT est_codigo,est_descripcion " & _
             " FROM est_pedido " & _
             " ORDER BY est_codigo"
    clsSql.Ejecutar strSql
    VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
    
    If VSFGPeds.Rows > numPeds Then
        'MsgBox "Nuevo pedido llegado...", vbInformation, "Pedido"
        Me.SetFocus
    End If
    numPeds = VSFGPeds.Rows
End Sub

Private Sub cmdExportar_Click()
    Dim sDir As String
    If LblPedido.Caption = "" Then
        MsgBox "Primero seleccione un Pedido", vbInformation, "Pedido"
        Exit Sub
    End If
    sDir = CurDir
    cdArchivo.ShowSave
    'cdArchivo.FileName
    If cdArchivo.FileName <> "" Then
        GuardarArchivo cdArchivo.FileName
    End If
    ChDir sDir
End Sub

Private Sub GuardarArchivo(strPath)
    Dim i As Long
    VSFGSave.Clear 1
    VSFGSave.Rows = 0
    For i = 1 To VSFG.Rows - 1
        VSFGSave.AddItem VSFG.TextMatrix(i, 1) & vbTab & VSFG.TextMatrix(i, 3) & vbTab & VSFGPeds.TextMatrix(VSFGPeds.Row, 3)
    Next i
    VSFGSave.SaveGrid strPath, flexFileTabText
End Sub

Private Sub cmdGuardar_Click()
    For i = 1 To VSFG.Rows - 1
        If LblPedido <> "-" Then
            'Actualiza un detalle de pedido
            strSql = " UPDATE det_pedido SET " & _
                     " det_ped_cant_confirmada=" & VSFG.TextMatrix(i, 4) & _
                     " ,det_ped_descripcion='" & VSFG.TextMatrix(i, 6) & _
                     "' WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' AND dep_codigo='" & VSFG.TextMatrix(i, 0) & "'"
            clsSql.Ejecutar (strSql), "M"
        End If
    Next i
    MsgBox "Detalle de Pedido guardado", vbInformation, "Guardar"
End Sub

Private Sub cmdImprimirEnBloque_Click()
    frmImprimirEnBloque.Show
    Unload Me
End Sub

Private Sub optListaPedido_Click()
    'TmrAct.Enabled = True
    cmdLimpiar_Click
    txtPedido.Enabled = False
    VSFGPeds.Height = 2055
    VSFGPeds.Rows = 1
    cmdActualizar.Height = 2055
    frmPed.Height = 2415
    
    frmDet.Top = 3840
    frmDet.Height = 3615
    VSFG.Height = 2175
    cmdGuardar.Top = 3000
    txtCantPed.Top = 3000
    txtCantEnt.Top = 3000
    cmbNegocio_Change
End Sub

Private Sub optNoPedido_Click()
    'TmrAct.Enabled = False
    cmdLimpiar_Click
    txtPedido.Enabled = True
    VSFGPeds.Height = 855
    VSFGPeds.Rows = 1
    cmdActualizar.Height = 855
    frmPed.Height = 1215
        
    frmDet.Top = 2400
    frmDet.Height = 4935
    VSFG.Height = 3615
    cmdGuardar.Top = 4440
    txtCantPed.Top = 4440
    txtCantEnt.Top = 4440
End Sub

Private Sub txtLector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AgregarProd UCase(txtLector.Text)
        txtLector.Text = ""
        CalculaCant
    End If
End Sub

'Función que verifica si es necesario realizar un backOrder del pedido
Private Function verifBackOr() As Integer
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, 3)) <> Val(VSFG.TextMatrix(i, 4)) Then
            verifBackOr = 1
            Exit For
        End If
    Next i
End Function


'Función que genera un backOrder de un pedido
Private Sub backOrder(codPed As Double, codEmp As String)
    Dim clsBack As New clsConsulta
    clsBack.Inicializar AdoConn, AdoConnMaster
    'Recupera el código con el cual se debe generar un nuevo backOrder
    strSql = " Select COALESCE(max(bac_codigo),0) as num " & _
             " From backorder " & _
             " Where emp_codigo='" & codEmp & "'" & _
             " GROUP BY emp_codigo"
    clsBack.Ejecutar (strSql)
    Dim codBac As Double
    codBac = clsBack.adorec_Def("num") + 1
    'Inserta la cabecera del backOrder
    strSql = " INSERT INTO backorder " & _
             " SELECT " & codBac & " AS bac_codigo, emp_codigo, ped_codigo, CURRENT_TIMESTAMP AS bac_fecha, " & _
             " 0 AS bac_baja, CURRENT_TIMESTAMP AS bac_fechamod, '" & strUsuario & "' AS bac_usumod " & _
             " From pedido " & _
             " WHERE ped_codigo=" & codPed & " AND emp_codigo='" & codEmp & "' "
    clsBack.Ejecutar (strSql), "M"
    'Inserta los detalles del backOrder
    strSql = " INSERT INTO det_backorder " & _
             " SELECT emp_codigo, prd_codigo, " & codBac & " AS bac_codigo, " & _
             " det_ped_cant_pedida-det_ped_cant_entregada AS det_bac_cantidad, " & _
             " det_ped_precio, CURRENT_TIMESTAMP AS det_bac_fechamod, " & _
             " '" & strUsuario & "' AS det_bac_usumod " & _
             " From det_pedido " & _
             " WHERE emp_codigo='" & codEmp & "' " & _
             " AND det_ped_cant_pedida > det_ped_cant_entregada " & _
             " AND ped_codigo= " & codPed & _
             " Order by prd_codigo "
    clsBack.Ejecutar (strSql), "M"
    Set clsBack = Nothing
End Sub


Private Sub cmdImprimir_Click()
    Dim RepPed As New frmReporte
    RepPed.strNumero = LblPedido.Caption
    RepPed.strTipo = 2
    RepPed.strReporte = "rptPedido"
    RepPed.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsPedidosV = Nothing
    Set clsPed = Nothing
    Set clsExiPrd = Nothing
    Set clsSql = Nothing
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
    
    Dim ban As Integer, Ban2 As Integer
    'TmrAct.Enabled = False
    'Actualiza la tabla de detalle de pedido con las cantidades a entregar
    
    If LblPedido <> "-" Then
        If FormatoD4(txtCantPed.Text) <> FormatoD4(txtCantEnt.Text) Then
            If MsgBox("La cantidad pedida es DIFERENTE a la confirmada" & _
                vbNewLine & "Desea confirmar el pedido?", vbYesNo + vbQuestion, "Confirmar Pedido") = vbNo Then
                booGuardar = False
                'TmrAct.Enabled = True
                Exit Sub
            End If
        End If
    Else
        MsgBox "Seleccione un pedido", vbInformation, "Pedido"
        'TmrAct.Enabled = True
        Exit Sub
    End If
    
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, 4)) > 0 And LblPedido <> "-" Then
            'Actualiza un detalle de pedido
            strSql = " UPDATE det_pedido SET " & _
                     " det_ped_cant_confirmada=" & VSFG.TextMatrix(i, 4) & _
                     " ,det_ped_descripcion='" & VSFG.TextMatrix(i, 6) & _
                     "' WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido & _
                     " AND prd_codigo='" & VSFG.TextMatrix(i, 1) & "' AND dep_codigo='" & VSFG.TextMatrix(i, 0) & "'"
            clsSql.Ejecutar (strSql), "M"
            ban = 1
        End If
        'Verifica si se puede completar con todo el pedido
        If Val(VSFG.TextMatrix(i, 3)) <> Val(VSFG.TextMatrix(i, 4)) And codCot > 0 Then
            Ban2 = 1
        End If
    Next i
    'Verifica que se haya hecho por lo menos una modificación al pedido
    If ban = 0 Then
        MsgBox "No se ha modificado nada del pedido.", vbInformation, "Pedido"
    Else
        'Verifica si se pudo completar con la cotización si es el caso
        Dim tFac As Integer
        If Ban2 = 0 And codCot > 0 Then
            tFac = 0
        Else
            tFac = 1
        End If
        'Actualiza el estado del pedido una vez que ha sido modificado
        strSql = " UPDATE pedido SET ped_estado=2, " & _
                 " ped_usumod='" & strUsuario & "', " & _
                 " ped_fechamod=CURRENT_TIMESTAMP,tipo_fac_codigo= " & tFac & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido & " "
        clsSql.Ejecutar (strSql), "M"
        
        frmImpresionDirecta.strNumero = strFac
        frmImpresionDirecta.strReporte = "rptFacturaSola"
        frmImpresionDirecta.Show
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.cmdImprimir_Click
        frmImpresionDirecta.CmdCerrar_Click
        frmImpresionDirecta.strNumero = LblPedido
        frmImpresionDirecta.strReporte = "rptSTKDespacho"
        frmImpresionDirecta.optImpresora.Value = True
        frmImpresionDirecta.cmdImprimir_Click
        
        frmImpresionDirecta.CmdCerrar_Click
        
        CmdLimpiar = True
    End If
    'TmrAct.Enabled = True
End Sub

Private Sub cmdLimpiar_Click()
    'Limpia el contenido del grid de detalles
    VSFG.Clear 1
    VSFG.Rows = 2
    LblPedido = "-"
    txtCantPed.Text = ""
    txtCantEnt.Text = ""
    txtLector.Text = ""
    strFac = ""
End Sub

Private Sub AgregarProd(codigo As String, Optional EsAux As Boolean = True)
  Dim i As Long
  Dim pas As Boolean
  pas = False
  With VSFG
    For i = 1 To .Rows - 1
      If codigo = .TextMatrix(i, 1) Then
        .ShowCell i, 1
        .Select i, 1
        If Val(Format(.TextMatrix(i, 3), "###0")) < Val(Format(.TextMatrix(i, 4), "###0")) + 1 Then
          MsgBox "El número de productos con código " & codigo & " Excede la cantidad requerida.", vbInformation, "Cantidad excedida"
        ElseIf Val(Format(.TextMatrix(i, 5), "###0")) < Val(Format(.TextMatrix(i, 4), "###0")) + 1 Then
          MsgBox "El número de productos con código " & codigo & " Excede la cantidad disponible en bodega.", vbInformation, "Cantidad excedida"
        Else
          .TextMatrix(i, 4) = Val(Format(.TextMatrix(i, 4), "###0")) + 1
          .Select i, 0, i, .Cols - 1
          LectorRow = i
'          Dim TempCod As String
'            Dim strSql As String
'            If chkTipoCod.Value = 0 Then
'              TempCod = VSFG.TextMatrix(i, 1)
'            Else
'              TempCod = CodAux2Cod(VSFG.TextMatrix(i, 1))
'            End If
'            'Actualiza un detalle de pedido
'            strSql = " UPDATE det_pedido SET " & _
'                     " det_ped_cant_entregada=" & VSFG.TextMatrix(i, 4) & _
'                     " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & LblPedido & _
'                     " AND prd_codigo='" & TempCod & "' AND dep_codigo='" & VSFG.TextMatrix(i, 0) & "'"
'            clsSql.Ejecutar (strSql)
        End If
        pas = True
      End If
    Next i
  End With
  If pas = False Then
    MsgBox "No se ha encontrado el producto con el código especificado." & vbCr & "Asegúrese el tipo de código del producto y que el mismo se encuentre en lista.", vbCritical, "Error de codigo"
  End If

End Sub


Private Sub Command1_Click()
    drptFActura.Tag = 14463
    drptFActura.Show
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Verifica cuado se presionó un enter para devolver un tab
    If KeyCode = vbKeyReturn And Screen.ActiveControl.Name <> "txtLector" And Screen.ActiveControl.Name <> "txtPedido" Then
        KeyCode = 0
        SendKeys vbKeyTab
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

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    'Inicializa los objetos de conexión con la base de datos
    clsPedidosV.Inicializar AdoConn, AdoConnMaster
    clsPed.Inicializar AdoConn, AdoConnMaster
    clsExiPrd.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsPedidos.Inicializar AdoConn, AdoConnMaster
    banTm = 0
    
    'Consulta todos los pedidos que pasan a bodega para ser revisados
    strSql = " SELECT RIGHT(ped_codigo,7)+0 as c,ped_fecha,ped_observacion,CONCAT(per_apellido,' ',per_nombre) as nombC,ped_estado,cot_codigo,ped_estado,pedido.ven_codigo,ped_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado " & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo = persona.emp_codigo AND pedido.per_codigo = persona.per_codigo " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' AND cat_p_tipo='C' AND ped_estado<>3 AND " & _
             " (IIF(ped_estado=2,ped_fecha='" & Format(HoyDia, "yyyy-MM-dd") & "',1=1)) " & _
             " AND pedido.ped_codigo LIKE CONCAT('" & strSucursal & strPtoFactura & "'+0,'%') " & _
             " ORDER BY ped_estado,ped_codigo "
    'clsPedidosV.Ejecutar (strSql)
    'Muestra los datos de los distintos pedidos en un listado
    'Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
    
    'Almacena el número de pedidos mostrados
    numPeds = VSFGPeds.Rows
    cargarTipoPedido
''''    'Consulta las existencias de productos en bodega
''''    strSql = " SELECT deposito.dep_codigo, producto.prd_codigo, Sum(existencia.exi_cantidad) AS Exis " & _
''''             " FROM (deposito INNER JOIN existencia ON (deposito.dep_codigo = existencia.dep_codigo) " & _
''''             " AND (deposito.emp_codigo = existencia.emp_codigo)) INNER JOIN producto " & _
''''             " ON (existencia.prd_codigo = producto.prd_codigo) AND (existencia.emp_codigo = producto.emp_codigo) " & _
''''             " Where producto.prd_baja=0 And deposito.emp_codigo='" & strEmpresa & "' " & _
''''             " GROUP BY deposito.dep_codigo, producto.prd_codigo, producto.prd_baja, producto.emp_codigo " & _
''''             " ORDER BY deposito.dep_codigo, producto.prd_codigo "
''''    clsExiPrd.Ejecutar (strSql)
End Sub

'Verifica si cada 10 segundos existe un nuevo pedido a revisar
Private Sub TmrAct_Timer()
    On Error Resume Next
    clsPedidosV.Actualizar
    banTm = banTm + 1
    If banTm = 2 Then
        'clsExiPrd.Actualizar
        banTm = 0
    End If
    'Muestra los datos de los distintos pedidos en un listado
    Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
    If VSFGPeds.Rows > numPeds Then
        MsgBox "Nuevo pedido llegado...", vbInformation, "Pedido"
        Me.SetFocus
    End If
    numPeds = VSFGPeds.Rows
End Sub

Private Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim noPasa As Boolean
    
    If KeyCode = vbKeyReturn Then
        Dim strCli As String
        cmdLimpiar_Click
        If cmbNegocio.BoundText <> "" Then
            
            strSql = " SELECT tip_ped_ptofac " & _
                     " FROM tipo_pedido " & _
                     " WHERE tip_ped_codigo='" & cmbNegocio.BoundText & "' "
            clsSql.Ejecutar strSql
            If chkCIRUC.Value = 1 Then
                strCli = "CONCAT(per_ruc,' ',per_apellido,' ',per_nombre)"
            Else
                strCli = "CONCAT(per_apellido,' ',per_nombre)"
            End If
            If Me.optNoPedido.Value = True Then
                If clsSql.adorec_Def.RecordCount > 0 Then
                    banTm = 0
                    'Consulta todos los pedidos que pasan a bodega para ser revisados
                    strSql = " SELECT RIGHT(ped_codigo,7)+0 as c,ped_fecha,ped_observacion," & strCli & " as nombC,ped_estado,cot_codigo,ped_estado,pedido.ven_codigo,ped_codigo,IIF(persona.per_bloqueado+persona.per_bloqueado_g=0,0,1) as per_bloqueado,ped_egr_codigo " & _
                             " FROM pedido INNER JOIN persona ON pedido.emp_codigo = persona.emp_codigo AND pedido.per_codigo = persona.per_codigo AND cat_p_tipo='C' AND persona.tip_ped_codigo='" & Me.cmbNegocio.BoundText & "'" & _
                             " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                             " pedido.ped_codigo = '" & txtPedido.Text & "' " & _
                             " ORDER BY pedido.ped_estado,pedido.ped_codigo "
                    clsPedidosV.Ejecutar (strSql)
                    'Muestra los datos de los distintos pedidos en un listado
                    Set VSFGPeds.DataSource = clsPedidosV.adorec_Def.DataSource
                    Pasa = True
                        strSql = " SELECT est_codigo,est_descripcion " & _
                                 " FROM est_pedido " & _
                                 " ORDER BY est_codigo"
                        clsSql.Ejecutar strSql
                        
                        'Carga los depósitos en el combo de la columna 1 del flexGrid vsfgImp
                        VSFGPeds.ColComboList(4) = VSFGPeds.BuildComboList(clsSql.adorec_Def, "est_descripcion", "est_codigo")
                    
                    If VSFGPeds.Rows > 1 Then
                        noPasa = EgresoRet(txtPedido.Text)
                        If noPasa = False Then
                            VSFGPeds.Row = 1
                            VSFGPeds.Col = 4
                            VSFGPeds.Select 1, 4
                            VSFGPeds_DblClick
                            txtLector.SetFocus
                        Else
                            MsgBox "Este pedido no puede salir, FACTURA RETENIDA", vbCritical
                        End If
                    End If
                End If
            End If
        Else
            Exit Sub
        End If
        banTm = banTm + 1
        If banTm = 2 Then
            'clsExiPrd.Actualizar
            banTm = 0
        End If
        numPeds = VSFGPeds.Rows
        txtPedido.Text = ""
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        'Verifca que solo se ingresen números en el campo de cantidad a entregar
        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en la cantidad.", vbInformation, "Cantidad"
            VSFG.TextMatrix(Row, 4) = intDato
        Else
            VSFG.TextMatrix(Row, 4) = FormatoD4(VSFG.TextMatrix(Row, 4))
        End If
        'Verifica que no se esté pidiendo más productos de los que hay en existencia
        If Val(VSFG.TextMatrix(Row, 4)) > Val(VSFG.TextMatrix(Row, 5)) And Left(VSFG.TextMatrix(Row, 1), 3) <> "PR-" Then
            If FormatoD4(VSFG.TextMatrix(Row, 5)) <= 0 Then
                MsgBox "No hay existencia del este producto en la bodega.", vbInformation, "Existencia"
                VSFG.TextMatrix(Row, 4) = 0
            Else
                MsgBox "Solo hay diponible " & VSFG.TextMatrix(Row, 5) & " unidades de este producto en esta bodega.", vbInformation, "Cantidad"
                VSFG.TextMatrix(Row, 4) = VSFG.TextMatrix(Row, 5)
            End If
        End If
        'Verifica que no se pidan más productos de los pedidos
        If Val(VSFG.TextMatrix(Row, 4)) > Val(VSFG.TextMatrix(Row, 3)) Then
            MsgBox "Solo puede entregar " & VSFG.TextMatrix(Row, 3) & " unidades del producto.", vbInformation, "Entregar"
            VSFG.TextMatrix(Row, 4) = VSFG.TextMatrix(Row, 3)
        End If
    End If
    CalculaCant
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Permite modificar solo la columna 4 de la cantidad a entregar
    If Col = 4 Or Col = 6 Then
        'Captura el valor actual de la celda de cantidad a entregar
        intDato = VSFG.TextMatrix(Row, Col)
    Else
        Cancel = True
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim strFiltroPed As String
    Dim cantExi As Long
    'Verifica cuando haya datos en una fila del grid tanto en bodega como en producto
    'para obtener la existencia de un producto en bodega
    If Row > 0 And VSFG.TextMatrix(Row, 0) <> "" And VSFG.TextMatrix(Row, 1) <> "" Then
        Dim strFiltro As String
        'Encuentra la existencia de un producto en una bodega específica
        strFiltro = "existencia.dep_codigo='" & VSFG.TextMatrix(Row, 0) & "' AND existencia.prd_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
        strFiltroPed = "det_pedido.dep_codigo='" & VSFG.TextMatrix(Row, 0) & "' AND det_pedido.prd_codigo='" & VSFG.TextMatrix(Row, 1) & "'"
''''ÁQUI**********************************
        strSql = " SELECT existencia.dep_codigo, existencia.prd_codigo, existencia.emp_codigo, Sum(existencia.exi_cantidad) AS Exis " & _
                 " FROM existencia " & _
                 " Where existencia.emp_codigo='" & strEmpresa & "' " & _
                 " AND " & strFiltro & _
                 " GROUP BY existencia.dep_codigo, existencia.prd_codigo, existencia.emp_codigo "
        clsExiPrd.Ejecutar (strSql)
        cantExi = clsExiPrd.adorec_Def("Exis")
        strSql = " SELECT det_pedido.dep_codigo, det_pedido.prd_codigo, det_pedido.emp_codigo, Sum(det_pedido.det_ped_cant_entregada) AS Exis " & _
                 " FROM pedido INNER JOIN det_pedido " & _
                 " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " Where pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_estado=1 " & _
                 " AND " & strFiltroPed & _
                 " GROUP BY det_pedido.dep_codigo, det_pedido.prd_codigo, det_pedido.emp_codigo "
        clsExiPrd.Ejecutar (strSql)
        If clsExiPrd.adorec_Def.RecordCount > 0 Then
            
            cantExi = cantExi - clsExiPrd.adorec_Def("Exis")
        End If
        'clsExiPrd.Filtrar (strFiltro)
        VSFG.TextMatrix(Row, 5) = cantExi + VSFG.TextMatrix(Row, 3)
        'clsExiPrd.QuitarFiltro
        If VSFG.TextMatrix(Row, 5) = "" Then VSFG.TextMatrix(Row, 5) = 0
    End If
End Sub

Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
    If VSFG.Col = 4 Then
        VSFG.AutoSearch = flexSearchNone
    Else
        VSFG.AutoSearch = flexSearchFromTop
    End If
End Sub

Private Sub VSFGPeds_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'Marca toda la fila con otra tonalidad si ese pedido necesita ser revisado
    If Col = 4 And Row <> 0 Then
        If Col = 4 And Row <> 0 And VSFGPeds.TextMatrix(Row, Col) = 0 Then
            VSFGPeds.Select Row, 0, Row, VSFGPeds.Cols - 1
            VSFGPeds.FillStyle = flexFillRepeat
            VSFGPeds.CellBackColor = &HC0C0FF
        End If
    End If
End Sub

Private Sub VSFGPeds_DblClick()
    'Verifica cuando se da un doble click sobre una fila del grid de pedidos
    If VSFGPeds.Row > 0 Then
        'Consulta el detalle de un pedido específico
        strSql = " SELECT dep_codigo, det_egreso.prd_codigo, prd_nombre, det_egr_cantidad, IIF(det_egreso.prd_codigo='PR-CARGOO100330TU',1,0) as det_egr_cantidad, '' as HH,'' as Descripcion " & _
                 " FROM (pedido INNER JOIN det_egreso ON (pedido.emp_codigo = det_egreso.emp_codigo) " & _
                 " AND (pedido.ped_egr_codigo = det_egreso.egr_codigo)" & _
                 " AND (pedido.ped_tip_egr_codigo = det_egreso.tip_egr_codigo)" & _
                 " INNER JOIN producto " & _
                 " ON (det_egreso.emp_codigo = producto.emp_codigo) AND (det_egreso.prd_codigo = producto.prd_codigo)) " & _
                 " Where pedido.emp_codigo='" & strEmpresa & "' AND " & _
                 " pedido.ped_codigo= " & VSFGPeds.TextMatrix(VSFGPeds.Row, 8) & _
                 " ORDER BY det_egreso.prd_codigo "
        clsPed.Ejecutar (strSql)
        'Debug.Print strSql
        'Muestra el detalle de pedido en un grid
        Set VSFG.DataSource = clsPed.adorec_Def.DataSource
        'Muestra el número del pedido a modificar
        LblPedido.Caption = VSFGPeds.TextMatrix(VSFGPeds.Row, 8)
        'Captura el código de la cotización si este existe
        codCot = Val(VSFGPeds.TextMatrix(VSFGPeds.Row, 4))
        strFac = VSFGPeds.TextMatrix(VSFGPeds.Row, 10)
        'Cliente Bloqueado
        If Abs(VSFGPeds.TextMatrix(VSFGPeds.Row, 9)) <> 0 Then
            CmdConfirmar.Enabled = False
            'CmdDeBaja.Enabled = False
            MsgBox "CLIENTE BLOQUEADO"
        Else
        
            'Verifica que no se pueda modificar un pedido ya revisado
            If VSFGPeds.TextMatrix(VSFGPeds.Row, 4) <> 8 Then
                VSFG.Editable = flexEDNone
                CmdConfirmar.Enabled = False
                'CmdDeBaja.Enabled = False
                MsgBox "El estado del pedido (" & VSFGPeds.Cell(flexcpTextDisplay, VSFGPeds.Row, 4) & ") no permite confirmar", vbCritical
            
            ElseIf VSFGPeds.TextMatrix(VSFGPeds.Row, 4) = 1 Then
                VSFG.Editable = flexEDNone
                CmdConfirmar.Enabled = False
                MsgBox "El estado del pedido (" & VSFGPeds.Cell(flexcpTextDisplay, VSFGPeds.Row, 4) & ") no permite confirmar", vbCritical
                'CmdDeBaja.Enabled = True
            ElseIf VSFGPeds.TextMatrix(VSFGPeds.Row, 4) = 8 Then
                VSFG.Editable = flexEDKbdMouse
                CmdConfirmar.Enabled = True
                'CmdDeBaja.Enabled = True
            End If
        End If
        CalculaCant
    End If
End Sub

Private Sub VSFGPeds_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        VSFGPeds_DblClick
    End If
End Sub

Private Sub CalculaCant()
    Dim i As Long
    Dim CP As Double
    Dim CE As Double
    
    CP = 0
    CE = 0
    For i = 1 To VSFG.Rows - 1
        CP = CP + FormatoD4(VSFG.TextMatrix(i, 3))
        CE = CE + FormatoD4(VSFG.TextMatrix(i, 4))
    Next i
    txtCantPed.Text = FormatoD2(CP)
    txtCantEnt.Text = FormatoD2(CE)

End Sub
