VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDesaplicarNota 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desaplicar Notas de Crédito"
   ClientHeight    =   3870
   ClientLeft      =   6660
   ClientTop       =   2460
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesaplicarNota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6990
   Begin VSFlex8Ctl.VSFlexGrid VSFG 
      Height          =   855
      Left            =   360
      TabIndex        =   22
      Top             =   2280
      Width           =   3615
      _cx             =   6376
      _cy             =   1508
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDesaplicarNota.frx":030A
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
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Desaplicar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista Previa"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
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
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtro de Notas de Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   788
      TabIndex        =   5
      Top             =   120
      Width           =   5415
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCotizacion 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   4065
         _ExtentX        =   7170
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   405
         Width           =   630
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Crédito:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   930
      End
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO ANULADO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   1800
      Width           =   2355
   End
   Begin VB.Label Label6 
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
      Left            =   4440
      TabIndex        =   19
      Top             =   2910
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recargos:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   15
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   14
      Top             =   2070
      Width           =   825
   End
   Begin VB.Label LblIva 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA X%"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   13
      Top             =   2310
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4440
      TabIndex        =   12
      Top             =   1830
      Width           =   630
   End
End
Attribute VB_Name = "frmDesaplicarNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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
Private strSql As String
Private clsSql As New clsConsulta
Private clsFPago As New clsConsulta
Private numaux As String
Private numaux2 As String

Private Sub cmbCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        strSql = " SELECT per_codigo,CONCAT(per_apellido,' ',per_nombre, ' (',tip_ped_nombre,')') as nombC " & _
                 " FROM persona " & _
                 " INNER JOIN tipo_pedido ON tipo_pedido.emp_codigo=persona.emp_codigo AND tipo_pedido.tip_ped_codigo=persona.tip_ped_codigo " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' AND persona.cat_p_tipo='C' " & _
                 " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " AND CONCAT(persona.per_apellido,' ',persona.per_nombre) like '" & cmbCliente.Text & "%'" & _
                 " ORDER BY nombC "
        clsSql.Ejecutar strSql
        
        Set cmbCliente.RowSource = clsSql.adorec_Def.DataSource
            
        cmbCliente.ListField = "nombC"
        cmbCliente.BoundColumn = "per_codigo"
    End If
End Sub

Private Sub cmbCotizacion_Change()
    numaux = ""
    numaux2 = ""
    If cmbCotizacion.MatchedWithList = True Then
        VSFG.Clear 1
        VSFG.Rows = 1
        strSql = " SELECT ing_anulado,CONCAT(per_apellido,' ',per_nombre) as nombC,ing_fecha, ing_subtotal,ing_dcto,ing_impuesto,ing_subtotal_o,ing_total,ing_numasiento as num,ing_saldo,ing_serie,ing_numero " & _
                 " FROM ingreso INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND  ingreso.per_codigo=persona.per_codigo " & _
                 " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND ingreso.ing_codigo='" & cmbCotizacion.Text & "' " & _
                 " AND ingreso.tip_ing_codigo='DCL' "
        clsSql.Ejecutar strSql
        txtFecha.Text = clsSql.adorec_Def("ing_fecha")
        TxtSubTotal.Text = Format(clsSql.adorec_Def("ing_subtotal"), "###0.00")
        TxtDesc.Text = Format(clsSql.adorec_Def("ing_dcto"), "###0.00")
        TxtIva.Text = Format(clsSql.adorec_Def("ing_impuesto"), "###0.00")
        TxtRecargo.Text = Format(clsSql.adorec_Def("ing_subtotal_o"), "###0.00")
        TxtTotal.Text = Format(clsSql.adorec_Def("ing_total"), "###0.00")
        cmbCotizacion.Tag = clsSql.adorec_Def("num")
        numaux = clsSql.adorec_Def("ing_serie") & "-" & clsSql.adorec_Def("ing_numero")
        numaux2 = clsSql.adorec_Def("ing_serie") & "-" & Format(clsSql.adorec_Def("ing_numero"), "0000000")
        If FormatoD0(clsSql.adorec_Def("ing_anulado")) = 1 Then
            lblEstado.Caption = "ANULADO"
            cmdAnular.Enabled = False
            cmdCambiar.Enabled = False
        Else
            lblEstado.Caption = ""
            cmdAnular.Enabled = True
            If MesCerrado(txtFecha.Text) = True Then
                cmdAnular.Enabled = False
            End If
            If FormatoD4(clsSql.adorec_Def("ing_saldo")) = 0 Then
                cmdCambiar.Enabled = False
            Else
                cmdCambiar.Enabled = True
                 strSql = " SELECT cuenta_p_c.cue_p_c_egr_codigo,pag_monto,pago.cue_p_c_codigo,pago.pag_codigo " & _
                          " FROM pago " & _
                          " INNER JOIN cuenta_p_c " & _
                          " ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                          " AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
                          " AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                          " WHERE pago.emp_codigo='" & strEmpresa & "' AND pago.cue_p_c_tipo='C' " & _
                          " AND pago.asi_numasiento='" & cmbCotizacion.BoundText & "' " & _
                          " AND pag_observacion LIKE '%NOTA DE CR%DITO%' "
                          
                          'AND (pag_no_doc='" & cmbCotizacion.BoundText & "' OR pag_no_doc LIKE '" & numaux & "' OR pag_no_doc LIKE '" & numaux2 & "')
                clsSql.Ejecutar strSql
                Set VSFG.DataSource = clsSql.adorec_Def.DataSource
                
            End If
        End If
        
    End If
End Sub

Private Sub cmdAnular_Click()
    Dim Motivo As String
    Dim Anula As Boolean
    Dim clsAsiento As New clsContable
    clsAsiento.Inicializar AdoConn, AdoConnMaster
    Dim Puede As Boolean: Puede = False
    If Left(txtFecha.Text, 7) = Left(HoyDia, 7) Then
        Puede = True
    ElseIf Left(txtFecha.Text, 7) = Left(DateAdd("m", -1, HoyDia), 7) And Right(Left(HoyDia, 10), 2) + 0 <= 5 Then
        Puede = True
    End If
    If Puede = False Then
        frmClave.strClaveMAESTRA = strClaveMAESTRA
        frmClave.dblPrecio = "Anulacion"
        frmClave.Show vbModal
        If frmClave.Ret = False Then
            Puede = False
        Else
            Puede = True
        End If
    End If
    
    If Puede = True Then
        Motivo = ""
        While Motivo = ""
            Motivo = InputBox("Motivo de Anulacion", "Contabilidad")
            Motivo = Motivo & vbNewLine & strUsuario & vbNewLine & HoyDia & " " & Format(Ahora, "HH:MM:SS")
            If Motivo = "" Then
                If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el Asiento?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                    Anula = False
                    Motivo = "NO ANULAR"
                End If
            Else
                Anula = True
            End If
        Wend
        If Anula = True Then
            clsAsiento.NumAsiento = ""
            If cmbCotizacion.Tag = "" Then
                'clsAsiento.AnularCX "C", Motivo, cmbCotizacion.BoundText
                If cmdCambiar.Enabled = True Then
                    cmdCambiar_Click
                End If
                
                Dim clsIngreso As New clsInventario
                clsIngreso.Inicializar AdoConn, AdoConnMaster
                clsIngreso.AnularIng cmbCotizacion.BoundText, "DCL", clsAsiento.NumAsiento, Motivo
                Set clsIngreso = Nothing
            Else
                strSql = " SELECT COALESCE(asi_descripcion,'') as descripcion FROM asiento WHERE emp_codigo='" & strEmpresa & "' AND asi_numasiento='" & cmbCotizacion.Tag & "' "
                clsSql.Ejecutar strSql
                clsAsiento.NumAsiento = Right(cmbCotizacion.Tag, 14)
                clsAsiento.AnularAsientoYOtros UCase(Motivo), clsSql.adorec_Def("descripcion")
            End If
            MsgBox "Nota de Crédito # " & cmbCotizacion.BoundText & " anulada", vbInformation, "Anular"
            cmbCotizacion_Change
        End If
    End If
    Set clsAsiento = Nothing
End Sub

Private Sub cmdCambiar_Click()
    Dim cls As New clsConsulta
    Dim facs As String
    Dim i As Long
    cls.Inicializar AdoConn, AdoConnMaster
    
    
        For i = 1 To VSFG.Rows - 1
            strSql = " UPDATE pago " & _
                    " SET pag_no_doc='NC DESAPLICADA', " & _
                    " pag_monto=0, " & _
                    " pag_observacion=CONCAT('NC','" & cmbCotizacion.BoundText & " - DESAPLICADA'), " & _
                    " pag_fechamod=CURRENT_TIMESTAMP, " & _
                    " pag_usumod='" & strUsuario & "' " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND cue_p_c_tipo='C' " & _
                    " AND pag_codigo='" & VSFG.TextMatrix(i, 3) & "' " & _
                    " AND cue_p_c_codigo='" & VSFG.TextMatrix(i, 2) & "' "
            cls.Ejecutar strSql, "M"
            
            strSql = " UPDATE cuenta_p_c " & _
                      " SET cue_p_c_pagado=0 " & _
                      " WHERE emp_codigo='" & strEmpresa & "' " & _
                      " AND cue_p_c_tipo='C' " & _
                      " AND cue_p_c_codigo='" & VSFG.TextMatrix(i, 2) & "' "
            cls.Ejecutar strSql, "M"
        
        Next i
        
        strSql = " UPDATE ingreso " & _
              " SET ing_saldo=0, " & _
              " ing_fechamod=CURRENT_TIMESTAMP, " & _
              " ing_usumod='" & strUsuario & "' " & _
              " WHERE ing_codigo='" & cmbCotizacion.Text & "' " & _
              " AND tip_ing_codigo='DCL' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql
    MsgBox "Nota Crédito " & cmbCotizacion.BoundText & " Desaplicada"
    Set cls = Nothing
    cmbCotizacion_Change
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub cmbCliente_Change()
    If cmbCliente.MatchedWithList = True Then
        strSql = " SELECT ing_codigo,ing_factura,ing_numasiento " & _
                 " FROM ingreso INNER JOIN persona ON (ingreso.emp_codigo = persona.emp_codigo) AND (ingreso.per_codigo = persona.per_codigo) " & _
                 " WHERE tip_ing_codigo='DCL' AND ingreso.emp_codigo='" & strEmpresa & "' AND persona.per_codigo='" & cmbCliente.BoundText & "' and persona.cat_p_tipo='C' " & _
                 " ORDER BY ingreso.ing_codigo "
        clsSql.Ejecutar strSql
        
        cmbCotizacion = ""
        Set cmbCotizacion.RowSource = clsSql.adorec_Def.DataSource
        cmbCotizacion.ListField = "ing_codigo"
        cmbCotizacion.BoundColumn = "ing_numasiento"
        cmbCotizacion.Tag = ""
        lblEstado.Caption = ""
        VSFG.Clear 1
        VSFG.Rows = 1
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVistaPrevia_Click()
    If cmbCliente <> "" And cmbCotizacion <> "" Then
        frmReporte.strNumero = cmbCotizacion.BoundText
        frmReporte.strReporte = "rptNotaCredito"
        frmReporte.Show

        Dim rpTNC2 As New frmReporte
        rpTNC2.strNumero = cmbCotizacion.BoundText
        rpTNC2.strReporte = "rptNotaCreditoUbicacion"
        rpTNC2.Show

    Else
        MsgBox "No ha seleccionado una Nota de Crédito", vbInformation, "Nota de Crédito"
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub cargarTipoPedido()
    
    Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
    cmbNegocio.ListField = "tip_ped_nombre"
    cmbNegocio.BoundColumn = "tip_ped_codigo"
    
    If Trim(strPtoFactura) = "" Then
        frmSelNegocio.Show vbModal
    End If
    
    strSql = " SELECT tip_ped_codigo " & _
             " FROM tipo_pedido " & _
             " WHERE tip_ped_ptofac='" & strPtoFactura & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        cmbNegocio.BoundText = clsSql.adorec_Def(0)
    End If
End Sub


Private Sub Form_Load()

    clsSql.Inicializar AdoConn, AdoConnMaster
    clsFPago.Inicializar AdoConn, AdoConnMaster
    
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    cmbCotizacion.Tag = ""
    lblEstado.Caption = ""
    cmdCambiar.Enabled = False
    
    cargarTipoPedido
    'cmbNegocio_Validate False
    strSql = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
    clsSql.Ejecutar (strSql)
    strClaveMAESTRA = clsSql.adorec_Def("par_texto")

End Sub
