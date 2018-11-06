VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGenerarArchivoCredito 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo Datos Crediticios"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerarArchivoCredito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13395
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
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13200
      Begin VB.TextBox txtCodigoEntidad 
         Height          =   315
         Left            =   11550
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker FechaValidez 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69599233
         CurrentDate     =   41632
      End
      Begin VB.CommandButton cmdGenerarArchivo 
         Caption         =   "&Generar Archivo"
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   2
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsultaCartera 
         Caption         =   "Consulta Cartera a Pagar"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   12975
         _cx             =   51861862
         _cy             =   51848527
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmGenerarArchivoCredito.frx":030A
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
      Begin NEED2.uctrVSFG ucrtVSFG1 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Entidad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   10320
         TabIndex        =   8
         Top             =   390
         Width           =   1110
      End
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   12720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
End
Attribute VB_Name = "frmGenerarArchivoCredito"
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
Private strSQL As String
Private Function JustificarIzquierdaConEspacios(cadena As String, largo As Long) As String
    Dim i As Long
    If largo < Len(cadena) Then
        JustificarIzquierdaConEspacios = Left(cadena, largo)
    Else
        JustificarIzquierdaConEspacios = cadena
        For i = 1 To largo - Len(cadena)
            JustificarIzquierdaConEspacios = JustificarIzquierdaConEspacios & " "
        Next i
    End If
End Function

Private Sub cmdConsultaCartera_Click()
    Dim i As Long
    
    strSQL = " CREATE TABLE CompRet " & _
             "( cue_p_c_codigo decimal(14,0)," & _
             " cue_p_c_tipo char(1)," & _
             " emp_codigo char(3)," & _
             " tipo_cartera char(3)," & _
             " reten decimal(14,2)," & _
             " key pri(cue_p_c_codigo,cue_p_c_tipo,emp_codigo,tipo_cartera))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TABLE CompRetV " & _
             "( cue_p_c_codigo decimal(14,0)," & _
             " cue_p_c_tipo char(1)," & _
             " emp_codigo char(3)," & _
             " tipo_cartera char(3)," & _
             " reten decimal(14,2)," & _
             " key pri(cue_p_c_codigo,cue_p_c_tipo,emp_codigo,tipo_cartera))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO CompRet " & _
             " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,'VEN' as tipo_cartera,COALESCE(comprobante_retencion.com_ret_total,0) as reten " & _
             " FROM persona INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN cuenta_p_c ON persona.emp_codigo=cuenta_p_c.emp_codigo AND persona.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo " & _
             " AND comprobante_retencion.com_ret_fecha <= '" & FechaValidez.Value & "'" & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta <= '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo" & _
             " ORDER BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO CompRetV " & _
             " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,'VIG' as tipo_cartera,COALESCE(comprobante_retencion.com_ret_total,0) as reten " & _
             " FROM persona INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN cuenta_p_c ON persona.emp_codigo=cuenta_p_c.emp_codigo AND persona.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo = comprobante_retencion.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = comprobante_retencion.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = comprobante_retencion.emp_codigo " & _
             " AND comprobante_retencion.com_ret_fecha <= '" & FechaValidez.Value & "'" & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta > '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo" & _
             " ORDER BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TABLE Pag " & _
             "( cue_p_c_codigo decimal(14,0)," & _
             " cue_p_c_tipo char(1)," & _
             " emp_codigo char(3)," & _
             " tipo_cartera char(3)," & _
             " abono decimal(14,2)," & _
             " estado decimal(14,2)," & _
             " abonoch decimal(14,2)," & _
             " estadoP decimal(14,2)," & _
             " key pri(cue_p_c_codigo,cue_p_c_tipo,emp_codigo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " CREATE TABLE PagV " & _
             "( cue_p_c_codigo decimal(14,0)," & _
             " cue_p_c_tipo char(1)," & _
             " emp_codigo char(3)," & _
             " tipo_cartera char(3)," & _
             " abono decimal(14,2)," & _
             " estado decimal(14,2)," & _
             " abonoch decimal(14,2)," & _
             " estadoP decimal(14,2)," & _
             " key pri(cue_p_c_codigo,cue_p_c_tipo,emp_codigo))"
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO Pag " & _
             " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,'VEN' as tipo_cartera,COALESCE(SUM(pag_monto),0) as abono, " & _
             " 0 as estado,0.00 as abonoch,0 as estadoP " & _
             " FROM persona INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN cuenta_p_c ON persona.emp_codigo=cuenta_p_c.emp_codigo AND persona.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo " & _
             " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = pago.emp_codigo " & _
             " AND pago.pag_fecha <= '" & FechaValidez.Value & "'" & _
             " LEFT JOIN doc_pago ON pago.doc_pag_codigo = doc_pago.doc_pag_codigo " & _
             " AND pago.emp_codigo = doc_pago.emp_codigo AND doc_pago.doc_pag_pendiente >= -1 " & _
             " AND doc_pago.doc_pag_fecha_recepcion <= '" & FechaValidez.Value & "'" & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta <= '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo" & _
             " ORDER BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " INSERT INTO PagV " & _
             " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,'VIG' as tipo_cartera,COALESCE(SUM(pag_monto),0) as abono, " & _
             " 0 as estado,0.00 as abonoch,0 as estadoP " & _
             " FROM persona INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN cuenta_p_c ON persona.emp_codigo=cuenta_p_c.emp_codigo AND persona.per_codigo=cuenta_p_c.per_codigo " & _
             " INNER JOIN pago ON cuenta_p_c.cue_p_c_codigo = pago.cue_p_c_codigo " & _
             " AND cuenta_p_c.cue_p_c_tipo = pago.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = pago.emp_codigo " & _
             " AND pago.pag_fecha <= '" & FechaValidez.Value & "'" & _
             " LEFT JOIN doc_pago ON pago.doc_pag_codigo = doc_pago.doc_pag_codigo " & _
             " AND pago.emp_codigo = doc_pago.emp_codigo AND doc_pago.doc_pag_pendiente >= -1 " & _
             " AND doc_pago.doc_pag_fecha_recepcion <= '" & FechaValidez.Value & "'" & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta > '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo" & _
             " ORDER BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo "
    clsCon_Def.Ejecutar strSQL
    strSQL = " SELECT '" & txtCodigoEntidad.Text & "' as codent,DATE_FORMAT('" & FechaValidez.Value & "','%d/%m/%Y') as fechadatos,IF(LEN(per_ruc)=13,'R',IF(LEN(per_ruc)=10,'C','E')) as tipiden,per_ruc, CONCAT(per_apellido,' ', per_nombre) as persona,  " & _
             " LEFT(per_tipo,1) as calsesujero, LEFT(ciu_codigo_din,2),RIGHT(ciu_codigo_din,2),'' as par_codigo_din,'' as sexo,'' as estadocivil,'' as origeningreso, cue_p_c_egr_codigo, " & _
             " COALESCE(cue_p_c_valor,0) as valor, ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00) ,2),DATE_FORMAT(cue_p_c_fechaemision,'%d/%m/%Y') as emision, '' as vencimiento, DATE_FORMAT(cue_p_c_fechapropuesta,'%d/%m/%Y') as vencimiento," & _
             " '' as plazo,1 as perio, (TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)) as diasmorosidad, ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00) ,2), 0 as montointeresmora, " & _
             " 0 as vencer1_30,0 as vencer31_90,0 as vencer91_180,0 as vencer181_360,0 as vencer360, " & _
             " IF(1 <= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)<= 30,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00),2),0) as vencido1_30," & _
             " IF(31 <= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)<= 90,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00),2),0) as vencido31_90," & _
             " IF(91 <= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)<= 180,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00),2),0) as vencido91_180, " & _
             " IF(181 <= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)<= 360,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00),2),0) as vencido181_360, " & _
             " IF(361 <= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) ,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00),2),0) as vencido361, " & _
             " 0 as demandajudicial,0 as carteracastigada,0 as cuotacredito,DATE_FORMAT(DATE_ADD(cue_p_c_fechaemision, INTERVAL 1 YEAR),'%d/%m/%Y') as fechacancela,'' as forpago "
    strSQL = strSQL & " FROM ((((cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo='C') INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN ciudad ON persona.ciu_codigo = ciudad.ciu_codigo" & _
             " LEFT JOIN egreso ON cuenta_p_c.emp_codigo = egreso.emp_codigo " & _
             " AND cuenta_p_c.cue_p_c_egr_codigo * 1 = egreso.egr_codigo " & _
             " AND cuenta_p_c.per_codigo = egreso.per_codigo" & _
             " AND egreso.egr_fecha <= '" & FechaValidez.Value & "') " & _
             " LEFT JOIN Pag ON cuenta_p_c.cue_p_c_codigo = Pag.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = Pag.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = Pag.emp_codigo AND Pag.tipo_cartera='VEN') " & _
             " LEFT JOIN CompRet ON cuenta_p_c.cue_p_c_codigo = CompRet.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = CompRet.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = CompRet.emp_codigo AND CompRet.tipo_cartera='VEN') "
    strSQL = strSQL & " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cue_p_c_egr_codigo NOT LIKE 'R%' AND tip_doc_cue_codigo=1 AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta <= '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " AND ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(Pag.abono,0) - COALESCE(CompRet.reten,0) - COALESCE(Pag.abonoch,0.00) ,2)>0.00 "
    strSQL = strSQL & " UNION" & _
             " SELECT '" & txtCodigoEntidad.Text & "' as codent,DATE_FORMAT('" & FechaValidez.Value & "','%d/%m/%Y') as fechadatos,IF(LEN(per_ruc)=13,'R',IF(LEN(per_ruc)=10,'C','E')) as tipiden,per_ruc, CONCAT(per_apellido,' ', per_nombre) as persona,  " & _
             " LEFT(per_tipo,1) as calsesujero, LEFT(ciu_codigo_din,2),RIGHT(ciu_codigo_din,2),'' as par_codigo_din,'' as sexo,'' as estadocivil,'' as origeningreso, cue_p_c_egr_codigo, " & _
             " COALESCE(cue_p_c_valor,0) as valor, ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00) ,2), DATE_FORMAT(cue_p_c_fechaemision,'%d/%m/%Y') as emision, '' as vencimiento, DATE_FORMAT(cue_p_c_fechapropuesta,'%d/%m/%Y') as vencimiento, " & _
             " '' as plazo,1 as perio, (TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)) as diasmorosidad, ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2), 0 as montointeresmora, " & _
             " IF(-1 >= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)>= -30,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2),0) as vencer1_30," & _
             " IF(-31 >= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)>= -90,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2),0) as vencer31_90," & _
             " IF(-91 >= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)>= -180,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2),0) as vencer91_180, " & _
             " IF(-181 >= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) AND TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta)>= -360,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2),0) as vencer181_360, " & _
             " IF(-361 >= TO_DAYS('" & FechaValidez.Value & "') - TO_DAYS(cue_p_c_fechapropuesta) ,ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00),2),0) as vencer361, " & _
             " 0 as vencido1_30,0 as vencido31_90,0 as vencido91_180,0 as vencido181_360,0 as vencido361, " & _
             " 0 as demandajudicial,0 as carteracastigada,0 as cuotacredito,DATE_FORMAT(DATE_ADD(cue_p_c_fechaemision, INTERVAL 1 YEAR),'%d/%m/%Y') as fechacancela,'' as forpago"
    strSQL = strSQL & " FROM ((((cuenta_p_c INNER JOIN persona ON cuenta_p_c.per_codigo = persona.per_codigo " & _
             " AND cuenta_p_c.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo='C') INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
             " AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " AND forma_pago.for_pag_tiempo!=0 " & _
             " INNER JOIN ciudad ON persona.ciu_codigo = ciudad.ciu_codigo " & _
             " LEFT JOIN egreso ON cuenta_p_c.emp_codigo = egreso.emp_codigo " & _
             " AND cuenta_p_c.cue_p_c_egr_codigo * 1 = egreso.egr_codigo " & _
             " AND cuenta_p_c.per_codigo = egreso.per_codigo" & _
             " AND egreso.egr_fecha > '" & FechaValidez.Value & "') " & _
             " LEFT JOIN PagV ON cuenta_p_c.cue_p_c_codigo = PagV.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = PagV.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = PagV.emp_codigo AND PagV.tipo_cartera='VIG') " & _
             " LEFT JOIN CompRetV ON cuenta_p_c.cue_p_c_codigo = CompRetV.cue_p_c_codigo  " & _
             " AND cuenta_p_c.cue_p_c_tipo = CompRetV.cue_p_c_tipo " & _
             " AND cuenta_p_c.emp_codigo = CompRetV.emp_codigo AND CompRetV.tipo_cartera='VIG') "
    strSQL = strSQL & " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cue_p_c_egr_codigo NOT LIKE 'R%' AND tip_doc_cue_codigo=1 AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
             " AND cue_p_c_fechapropuesta > '" & FechaValidez.Value & "'" & _
             " AND cue_p_c_fechaemision <= '" & FechaValidez.Value & "'" & _
             " AND ROUND(COALESCE(cue_p_c_valor,0) - COALESCE(PagV.abono,0) - COALESCE(CompRetV.reten,0) - COALESCE(PagV.abonoch,0.00) ,2)!=0 "
'    strSql = strSql & " UNION " & _
'             " SELECT CONCAT(per_apellido,' ', per_nombre,' (',tip_ped_codigo,')') as persona, '1/1' as pagos, " & _
'             " ing_fecha as emision, ing_fecha as vencimiento, ing_fecha as ultimo, " & _
'             " COALESCE(-1 * ing_total,0.000) as valor,COALESCE(-1 * ing_saldo,0.000) as abonoNC, 'NOTA DE CREDITO' as descripcion," & _
'             " 'NC' as numero, ing_codigo, (TO_DAYS('" & FechaValidez.value & "') - TO_DAYS(ing_fecha)) as tipoCar, ingreso.per_codigo as p_cod, " & _
'             " 0 as reten,0 as estado,0 as abonoch,0 as estadoP,per_direccion,per_telf,'VEN' as tipo_cartera,persona.per_ruc" & _
'             " FROM ingreso INNER JOIN persona ON ingreso.per_codigo = persona.per_codigo " & _
'             " AND ingreso.emp_codigo = persona.emp_codigo AND persona.cat_p_tipo = 'C' " & _
'             " WHERE ingreso.emp_codigo = '" & strEmpresa & "' " & _
'             " AND ingreso.tip_ing_codigo='DCL' " & _
'             " AND ingreso.ing_fecha <= '" & FechaValidez.value & "' AND ingreso.ing_anulado=0" & _
'             " AND ROUND(COALESCE(ing_total,0.000) - COALESCE(ing_saldo,0.000),2)>0 " & _
'             " ORDER BY persona,tipoCar,cue_p_c_egr_codigo "
    clsCon_Def.Ejecutar strSQL
    
    Set VSFG2.DataSource = clsCon_Def.adorec_Def.DataSource
    
    strSQL = " DROP TABLE CompRet "
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE Pag "
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE CompRetV "
    clsCon_Def.Ejecutar strSQL
    strSQL = " DROP TABLE PagV "
    clsCon_Def.Ejecutar strSQL
    
End Sub

Private Function QuitarCaracteresEspeciales(cadena As String) As String
    Dim i As Long
    Dim CadenaFinal As String
    Dim Caracter As String
    CadenaFinal = ""
    For i = 1 To Len(cadena)
        Caracter = Mid(cadena, i, 1)
        If Caracter <> " " Then
            '65-90 A-Z 165 Ñ
            Caracter = Replace(Replace(Caracter, "Ñ", "N"), vbNewLine, " ")
            If Not (Asc(Caracter) >= 65 And Asc(Caracter) <= 90) Or IsNumeric(Caracter) = True Then
                Caracter = ""
            End If
        End If
        CadenaFinal = CadenaFinal & Caracter
    Next i
    QuitarCaracteresEspeciales = CadenaFinal
End Function

Private Sub cmdGenerarArchivo_Click()
    Dim sDir As String
    Dim ArchivoDefecto As String
    ArchivoDefecto = ""
    sDir = CurDir
    cdArchivo.Filter = "Todos los Archivos|*.*|Archivos de texto .txt|*.txt"
    cdArchivo.FileName = ArchivoDefecto
    cdArchivo.ShowSave
    ChDir sDir
    If (cdArchivo.FileName <> "") Then
        Me.MousePointer = 11
        VSFG2.ClipSeparators = "|" & vbCr
        VSFG2.SaveGrid cdArchivo.FileName, flexFileCustomText
        Me.MousePointer = 0
    End If

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
    Set ucrtVSFG1.VSFGControl = VSFG2
    ucrtVSFG1.Inicializar False, False, False
    FechaValidez.Value = HoyDia
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

