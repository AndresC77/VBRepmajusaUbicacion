VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmImpresionDirecta 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion"
   ClientHeight    =   6960
   ClientLeft      =   3180
   ClientTop       =   1845
   ClientWidth     =   7335
   Icon            =   "frmImpresionDirecta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7335
   Begin VB.PictureBox pic 
      Height          =   3855
      Left            =   4920
      Picture         =   "frmImpresionDirecta.frx":030A
      ScaleHeight     =   3795
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox rtxtImpresion 
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   11668
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmImpresionDirecta.frx":59A3
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6600
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Destino"
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
      Height          =   1095
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pantalla"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optImpresora 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Impresora"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   1700
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5400
      TabIndex        =   0
      Top             =   6360
      Width           =   1700
   End
End
Attribute VB_Name = "frmImpresionDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mod = 0 NADA - 1 ELIMINAR - 2 INSERTAR - 3 MODIFICAR - -2 NADA INSERTAR - -3 NADA MODIF
Private clsCon_Def As New clsConsulta
Private strSql As String
Public strNumero As String
Public strReporte As String
Public lngPag As Long

Public Sub cmdImprimir_Click()
    Dim clsConAUX As New clsConsulta
    Dim CDEClaveAcceso As String
    Dim ClaveAcceso As String
    Dim Autori As String
    Dim TipoPed As String
    Dim dirEnvio As String
    Dim canti As Long
    Dim cantiRepro As Long
    Dim subTotalRepro As Double
    Dim dctoRepro As Double
    Dim IvaRepro As Double
    Dim TOTALPEDIDO As Double
    
    Dim LineasEnBlanco As Long
    Dim CDEPedido As String
    Dim Pedido As String
    Dim lngPagina As Long
    
    Dim strLinea As String
    
    Dim i As Long
    
    Dim TamLetraFactura As Integer
    Dim TamLetraPedido As Integer
    Dim TamLetraSTKDespacho As Integer
    TamLetraFactura = 6
    TamLetraPedido = 7
    TamLetraSTKDespacho = 10
    
    Dim fso As Object
      
    'Instanciar el objeto FSO para poder _
     usar las funciones FileExists y FolderExists
    Set fso = CreateObject("Scripting.FileSystemObject")
      
    ' Comprobar archivo
    If fso.FileExists(Trim(App.Path) & "\Imagen\PieFactura.jpg") = True Then
        pic.Picture = LoadPicture(Trim(App.Path) & "\Imagen\PieFactura.jpg")
    End If
    
    If optImpresora.Value = True Then
        If strReporte <> "rptSTKDespacho" Then
            DefinirImpresoraPorDefecto ImpresoraTicket
        Else
            DefinirImpresoraPorDefecto ImpresoraEtiqueta
        End If
    End If
'    DoEvents
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    If strReporte = "rptFacturaSola" Then
        strSql = " SELECT DISTINCT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp," & _
                 " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
        strSql = strSql & " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo "
        clsCon_Def.Ejecutar strSql
        While Not clsCon_Def.adorec_Def.EOF
        
            Me.Caption = "Impresion Factura - " & strNumero
        
            
            strSql = " SELECT COALESCE(doc_ele_claveacceso,'') as doc_ele_claveacceso,COALESCE(doc_ele_autorizacion,'') as doc_ele_autorizacion " & _
                     " FROM doc_electronico " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_ele_coddoc='01' " & _
                     " ANd doc_ele_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "'"
            
            clsConAUX.Ejecutar strSql
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                CDEClaveAcceso = Replace(code128$(clsConAUX.adorec_Def("doc_ele_claveacceso")), "'", "''")
                ClaveAcceso = clsConAUX.adorec_Def("doc_ele_claveacceso")
                Autori = clsConAUX.adorec_Def("doc_ele_autorizacion")
            Else
                CDEClaveAcceso = ""
            End If
            
            strSqlAux = " SELECT egreso.emp_codigo,COALESCE(ped_codigo,'0') as ped_codigo,COALESCE(sum(det_egr_cantidad),0) as n,COALESCE(ped_direccion_envio,'') as dir " & _
                       " FROM egreso LEFT JOIN pedido " & _
                       " ON egreso.emp_codigo=pedido.emp_codigo" & _
                       " AND egreso.tip_egr_codigo=pedido.ped_tip_egr_codigo" & _
                       " AND egreso.egr_codigo=pedido.ped_egr_codigo AND pedido.ped_estado IN (2,8) " & _
                       " AND pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_tip_egr_codigo='FAC' " & _
                       " AND pedido.ped_egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "'" & _
                       " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                       " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                       " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                       " AND det_egreso.prd_codigo!='PR-CARGOO100330TU' " & _
                       " AND det_egreso.emp_codigo='" & strEmpresa & "' AND det_egreso.tip_egr_codigo='FAC' " & _
                       " AND det_egreso.egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "'" & _
                       " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo='FAC' " & _
                       " AND egreso.egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "' " & _
                       " GROUP BY egreso.emp_codigo,egr_total,pedido.ped_codigo,ped_direccion_envio"
            clsConAUX.Ejecutar strSqlAux
            dirEnvio = UCase(clsConAUX.adorec_Def("dir"))
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                canti = FormatoD2(clsConAUX.adorec_Def("n"))
            Else
                canti = 0
            End If
                
            strSql = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp ,orden," & _
                     " emp_nombre,emp_direccion,emp_telf,emp_ruc,emp_contribuyenteespecial,persona.tip_ped_codigo," & _
                     " egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                     " egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf," & _
                     " ciu_nombre as ciudad,vendedor.ven_codigo as ven," & _
                     " CONCAT(ven_apellido,' ',ven_nombre) as vendedor,egr_dcto," & _
                     " de.prd_codigo as prd_codigo,de.nombre as nombre,de.prd_ubica_linea,ROUND(cantidad,2) as cantidad," & _
                     " ROUND(det_egr_precio,3) as det_egr_precio,utot,egr_subtotal," & _
                     " egr_dcto,egr_subtotal_o,egr_impuesto,det_egr_dcto,egr_total," & _
                     " CONCAT('Observaciones: ',egreso.egr_observacion) as egr_observacion," & _
                     " cod_iva_porcentaje as Piva,de.mar_codigo,de.gru_codigo,de.gru_nombre," & _
                     " CONCAT('Nº de Factura: ',FORMAT(RIGHT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 7),3)*1,'000'),'-',FORMAT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 10)*1,'000'),'-',FORMAT(Right(egreso.egr_codigo, 7)*1,'000000000'),' - ',FORMAT(current_timestamp,'HH:MM'),persona.cat_p_codigo) as todo," & _
                     " FORMAT(egr_fecha,'yyyy-MM-dd') as fech,for_ent_nombre,"
            strSql = strSql & " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')))>2,CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')))>2,CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')))>2,CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,''))," & _
                     " IIF(LEN(CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')))>2,CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')),''))))))))) as lider,"
            strSql = strSql & " egr_fechamod as fechamod, egr_usumod as usumod," & _
                     " persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9, " & _
                     " CONCAT('Estimados. De acuerdo con la información registrada en nuestro sistema, en tu  mail: - ',persona.per_email,' - recibirás tus documentos electrónicos autorizada por el SRI, según las nueva ley en vigencia. ',IIF(persona.tip_ped_codigo='JON','Si no tienes actualizados tus datos comunicate al 1800CATALOGOS para pedir esta actualización','')) as mensaje, " & _
                     " IIF('" & Autori & "'='','DOCUMENTO SIN VALIDEZ TRIBUTARIA','') as mensaje2 "
            strSql = strSql & " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                     " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                     " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                     " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                     " INNER JOIN codigo_iva ON egreso.cod_iva_codigo=codigo_iva.cod_iva_codigo" & _
                     " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo " & _
                     " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                     " INNER JOIN ("
            strSql = strSql & " SELECT '0' AS orden,det_egreso.emp_codigo,det_egreso.tip_egr_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo, " & _
                     " ROUND(det_egr_cantidad,2) AS cantidad," & _
                     " ROUND(det_egr_precio,3) AS det_egr_precio,det_egr_precio*(det_egr_cantidad) AS utot, " & _
                     " COALESCE(det_egr_pdcto,(det_egr_dcto/det_egr_cantidad* det_egr_cantidad)) AS det_egr_dcto, " & _
                     " IIF(det_egr_pdcto IS NOT NULL,'%','') AS s,prd_nombre AS nombre,mar_codigo,producto.gru_codigo,gru_nombre,det_egreso.prd_ubica_linea " & _
                     " FROM det_egreso_ubicacion det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                     " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
            strSql = strSql & " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                     " AND det_egreso.tip_egr_codigo='FAC' " & _
                     " AND det_egreso.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ")"
            strSql = strSql & " UNION SELECT '2' as orden,det_egreso_c.emp_codigo,det_egreso_c.tip_egr_codigo,det_egreso_c.egr_codigo,det_egreso_c.oca_codigo, " & _
                     " ROUND(det_egr_c_cantidad,2) as cantidad,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,'0.0000' as det_egr_dcto, " & _
                     " '' as s,oca_nombre as nombre,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,'' as prd_ubica_linea " & _
                     " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                     " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                     " AND det_egreso_c.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ")"
            strSql = strSql & ") de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo "
            strSql = strSql & " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                     " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p2.per_es_di=1 " & _
                     " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                     " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                     " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                     " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
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
                     " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                     " AND egreso.tip_egr_codigo='FAC' " & _
                     " AND egreso.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ") " & _
                     " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                     " egr_codigo,orden,prd_ubica_linea ,mar_codigo,LEFT(gru_codigo,2),gru_nombre,nombre  "
            clsConAUX.Ejecutar strSql
            
            TipoPed = clsConAUX.adorec_Def("tip_ped_codigo")
            
            strLinea = Left(clsConAUX.adorec_Def("emp_nombre"), 57)
            ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraFactura, True
            strLinea = Left("R.U.C. " & clsConAUX.adorec_Def("emp_ruc"), 57)
            ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraFactura, True
            strLinea = Left("Dir.Matriz: " & clsConAUX.adorec_Def("emp_direccion"), 57)
            ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraFactura, True
            strLinea = Left("Telf: " & clsConAUX.adorec_Def("emp_telf"), 57)
            ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraFactura, True
            If Trim(clsConAUX.adorec_Def("emp_contribuyenteespecial")) <> "" Then
                strLinea = Left("Contribuyente Especial: NAC-PCTRSGE12-0018", 57)
                ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraFactura, True
            End If
            ImprimeLinea "", "Monaco", TamLetraFactura, True
            ImprimeLinea "Factura No.: " & Format(Mid(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 9, 3), "000") & "-" & Format(Left(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 10), "000") & "-" & Format(Right(clsConAUX.adorec_Def("egr_codigo"), 7), "000000000"), "Monaco", TamLetraFactura, True
            ImprimeLinea "Autorizacion:", "Monaco", TamLetraFactura
            ImprimeLinea Autori, "Monaco", TamLetraFactura
            ImprimeLinea "Clave Acceso:", "Monaco", TamLetraFactura
            ImprimeLinea ClaveAcceso, "Monaco", TamLetraFactura
            ImprimeLinea CDEClaveAcceso, "Code 128", 26
            ImprimeLinea "", "Monaco", TamLetraFactura
            ImprimeLinea "Fecha: " & clsConAUX.adorec_Def("fech"), "Monaco", TamLetraFactura
            ImprimeLinea "Cliente: " & clsConAUX.adorec_Def("per"), "Monaco", TamLetraFactura
            ImprimeLinea "CI/RUC: " & clsConAUX.adorec_Def("per_ruc"), "Monaco", TamLetraFactura
            ImprimeLinea "Telf: " & clsConAUX.adorec_Def("per_telf"), "Monaco", TamLetraFactura
            ImprimeLinea "Dir: " & clsConAUX.adorec_Def("per_direccion"), "Monaco", TamLetraFactura
            
            ImprimeLinea " Codigo      Descripcion          Cant  P.Unit    Total  ", "Monaco", TamLetraFactura
           'ImprimeLinea " Codigo |    Descripcion       | Cant | P.Unit |  Total  ", "Monaco", TamLetraFactura
           LineasEnBlanco = 0
            While Not clsConAUX.adorec_Def.EOF
               'ImprimeLinea " Codigo |     Descripcion      |  Cant  | P.Unit |  Total  ", "Monaco", TamLetraFactura
                ImprimeLinea Rpad(Right(clsConAUX.adorec_Def("prd_codigo"), 7), " ", 7) & " " & _
                             Rpad(clsConAUX.adorec_Def("nombre"), " ", 24) & " " & _
                             Lpad(FormatoD0(clsConAUX.adorec_Def("cantidad")), " ", 4) & " " & _
                             Lpad(FormatoD2(clsConAUX.adorec_Def("det_egr_precio")), " ", 8) & " " & _
                             Lpad(FormatoD2(clsConAUX.adorec_Def("utot")), " ", 9), "Monaco", TamLetraFactura
                clsConAUX.adorec_Def.MoveNext
                LineasEnBlanco = LineasEnBlanco + 1
            Wend
            
            clsConAUX.adorec_Def.MoveLast
            ImprimeLinea "_________________________________________________________", "Monaco", TamLetraFactura
            ImprimeLinea "CANTIDAD:" & Rpad(canti, " ", 5) & "                       SUMA:" & Lpad(clsConAUX.adorec_Def("egr_subtotal") + clsConAUX.adorec_Def("egr_subtotal_o"), " ", 14), "Monaco", TamLetraFactura
            ImprimeLinea "                                     DCTO:" & Lpad(clsConAUX.adorec_Def("egr_dcto"), " ", 14), "Monaco", TamLetraFactura
            ImprimeLinea "                             BASE IVA " & Lpad(FormatoD0(clsConAUX.adorec_Def("Piva")), " ", 2) & "%:" & Lpad(clsConAUX.adorec_Def("egr_subtotal") - clsConAUX.adorec_Def("egr_dcto"), " ", 14), "Monaco", TamLetraFactura
            ImprimeLinea "                                  IVA " & Lpad(FormatoD0(clsConAUX.adorec_Def("Piva")), " ", 2) & "%:" & Lpad(clsConAUX.adorec_Def("egr_impuesto"), " ", 14), "Monaco", TamLetraFactura
            ImprimeLinea "                                    TOTAL:" & Lpad(clsConAUX.adorec_Def("egr_total"), " ", 14), "Monaco", TamLetraFactura
            TOTALPEDIDO = clsConAUX.adorec_Def("egr_total")
            ImprimeLinea "", "Monaco", TamLetraFactura
            
            If TipoPed = "JON" Then
                ImprimeLinea "Lider:" & clsConAUX.adorec_Def("lider"), "Monaco", TamLetraFactura
                ImprimeLinea "Forma de Entrega:" & clsConAUX.adorec_Def("for_ent_nombre"), "Monaco", TamLetraFactura
            End If
            ImprimeLinea clsConAUX.adorec_Def("egr_observacion"), "Monaco", TamLetraFactura
            ImprimeLinea "", "Monaco", TamLetraFactura
            ImprimeLinea "Recibí conforme y acepto el importe de la presente factura comercial, obligándome al pago de la misma sin protesto. Declaro que la relación generada en virtud del presente documento es meramente mercantil. En caso de mora pagaré la tasa máxima autorizada por el emisor.", "Monaco", TamLetraFactura
            ImprimeLinea "", "Monaco", TamLetraFactura
            ImprimeLinea clsConAUX.adorec_Def("mensaje"), "Monaco", TamLetraFactura
            ImprimeLinea "", "Monaco", TamLetraFactura
            ImprimeLinea clsConAUX.adorec_Def("usumod") & " " & clsConAUX.adorec_Def("fechamod"), "Monaco", TamLetraFactura
            ImprimeLinea "", "Monaco", TamLetraFactura
            ImprimeLinea "", "Monaco", TamLetraFactura
            If 5 - LineasEnBlanco > 0 Then
                For i = 1 To 5 - LineasEnBlanco
                    ImprimeLinea "", "Monaco", TamLetraPedido
                Next i
            End If
            LineasEnBlanco = 0
            
            'PEDIDO REPROGRAMADO
            strSql = " SELECT * " & _
                     " FROM Fn_DetFacturaReprogamada('" & strEmpresa & "'," & strNumero & ")"
                 
            clsConAUX.Ejecutar strSql
            
            If clsConAUX.adorec_Def.RecordCount > 0 Then
            
                 ImprimeLinea "_________________________________________________________", "Monaco", TamLetraFactura
                 ImprimeLinea "", "Monaco", TamLetraFactura, True
            
                 ImprimeLinea "REPROGRAMACION DE PEDIDO", "Monaco", TamLetraFactura
                 
                 ImprimeLinea "", "Monaco", TamLetraFactura, True
                 
                 ImprimeLinea " Codigo      Descripcion          Cant  P.Unit    Total  ", "Monaco", TamLetraFactura
                'ImprimeLinea " Codigo |    Descripcion       | Cant | P.Unit |  Total  ", "Monaco", TamLetraFactura
                 LineasEnBlanco = 0
                 
                 cantiRepro = 0
                 subTotalRepro = 0
                 dctoRepro = 0
                 IvaRepro = 0
                 'TOTALPEDIDO = 0
                 While Not clsConAUX.adorec_Def.EOF
                    'ImprimeLinea " Codigo |     Descripcion      |  Cant  | P.Unit |  Total  ", "Monaco", TamLetraFactura
                     ImprimeLinea "Pedido: " & clsConAUX.adorec_Def("det_ped_ped_reprogramado"), "Monaco", TamLetraFactura
                     ImprimeLinea Rpad(Right(clsConAUX.adorec_Def("prd_codigo"), 7), " ", 7) & " " & _
                                  Rpad(clsConAUX.adorec_Def("nombre"), " ", 24) & " " & _
                                  Lpad(FormatoD0(clsConAUX.adorec_Def("cantidad")), " ", 4) & " " & _
                                  Lpad(FormatoD2(clsConAUX.adorec_Def("det_ped_precio")), " ", 8) & " " & _
                                  Lpad(FormatoD2(clsConAUX.adorec_Def("utot")), " ", 9), "Monaco", TamLetraFactura
                     
                     LineasEnBlanco = LineasEnBlanco + 1
                     cantiRepro = cantiRepro + FormatoD0(clsConAUX.adorec_Def("cantidad"))
                     subTotalRepro = subTotalRepro + FormatoD2(clsConAUX.adorec_Def("utot"))
                     dctoRepro = dctoRepro + FormatoD2(clsConAUX.adorec_Def("dcto"))
                     clsConAUX.adorec_Def.MoveNext
                 Wend
                 
                 clsConAUX.adorec_Def.MoveLast
                 ImprimeLinea "_________________________________________________________", "Monaco", TamLetraFactura
                 ImprimeLinea "CANTIDAD:" & Rpad(cantiRepro, " ", 5) & "                SUMA Repro.:" & Lpad(FormatoD2(subTotalRepro), " ", 14), "Monaco", TamLetraFactura
                 ImprimeLinea "                              DCTO Repro.:" & Lpad(FormatoD2(dctoRepro), " ", 14), "Monaco", TamLetraFactura
                 ImprimeLinea "                      BASE IVA " & Lpad(FormatoD0(clsConAUX.adorec_Def("Piva")), " ", 2) & "% Repro.:" & Lpad(FormatoD2(FormatoD2(subTotalRepro) - FormatoD2(dctoRepro)), " ", 14), "Monaco", TamLetraFactura
                 ImprimeLinea "                           IVA " & Lpad(FormatoD0(clsConAUX.adorec_Def("Piva")), " ", 2) & "% Repro.:" & Lpad(FormatoD2(subTotalRepro * 0.12), " ", 14), "Monaco", TamLetraFactura
                 ImprimeLinea "                             TOTAL Repro.:" & Lpad(FormatoD2(FormatoD2(subTotalRepro) - FormatoD2(dctoRepro) + FormatoD2(subTotalRepro * 0.12)), " ", 14), "Monaco", TamLetraFactura
                 TOTALPEDIDO = TOTALPEDIDO + FormatoD2(FormatoD2(subTotalRepro) - FormatoD2(dctoRepro) + FormatoD2(subTotalRepro * 0.12))
                 ImprimeLinea "", "Monaco", TamLetraFactura
                 ImprimeLinea "", "Monaco", TamLetraFactura
                 ImprimeLinea "        TOTAL PEDIDO:" & Lpad(FormatoD2(TOTALPEDIDO), " ", 14), "Monaco", TamLetraFactura, True
                 
                 ImprimeLinea "", "Monaco", TamLetraFactura
            End If
            
            If fso.FileExists(Trim(App.Path) & "\Imagen\PieFactura.jpg") = True Then
                If optImpresora.Value = True Then
                    Printer.PaintPicture pic, 0, Printer.CurrentY
                End If
            End If
            
            TerminarHoja
            
            clsCon_Def.adorec_Def.MoveNext
        Wend
    ElseIf strReporte = "rptPedido" Then
        
        strSql = " SELECT DISTINCT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp," & _
                 " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
        strSql = strSql & " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo "
        clsCon_Def.Ejecutar strSql
        
        lngPagina = 0
        While Not clsCon_Def.adorec_Def.EOF
            lngPagina = lngPagina + 1
            Me.Caption = "Impresion Pedido - " & strNumero
            DoEvents
            
            If lngPag = 0 Or lngPag <= lngPagina Then
            
                 strSql = " SELECT COALESCE(ped_codigo,'0') as ped_codigo,COALESCE(sum(det_egr_cantidad),0) as n,COALESCE(ped_direccion_envio,'') as dir " & _
                          " FROM egreso LEFT JOIN pedido " & _
                          " ON egreso.emp_codigo=pedido.emp_codigo" & _
                          " AND egreso.tip_egr_codigo=pedido.ped_tip_egr_codigo" & _
                          " AND egreso.egr_codigo=pedido.ped_egr_codigo AND pedido.ped_estado IN (2,8) " & _
                          " AND pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_tip_egr_codigo='FAC' " & _
                          " AND pedido.ped_egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "'" & _
                          " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                          " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                          " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                          " AND det_egreso.prd_codigo!='PR-CARGOO100330TU' " & _
                          " AND det_egreso.emp_codigo='" & strEmpresa & "' AND det_egreso.tip_egr_codigo='FAC' " & _
                          " AND det_egreso.egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "'" & _
                          " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo='FAC' " & _
                          " AND egreso.egr_codigo='" & clsCon_Def.adorec_Def("egr_codigo") & "' " & _
                          " GROUP BY pedido.ped_codigo,ped_direccion_envio"
                 
                 clsConAUX.Ejecutar strSql
                 dirEnvio = UCase(clsConAUX.adorec_Def("dir"))
                 If clsConAUX.adorec_Def.RecordCount > 0 Then
                     CDEPedido = Replace(code128$(clsConAUX.adorec_Def("ped_codigo")), "'", "''")
                     Pedido = clsConAUX.adorec_Def("ped_codigo")
                     canti = clsConAUX.adorec_Def("n")
                 Else
                     CDEPedido = ""
                     Pedido = ""
                     canti = 0
                 End If
                 
                 strSql = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp ,orden," & _
                          " egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                          " de.prd_codigo as prd_codigo,de.nombre as nombre,de.prd_ubica_linea,ROUND(cantidad,2) as cantidad," & _
                          " de.mar_codigo,de.gru_codigo,de.gru_nombre,egr_observacion,IIF('" & dirEnvio & "'!='','" & dirEnvio & "',per_direccion2) as direnvio," & _
                          " FORMAT(egr_fecha,'yyyy-MM-dd') as fech,for_ent_nombre,"
                 strSql = strSql & " egr_fechamod as fechamod, egr_usumod as usumod "
                 strSql = strSql & " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                          " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                          " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo " & _
                          " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                          " INNER JOIN ("
                 strSql = strSql & " SELECT '0' AS orden,det_egreso.emp_codigo,det_egreso.tip_egr_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo, " & _
                          " ROUND(det_egr_cantidad,2) AS cantidad," & _
                          " ROUND(det_egr_precio,3) AS det_egr_precio,det_egr_precio*(det_egr_cantidad) AS utot, " & _
                          " COALESCE(det_egr_pdcto,(det_egr_dcto/det_egr_cantidad* det_egr_cantidad)) AS det_egr_dcto, " & _
                          " IIF(det_egr_pdcto IS NOT NULL,'%','') AS s,prd_nombre AS nombre,mar_codigo,producto.gru_codigo,gru_nombre,det_egreso.prd_ubica_linea " & _
                          " FROM det_egreso_ubicacion det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                          " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
                 strSql = strSql & " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                          " AND det_egreso.tip_egr_codigo='FAC' " & _
                          " AND det_egreso.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ")"
                 strSql = strSql & " UNION SELECT '2' as orden,det_egreso_c.emp_codigo,det_egreso_c.tip_egr_codigo,det_egreso_c.egr_codigo,det_egreso_c.oca_codigo, " & _
                          " ROUND(det_egr_c_cantidad,2) as cantidad,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,'0.0000' as det_egr_dcto, " & _
                          " '' as s,oca_nombre as nombre,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,'' as prd_ubica_linea " & _
                          " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                          " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                          " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                          " AND det_egreso_c.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ")"
                 strSql = strSql & ") de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo "
                 strSql = strSql & " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                          " AND egreso.tip_egr_codigo='FAC' " & _
                          " AND egreso.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ") " & _
                          " ORDER BY ordenfp,persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9," & _
                          " egr_codigo,orden,prd_ubica_linea ,mar_codigo,LEFT(gru_codigo,2),gru_nombre,nombre  "
                 clsConAUX.Ejecutar strSql
                 ImprimeLinea "P-" & lngPagina, "Monaco", TamLetraSTKDespacho, True
                 strLinea = Left("Factura No.: " & Format(Mid(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 9, 3), "000") & "-" & Format(Left(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 10), "000") & "-" & Format(Right(clsConAUX.adorec_Def("egr_codigo"), 7), "000000000"), 57)
                 ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraPedido, True
                 ImprimeLinea CDEPedido, "Code 128", 36
                 ImprimeLinea "Pedido: " & Pedido, "Monaco", TamLetraPedido
                 ImprimeLinea "Fecha: " & clsConAUX.adorec_Def("fech"), "Monaco", TamLetraPedido
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 ImprimeLinea "   Codigo             Ubicacion       Cant ", "Monaco", TamLetraPedido
                'ImprimeLinea "   Codigo       |     Ubicacion     | Cant ", "Monaco", TamLetraPedido
                 LineasEnBlanco = 0
                 While Not clsConAUX.adorec_Def.EOF
                    'ImprimeLinea "   Ubicacion     |Cant|     Descripcion         ", "Monaco", TamLetraPedido
                     ImprimeLinea Rpad(clsConAUX.adorec_Def("prd_codigo"), " ", 16) & " " & _
                                  Rpad(clsConAUX.adorec_Def("prd_ubica_linea"), " ", 19) & " " & _
                                  Lpad(FormatoD0(clsConAUX.adorec_Def("cantidad")), " ", 4), "Monaco", TamLetraPedido
                     ImprimeLinea Rpad(clsConAUX.adorec_Def("nombre"), " ", 25), "Monaco", TamLetraPedido
                     clsConAUX.adorec_Def.MoveNext
                     LineasEnBlanco = LineasEnBlanco + 1
                 Wend
                 clsConAUX.adorec_Def.MoveLast
                 ImprimeLinea "_______________________________________________", "Monaco", TamLetraPedido
                'ImprimeLinea "   Codigo       |     Ubicacion     | Cant ", "Monaco", TamLetraPedido
                 ImprimeLinea "  TOTAL UNIDADES: " & canti, "Monaco", TamLetraPedido, True
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 
                 ImprimeLinea "Forma de Entrega:" & clsConAUX.adorec_Def("for_ent_nombre"), "Monaco", TamLetraPedido
                 ImprimeLinea "Dir-Envio:" & clsConAUX.adorec_Def("direnvio"), "Monaco", TamLetraPedido
                 
                 ImprimeLinea clsConAUX.adorec_Def("egr_observacion"), "Monaco", TamLetraPedido
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 ImprimeLinea clsConAUX.adorec_Def("usumod") & " " & clsConAUX.adorec_Def("fechamod"), "Monaco", TamLetraPedido
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 
                 strSql = " SELECT orden,de.prd_codigo as prd_codigo,de.nombre as nombre,ROUND(cantidad,2) as cantidad " & _
                          " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
                          " AND egreso.per_codigo=persona.per_codigo " & _
                          " INNER JOIN ("
                 strSql = strSql & " SELECT '0' as orden,pedido.emp_codigo,pedido.ped_tip_egr_codigo as tip_egr_codigo," & _
                          " pedido.ped_egr_codigo as egr_codigo,'' as prd_codigo, 0 as cantidad, " & _
                          " 'VACIO' as nombre" & _
                          " FROM pedido " & _
                          " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                          " AND pedido.ped_tip_egr_codigo='FAC' " & _
                          " AND pedido.ped_egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ") " & _
                          " UNION" & _
                          " SELECT '1' as orden,pedido.emp_codigo,pedido.ped_tip_egr_codigo as tip_egr_codigo," & _
                          " pedido.ped_egr_codigo as egr_codigo,det_pedido.prd_codigo, ROUND(det_ped_cant_entregada,2) as cantidad," & _
                          " prd_nombre as nombre" & _
                          " FROM pedido INNER JOIN det_pedido " & _
                          " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                          " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                          " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo " & _
                          " AND det_pedido.prd_codigo=producto.prd_codigo " & _
                          " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                          " AND pedido.ped_tip_egr_codigo='FAC' " & _
                          " AND pedido.ped_egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ") " & _
                          " AND det_pedido.det_ped_incentivo=1 " & _
                          " AND ROUND(det_ped_cant_entregada,2)!=0 "
                 strSql = strSql & " ) de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo " & _
                          " AND egreso.egr_codigo=de.egr_codigo " & _
                          " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                          " AND egreso.tip_egr_codigo='FAC' " & _
                          " AND egreso.egr_codigo in (" & clsCon_Def.adorec_Def("egr_codigo") & ") " & _
                          " ORDER BY orden DESC,nombre"
                 
                 clsConAUX.Ejecutar strSql
                 strLinea = Left("-- ADICIONALES --", 57)
                 ImprimeLinea Lpad(strLinea, " ", 24 + (FormatoD0(Len(strLinea) / 2))), "Monaco", TamLetraPedido, True
                 
                 ImprimeLinea "    Nombre                         Cant  ", "Monaco", TamLetraPedido
                'ImprimeLinea "    Codigo    |     Descripcio  |  Cant  ", "Monaco", TamLetrapedido
                 canti = 0
                 While Not clsConAUX.adorec_Def.EOF
                     If FormatoD0(clsConAUX.adorec_Def("cantidad")) <> 0 And FormatoD0(clsConAUX.adorec_Def("orden")) <> 0 Then
                         ImprimeLinea Rpad(Right(clsConAUX.adorec_Def("nombre"), 30), " ", 30) & " " & _
                                      Lpad(FormatoD0(clsConAUX.adorec_Def("cantidad")), " ", 8), "Monaco", TamLetraPedido
                         canti = canti + FormatoD0(clsConAUX.adorec_Def("cantidad"))
                     End If
                     clsConAUX.adorec_Def.MoveNext
                 Wend
                 clsConAUX.adorec_Def.MoveLast
                 ImprimeLinea "________________________________________________", "Monaco", TamLetraPedido
                'ImprimeLinea "    Codigo    |     Descripcion     |Cant| Ubicacion       ", "Monaco", TamLetrapedido
                 ImprimeLinea "  TOTAL UNIDADES: " & canti, "Monaco", TamLetraPedido, True
                 If 5 - LineasEnBlanco > 0 Then
                     For i = 1 To 5 - LineasEnBlanco
                         ImprimeLinea "", "Monaco", TamLetraPedido
                     Next i
                 End If
                 ImprimeLinea "", "Monaco", TamLetraPedido
                 ImprimeLinea "Impr:" & Now(), "Monaco", TamLetraPedido
                 LineasEnBlanco = 0
                 TerminarHoja
            End If
            clsCon_Def.adorec_Def.MoveNext
        Wend
    ElseIf strReporte = "rptSTKDespacho" Then
        strSql = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)IN ('EFE','CONT'),1,0) as ordenfp,pedido.ped_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " IIF(persona.for_pag_codigo NOT IN ('EFE','CONT') OR pedido.ped_direccion_envio IS NULL or pedido.ped_direccion_envio='' OR LEFT(pedido.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,'-',persona.per_direccion2,' (',for_ent_nombre,')'),CONCAT(pedido.ped_direccion_envio,' (',for_ent_nombre,')')) as per_direccion2,persona.per_direccion,CONCAT(persona.per_telf,'/',persona.per_fax,'/',persona.per_celular) as per_tfc,persona.per_codigo_postal," & _
                 " ped_fecha,for_pag_nombre,COALESCE(dis_pol_nombre,'') as dis_pol_nombre," & _
                 " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1, " & _
                 " CURRENT_TIMESTAMP as hoy, " & _
                 " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')),''))))))))) as papa"
        strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
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
                 " LEFT JOIN distribucion_politica ON persona.dis_pol_codigo=distribucion_politica.dis_pol_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9,pedido.ped_codigo "
        clsCon_Def.Ejecutar strSql
        While Not clsCon_Def.adorec_Def.EOF
            CDEPedido = Replace(code128$(clsCon_Def.adorec_Def("ped_codigo")), "'", "''")
            ImprimeLinea CDEPedido, "Code 128", 36
            ImprimeLinea "Pedido: " & clsCon_Def.adorec_Def("ped_codigo"), "Monaco", TamLetraSTKDespacho
            ImprimeLinea "Cliente: " & clsCon_Def.adorec_Def("per"), "Monaco", TamLetraSTKDespacho
            ImprimeLinea "Envio: " & clsCon_Def.adorec_Def("per_direccion2"), "Monaco", TamLetraSTKDespacho, True
            ImprimeLinea "Status:", "Monaco", TamLetraSTKDespacho
            ImprimeLinea clsCon_Def.adorec_Def("for_pag_nombre"), "Monaco", 15, True
            ImprimeLinea ""
            ImprimeLinea "Lider: " & clsCon_Def.adorec_Def("papa"), "Monaco", TamLetraSTKDespacho
            ImprimeLinea "N1: " & clsCon_Def.adorec_Def("nn1"), "Monaco", TamLetraSTKDespacho
            ImprimeLinea clsCon_Def.adorec_Def("hoy"), "Monaco", TamLetraSTKDespacho
            TerminarHoja
            clsCon_Def.adorec_Def.MoveNext
        Wend
    ElseIf strReporte = "rptListaEmbarque" Then
    
        strSql = " SELECT empresa.emp_nombre, contenedor.con_codigo,con_fecha,con_observacion,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " con_guia,con_peso,cou_nombre,CAST(CHARINDEX(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre,' - ',persona.per_direccion2),d.ped_direccion_envio),CHARINDEX(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre,' - ',persona.per_direccion2),d.ped_direccion_envio),' - ')+3,500) as varchar) as per_direccion2,persona.per_telf,IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre),LEFT(d.ped_direccion_envio,CHARINDEX(d.ped_direccion_envio,' - '))) as ciu_nombre,zona.zon_nombre, " & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet," & _
                 " CAST(egreso.egr_codigo as varchar) as egr_codigo,pedido.ped_codigo,con_usumod,egr_total,sum(det_egr_cantidad) as egr_unidades " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo AND contenedor.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo " & _
                 " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                 " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo "
        strSql = strSql & " INNER JOIN egreso ON det_contenedor.emp_codigo=egreso.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_contenedor.tip_egr_codigo AND egreso.egr_anulado=0 " & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND det_egreso.prd_codigo NOT LIKE 'PR-%'" & _
                 " INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona pd ON egreso.emp_codigo=pd.emp_codigo " & _
                 " AND egreso.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo "
        strSql = strSql & " LEFT JOIN (SELECT TOP 1 pedido.emp_codigo,pedido.per_codigo,pedido.ped_direccion_envio " & _
                 " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo " & _
                 " AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.for_pag_codigo IN ('CONT','EFE')" & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & strNumero & "') as d" & _
                 " ON persona.emp_codigo=d.emp_codigo " & _
                 " AND persona.per_codigo=d.per_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' GROUP BY empresa.emp_nombre,contenedor.con_codigo,con_fecha,con_observacion,persona.per_nombre,persona.per_apellido,persona.per_telf, con_guia,con_peso,cou_nombre,d.ped_direccion_envio,ciudad.ciu_nombre,canton.can_nombre,pais.pai_nombre,persona.per_direccion2, zona.zon_nombre,pd.per_apellido,pd.per_nombre,cd.ciu_nombre,zd.zon_nombre, egreso.egr_codigo,pedido.ped_codigo,con_usumod,egr_total"
        strSql = strSql & " UNION SELECT empresa.emp_nombre, contenedor.con_codigo,con_fecha,con_observacion,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " con_guia,con_peso,cou_nombre,persona.per_direccion2,persona.per_telf,ciudad.ciu_nombre,zona.zon_nombre, " & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet," & _
                 " det_contenedor_per.det_con_per_detalle,'0' as ped_codigo,con_usumod,0,1 " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo AND contenedor.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_contenedor_per ON contenedor.emp_codigo=det_contenedor_per.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_per.con_codigo " & _
                 " INNER JOIN persona pd ON det_contenedor_per.emp_codigo=pd.emp_codigo " & _
                 " AND det_contenedor_per.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' "
        clsCon_Def.Ejecutar strSql
        ImprimeLinea "LISTA DE EMPAQUE", "Monaco", TamLetraSTKDespacho, True
        ImprimeLinea ""
        ImprimeLinea "Courier: " & clsCon_Def.adorec_Def("cou_nombre"), "Monaco", TamLetraSTKDespacho, True
        ImprimeLinea "Fecha: " & Format(clsCon_Def.adorec_Def("con_fecha"), "yyyy-mm-dd"), "Monaco", TamLetraSTKDespacho, True
        ImprimeLinea "No.Guia: " & clsCon_Def.adorec_Def("con_guia"), "Monaco", TamLetraSTKDespacho, True
        ImprimeLinea "No.Contenedor: " & clsCon_Def.adorec_Def("con_codigo"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea ""
        ImprimeLinea "Lider: " & clsCon_Def.adorec_Def("per"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea "Dir: " & clsCon_Def.adorec_Def("per_direccion2"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea "Telf: " & clsCon_Def.adorec_Def("per_telf"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea "Ciudad: " & clsCon_Def.adorec_Def("ciu_nombre"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea "Zona: " & clsCon_Def.adorec_Def("zon_nombre"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea "Obs: " & clsCon_Def.adorec_Def("con_observacion"), "Monaco", TamLetraSTKDespacho
        ImprimeLinea ""
        ImprimeLinea "Cliente", "Monaco", TamLetraPedido
        ImprimeLinea "    Pedido        Factura       Ciudad     Valor    Uni. ", "Monaco", TamLetraFactura
        ImprimeLinea "_________________________________________________________", "Monaco", TamLetraFactura
        'ImprimeLinea "    Pedido    |   Factura    |  Ciudad  |  Valor  | Uni. ", "Monaco", TamLetraFactura
        i = 0
        While Not clsCon_Def.adorec_Def.EOF
            ImprimeLinea clsCon_Def.adorec_Def("perdet"), "Monaco", TamLetraPedido
            ImprimeLinea Rpad(Right(clsCon_Def.adorec_Def("ped_codigo"), 14), " ", 14) & " " & _
                         Rpad(Right(clsCon_Def.adorec_Def("egr_codigo"), 14), " ", 14) & " " & _
                         Rpad(Right(clsCon_Def.adorec_Def("ciudet"), 10), " ", 10) & " " & _
                         Lpad(clsCon_Def.adorec_Def("egr_total"), " ", 9) & " " & _
                         Lpad(clsCon_Def.adorec_Def("egr_unidades"), " ", 6), "Monaco", TamLetraFactura
            clsCon_Def.adorec_Def.MoveNext
            i = i + 1
        Wend
        ImprimeLinea "_________________________________________________________", "Monaco", TamLetraFactura
        ImprimeLinea "TOTAL PEDIDOS: " & i, "Monaco", TamLetraSTKDespacho, True
        ImprimeLinea ""
        ImprimeLinea "Al momento de recibir el paquete, verificar su contenido y realizar el cuadre de acuerdo a este Listado de Embarque. " & _
                     "Nuestra garantía es de 24 HORAS despues de haber recibido su pedido.", "Monaco", TamLetraFactura
        
        TerminarHoja
    ElseIf strReporte = "rptContenedorMercaderia" Then
    
    End If
    If optImpresora.Value = True Then
        DefinirImpresoraPorDefecto ImpresoraPorDefecto
    End If
    
End Sub

Private Sub TerminarHoja()
    If optImpresora.Value = True Then
        ImprimeLinea "."
        Printer.EndDoc
    ElseIf optPantalla.Value = True Then
        'ImprimeLinea "123456789112345678921234567893123456789412345678951234567", , 10
        ImprimeLinea "_________________________________________________________"
        ImprimeLinea ""
    End If
End Sub

Private Sub ImprimeLinea(strLinea As String, Optional strFont As String = "Monaco", Optional intTama As Integer = 6, Optional booNegrita As Boolean = False, Optional booItalica As Boolean = False)
    Dim ini As Long
    Dim largo As Long
    Dim strParcial As String
    Dim MaxCaracteres As Integer
    If intTama = 6 Then
        MaxCaracteres = 57
    ElseIf intTama = 7 Then
        MaxCaracteres = 48
    ElseIf intTama = 10 Then
        MaxCaracteres = 36
    ElseIf intTama = 15 Then
        MaxCaracteres = 25
    Else
        MaxCaracteres = 10000
    End If
    If optImpresora.Value = True Then
        Printer.FontBold = booNegrita
        Printer.FontSize = intTama
        Printer.FontItalic = booItalica
        Printer.Font = strFont
        While Len(strLinea) > MaxCaracteres
            strParcial = Left(strLinea, MaxCaracteres)
            If InStrRev(strParcial, " ") > 0 Then
                strParcial = Left(strLinea, InStrRev(strParcial, " "))
                strLinea = Mid(strLinea, Len(strParcial) + 1)
            Else
                strParcial = Left(strLinea, MaxCaracteres)
                strLinea = Mid(strLinea, MaxCaracteres + 1)
            End If
            Printer.Print strParcial
        Wend
        Printer.Print strLinea
        While Len(strLinea) > MaxCaracteres
            strParcial = Left(strLinea, MaxCaracteres)
            If InStrRev(strParcial, " ") > 0 Then
                strParcial = Left(strLinea, InStrRev(strParcial, " "))
                strLinea = Mid(strLinea, Len(strParcial) + 1)
            Else
                strParcial = Left(strLinea, MaxCaracteres)
                strLinea = Mid(strLinea, MaxCaracteres + 1)
            End If
            Printer.Print strParcial
        Wend
    Else
        ini = Len(rtxtImpresion.Text) + 2
        largo = Len(strLinea)
        rtxtImpresion.TextRTF = Replace(rtxtImpresion.TextRTF, "\par }", vbNewLine & "\par " & strLinea & "\par }")
        rtxtImpresion.SelStart = ini
        rtxtImpresion.SelLength = largo
        rtxtImpresion.SelBold = booNegrita
        rtxtImpresion.SelFontSize = intTama
        rtxtImpresion.SelItalic = booItalica
        rtxtImpresion.SelFontName = strFont
        rtxtImpresion.Refresh
    End If
'    Me.Refresh
    
End Sub


Private Sub Command1_Click()
    strNumero = "40020409316"
    strReporte = "rptFacturaSola"
'    strReporte = "rptPedido"
    strReporte = "rptSTKDespacho"
    strNumero = "10020936158,10020936159,10020936160"
    
    strReporte = "rptListaEmbarque"
    strNumero = "31388"
    
    
    
    'rtxtImpresion.SaveFile "c:\Documento.rtf"
    
    'Printer.EndDoc
    'rtxtImpresion.FileName = "c:\Documento.rtf"
    'cd.ShowFont
    'MsgBox cd.FontName
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

Public Sub CmdCerrar_Click()
    Unload Me
End Sub

Public Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
     
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

