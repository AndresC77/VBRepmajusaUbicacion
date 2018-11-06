VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEnvioCartera 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de Cartera"
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
   Icon            =   "frmEnvioCartera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13395
   Begin VB.CheckBox chkIncluyeContado 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Incluir Clientes de contado"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   150
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Cartera"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   13200
      Begin VB.CommandButton cmdEnviarCorreos 
         Caption         =   "&Enviar Correos"
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsultaCartera 
         Caption         =   "Consulta Cartera a Cobrar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTotalACobrar 
         Height          =   375
         Left            =   11280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3720
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   12975
         _cx             =   22886
         _cy             =   5318
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEnvioCartera.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFGExportar 
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   12975
         _cx             =   2088786278
         _cy             =   2088766382
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEnvioCartera.frx":0488
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
   End
   Begin MSDataListLib.DataCombo cmbNegocio 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
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
      Left            =   240
      TabIndex        =   1
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmEnvioCartera"
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
Private Destino As String

Private Sub cmdConsultaCartera_Click()
    Dim i As Long
    
    strSql = " INSERT INTO personaT SELECT per_codigo,emp_codigo,for_pag_codigo,per_ruc,per_nombre,per_apellido,tip_ped_codigo,per_codigo_resp,per_email,per_celular " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND cat_p_tipo='C' " & _
             " AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'"
    If chkIncluyeContado.Value = 0 Then
        strSql = strSql & " AND persona.for_pag_codigo NOT IN ('EFE','CONT') "
    End If
    clsCon_Def.Ejecutar strSql
    
    strSql = " INSERT INTO cuenta_p_cT SELECT cuenta_p_c.* " & _
             " FROM cuenta_p_c INNER JOIN personaT ON cuenta_p_c.emp_codigo=personaT.emp_codigo " & _
             " AND cuenta_p_c.per_codigo=personaT.per_codigo " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo = 'C'" & _
             " AND cue_p_c_pagado='0'" & _
             " AND cue_p_c_egr_codigo NOT LIKE 'R%'"
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO pagoT SELECT pago.pag_codigo,pago.emp_codigo,pago.cue_p_c_codigo,pago.cue_p_c_tipo,pago.pag_monto " & _
                 " FROM cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo " & _
                 " INNER JOIN pago ON cuenta_p_cT.emp_codigo=pago.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                 " AND cuenta_p_cT.cue_p_c_codigo=pago.cue_p_c_codigo AND pag_monto!=0 " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO comprobante_retencionT SELECT comprobante_retencion.emp_codigo,comprobante_retencion.cue_p_c_codigo,comprobante_retencion.cue_p_c_tipo,comprobante_retencion.com_ret_total " & _
                 " FROM cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo " & _
                 " INNER JOIN comprobante_retencion ON cuenta_p_cT.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' "
    clsCon_Def.Ejecutar strSql
    strSql = " SELECT CONCAT(Deu.per_nombre, ' ',Deu.per_apellido) as Deudor,Deu.per_email,Deu.per_celular,COALESCE(cco.per_email_cco,''),cuenta_p_cT.cue_p_c_codigo as c1,CONCAT(personaT.per_apellido, ' ',personaT.per_nombre) as cli,personaT.per_ruc, RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, DATEADD(d,for_pag_tiempo,cue_p_c_fechaemision) as cue_p_c_fechapropuesta,cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                 " FROM cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " INNER JOIN personaT Deu ON personaT.emp_codigo=Deu.emp_codigo" & _
                 " AND personaT.per_codigo_resp=Deu.per_codigo AND Deu.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " INNER JOIN forma_pago ON personaT.emp_codigo=forma_pago.emp_codigo " & _
                 " AND personaT.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " LEFT JOIN (" & _
                    " SELECT persona.emp_codigo,persona.per_codigo,li.per_email as per_email_cco" & _
                    " FROM persona_copia_de INNER JOIN persona li ON persona_copia_de.emp_codigo=li.emp_codigo" & _
                    " AND persona_copia_de.per_codigo=li.per_codigo" & _
                    " INNER JOIN persona ON persona_copia_de.emp_codigo=persona.emp_codigo" & _
                    " AND persona_copia_de.per_codigo=persona.per_codigo_ref" & _
                    " WHERE persona_copia_de.emp_codigo='" & strEmpresa & "' AND per_cop_de='CARTERA'" & _
                 ") cco ON personaT.emp_codigo=cco.emp_codigo " & _
                 " AND personaT.per_codigo=cco.per_codigo " & _
                 " LEFT JOIN pagoT ON cuenta_p_cT.emp_codigo=pagoT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=pagoT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=pagoT.cue_p_c_codigo " & _
                 " LEFT JOIN comprobante_retencionT ON cuenta_p_cT.emp_codigo=comprobante_retencionT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=comprobante_retencionT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=comprobante_retencionT.cue_p_c_codigo " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' " & _
                 " AND personaT.for_pag_codigo NOT IN ('EFE','CONT') " & _
                 " GROUP BY Deu.per_apellido,Deu.per_nombre,Deu.per_email,Deu.per_celular,COALESCE(cco.per_email_cco,''),cuenta_p_cT.cue_p_c_codigo,personaT.per_apellido,personaT.per_nombre,personaT.per_ruc,cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, for_pag_tiempo,cue_p_c_fechaemision,cue_p_c_valor ,cue_p_c_valor,com_ret_total " & _
                 " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)>0 "
    strSql = strSql & " UNION " & _
                 " SELECT CONCAT(Deu.per_nombre, ' ',Deu.per_apellido) as Deudor,Deu.per_email,Deu.per_celular,COALESCE(cco.per_email_cco,''),ingreso.ing_codigo as c1,CONCAT(personaT.per_apellido, ' ',personaT.per_nombre) as cli,personaT.per_ruc, RIGHT(ing_codigo,7) as cue_p_c_egr_codigo, 'NOTA CREDITO', ing_fecha, ing_fecha,-1*ing_total ,-1*(ing_total-ing_saldo) as d " & _
                 " FROM ingreso INNER JOIN personaT ON ingreso.emp_codigo=personaT.emp_codigo" & _
                 " AND ingreso.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " INNER JOIN personaT Deu ON personaT.emp_codigo=Deu.emp_codigo" & _
                 " AND personaT.per_codigo_resp=Deu.per_codigo AND Deu.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
                 " INNER JOIN forma_pago ON personaT.emp_codigo=forma_pago.emp_codigo " & _
                 " AND personaT.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " LEFT JOIN (" & _
                    " SELECT persona.emp_codigo,persona.per_codigo,li.per_email as per_email_cco" & _
                    " FROM persona_copia_de INNER JOIN persona li ON persona_copia_de.emp_codigo=li.emp_codigo" & _
                    " AND persona_copia_de.per_codigo=li.per_codigo" & _
                    " INNER JOIN persona ON persona_copia_de.emp_codigo=persona.emp_codigo" & _
                    " AND persona_copia_de.per_codigo=persona.per_codigo_ref" & _
                    " WHERE persona_copia_de.emp_codigo='" & strEmpresa & "' AND per_cop_de='CARTERA'" & _
                 ") cco ON personaT.emp_codigo=cco.emp_codigo " & _
                 " AND personaT.per_codigo=cco.per_codigo " & _
                 " WHERE ingreso.emp_codigo = '" & strEmpresa & "' AND ingreso.tip_ing_codigo = 'DCL' " & _
                 " AND ing_anulado=0 " & _
                 " AND personaT.for_pag_codigo NOT IN ('EFE','CONT') " & _
                 " AND ROUND(ing_total-ing_saldo,2)>0 " & _
                 " ORDER BY Deudor,cue_p_c_fechapropuesta,cli,cue_p_c_fechaemision,cue_p_c_egr_codigo,c1"

    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    txtTotalACobrar.Text = FormatoD2(0)
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 5) = QuitarCaracteresEspecialesYNumeros(VSFG.TextMatrix(i, 5))
        VSFG.TextMatrix(i, 8) = QuitarCaracteresEspeciales(Left(VSFG.TextMatrix(i, 8), 25))
        txtTotalACobrar.Text = FormatoD2(FormatoD2(txtTotalACobrar.Text) + FormatoD2(VSFG.TextMatrix(i, 12)))
    Next i
    strSql = " DELETE FROM pagoT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM comprobante_retencionT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM cuenta_p_cT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM personaT "
    clsCon_Def.Ejecutar strSql
    VSFG.MergeCol(0) = True: VSFG.MergeCol(1) = True: VSFG.MergeCol(2) = True: VSFG.MergeCol(3) = True: VSFG.MergeCol(4) = True
    MsgBox Now
End Sub

Private Function QuitarCaracteresEspecialesYNumeros(cadena As String) As String
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
    QuitarCaracteresEspecialesYNumeros = CadenaFinal
End Function

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
            If (Not (Asc(Caracter) >= 65 And Asc(Caracter) <= 90)) And IsNumeric(Caracter) = False Then
                Caracter = ""
            End If
        End If
        CadenaFinal = CadenaFinal & Caracter
    Next i
    QuitarCaracteresEspeciales = CadenaFinal
End Function

Private Sub cmdEnviarCorreos_Click()
    
    Dim clsProveeSMS As New clsConsulta
    Dim SMS As New clsEnvioSMS
    clsProveeSMS.Inicializar AdoConn, AdoConnMaster
    clsProveeSMS.Ejecutar " SELECT par_texto FROM parametro WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='SMS'", "L"
    SMS.Inicializar (clsProveeSMS.adorec_Def("par_texto"))
    Dim CantSMS As Long
    Dim i As Long
    Dim j As Long
    Dim ii As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim CopiarDesde As Long
    Dim MensajeSMS As String
    Dim CarteraRoja As Double
    Dim CarteraAmarilla As Double
    Dim CarteraBlanca As Double

    CopiarDesde = 2
    CantSMS = 0
    CarteraRoja = 0
    CarteraAmarilla = 0
    CarteraBlanca = 0
    Destino = Buscar_Carpeta(Me.hwnd, "Carpetas a Subir")
    For i = 1 To VSFG.Rows - 1
        VSFG.GetMergedRange i, CopiarDesde, r1, c1, r2, c2
        VSFGExportar.Clear flexClearScrollable, flexClearText
        VSFGExportar.Rows = 1
        k = 1
        For ii = r1 To r2
            VSFGExportar.AddItem ""
            For j = CopiarDesde + 3 To VSFG.Cols - 1
                VSFGExportar.TextMatrix(k, j - (CopiarDesde + 2)) = VSFG.TextMatrix(ii, j)
            Next j
            If VSFGExportar.TextMatrix(k, 6) <= HoyDia Then
                VSFGExportar.Cell(flexcpBackColor, k, 0, k, VSFGExportar.Cols - 1) = vbRed
                If VSFGExportar.TextMatrix(k, 4) <> "NOTA CREDITO" Then
                    CarteraRoja = CarteraRoja + VSFGExportar.TextMatrix(k, 8)
                End If
            ElseIf VSFGExportar.TextMatrix(k, 6) < DateAdd("d", 7, HoyDia) Then
                VSFGExportar.Cell(flexcpBackColor, k, 0, k, VSFGExportar.Cols - 1) = vbYellow
                If VSFGExportar.TextMatrix(k, 4) <> "NOTA CREDITO" Then
                    CarteraAmarilla = CarteraAmarilla + VSFGExportar.TextMatrix(k, 8)
                End If
            Else
                If VSFGExportar.TextMatrix(k, 4) <> "NOTA CREDITO" Then
                    CarteraBlanca = CarteraBlanca + VSFGExportar.TextMatrix(k, 8)
                End If
            End If
            k = k + 1
        Next ii
        
        VSFGExportar.AddItem ""
        
        VSFGExportar.AddItem ""
        VSFGExportar.TextMatrix(k + 1, VSFGExportar.Cols - 5) = "SubTotal Cartera Vencida"
        VSFGExportar.TextMatrix(k + 1, VSFGExportar.Cols - 1) = CarteraRoja
        VSFGExportar.Cell(flexcpBackColor, k + 1, 0, k + 1, VSFGExportar.Cols - 1) = vbRed
        VSFGExportar.AddItem ""
        VSFGExportar.TextMatrix(k + 2, VSFGExportar.Cols - 5) = "SubTotal Cartera X Vencer (7 días)"
        VSFGExportar.TextMatrix(k + 2, VSFGExportar.Cols - 1) = CarteraAmarilla
        VSFGExportar.Cell(flexcpBackColor, k + 2, 0, k + 2, VSFGExportar.Cols - 1) = vbYellow
        VSFGExportar.AddItem ""
        VSFGExportar.TextMatrix(k + 3, VSFGExportar.Cols - 5) = "SubTotal Cartera X Vencer (más de 7 días)"
        VSFGExportar.TextMatrix(k + 3, VSFGExportar.Cols - 1) = CarteraBlanca
        VSFGExportar.AddItem ""
        VSFGExportar.TextMatrix(k + 4, VSFGExportar.Cols - 5) = "TOTAL CARTERA"
        VSFGExportar.TextMatrix(k + 4, VSFGExportar.Cols - 1) = CarteraBlanca + CarteraAmarilla + CarteraRoja
        VSFGExportar.Cell(flexcpFontBold, k + 4, 0, k + 4, VSFGExportar.Cols - 1) = True
        VSFGExportar.CellBorderRange k + 1, VSFGExportar.Cols - 5, k + 4, VSFGExportar.Cols - 1, vbBlack, 2, 2, 2, 2, 0, 2
        VSFGExportar.ShowCell k + 4, VSFGExportar.Cols - 1
        
        MensajeSMS = "Estimad@ " & Left(Trim(VSFG.TextMatrix(i, 0)), 14) & ",JSN informa q su saldo vencido al " & UCase(Format(HoyDia, "d mmm")) & _
                    " es de $" & Format(CarteraRoja, "#,##0.00") & ".Agradecemos su pago." & _
                    "Cartera a vencer el " & UCase(Format(DateAdd("d", 7, HoyDia), "d mmm")) & " $" & Format(CarteraAmarilla, "#,##0.00") & ".Más detalle en su mail"
        
        If Trim(VSFG.TextMatrix(i, 2)) <> "" Then
            CantSMS = CantSMS + 1
            SMS.Enviar MensajeSMS, Trim(VSFG.TextMatrix(i, 2))
        End If
        
        VSFGExportar.SaveGrid Destino & "\" & Replace(Replace(Replace(VSFG.TextMatrix(i, 0), " ", "_"), "/", "_"), "\", "_") & ".xls", flexFileExcel, flexXLSaveFixedCells
        EnviarMail NombreComercial & " Cartera", CorreoCartera, VSFG.TextMatrix(i, 0), Trim(VSFG.TextMatrix(i, 1)), VSFG.TextMatrix(i, 3), "Cartera General al " & HoyDia, _
                    "Estimad@" & vbNewLine & _
                    VSFG.TextMatrix(i, 0) & vbNewLine & vbNewLine & _
                    "Adjunto encontrarás el Reporte de Cartera donde estan todas las cuentas pendientes que tienes con nosotros." & vbNewLine & vbNewLine & _
                    "Recuerda que:" & vbNewLine & _
                    "* Las facturas resaltadas con color ROJO son facturas que debieron ser canceladas hasta el dia de hoy," & vbNewLine & _
                    "* Las facturas resaltadas con color AMARILLO deben ser canceladas en el transcurso de la proxima semana, y " & vbNewLine & _
                    "* El resto de facturas deberán ser canceladas posteriormente según su vencimiento." & vbNewLine & _
                    "Si tienes alguna novedad por favor no dudes en comunicarte con tu Gestor de Cobranza." & vbNewLine & vbNewLine & _
                    "Saludos Cordiales" & vbNewLine & _
                    "Cartera" & vbNewLine & _
                    NombreComercial, Destino & "\" & Replace(Replace(Replace(VSFG.TextMatrix(i, 0), " ", "_"), "/", "_"), "\", "_") & ".xls"
        i = r2
        
        CarteraRoja = 0
        CarteraAmarilla = 0
        CarteraBlanca = 0
    Next i
    MsgBox "Envio terminado " & CantSMS & vbNewLine & Now
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
    Set ucrtVSFG1.VSFGControl = VSFG
    ucrtVSFG1.Inicializar False, False, False
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
