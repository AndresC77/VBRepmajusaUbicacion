VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEgresoComun 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Detalle de Egresos Comunes"
   ClientHeight    =   5535
   ClientLeft      =   4635
   ClientTop       =   3135
   ClientWidth     =   7965
   Icon            =   "frmEgresoComun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7965
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Egresos Comu nes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   135
      TabIndex        =   12
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   7335
      End
      Begin VB.TextBox TxtTotalHaber 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   4320
         Width           =   1815
      End
      Begin VB.ComboBox cmbDia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEgresoComun.frx":030A
         Left            =   6765
         List            =   "frmEgresoComun.frx":036B
         TabIndex        =   4
         Text            =   "DIA"
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cmbMes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEgresoComun.frx":03E2
         Left            =   5970
         List            =   "frmEgresoComun.frx":040D
         TabIndex        =   3
         Text            =   "MES"
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cmbAño 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEgresoComun.frx":044D
         Left            =   5160
         List            =   "frmEgresoComun.frx":04AE
         TabIndex        =   2
         Text            =   "AÑO"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txtTotalDebe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3795
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   255
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcmbCuenta 
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbBanco 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   7320
         _cx             =   12912
         _cy             =   3625
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEgresoComun.frx":056C
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
      Begin VB.Image imgBtnDn 
         Height          =   210
         Left            =   3000
         Picture         =   "frmEgresoComun.frx":0622
         Top             =   1080
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgBtnUp 
         Height          =   210
         Left            =   3240
         Picture         =   "frmEgresoComun.frx":074E
         ToolTipText     =   "Elimina una Fila"
         Top             =   1080
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lblCodigo 
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
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   540
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
         Left            =   120
         TabIndex        =   17
         Top             =   645
         Width           =   510
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
         TabIndex        =   16
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblcuentaban 
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
         Left            =   3840
         TabIndex        =   15
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label lblFecha 
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
         Height          =   165
         Left            =   3840
         TabIndex        =   14
         Top             =   330
         Width           =   525
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
         Left            =   2955
         TabIndex        =   13
         Top             =   4380
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4042
      TabIndex        =   11
      Top             =   5040
      Width           =   1920
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2002
      TabIndex        =   10
      Top             =   5040
      Width           =   1920
   End
End
Attribute VB_Name = "frmEgresoComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma ingreso y modificación de egresos comunes                             #
'#  frmEgresoComun V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de egresos comunes de la empresa.   #
'#  Permitirá almacenar en la base de datos egresos comunes en las tablas       #
'#  egreso_comun y det_egreso_comun tanto en la modificación como en el ingreso #
'#  de nuevos datos, para modificar o crear un egreso nuevo los valores vienen  #
'#  de frmSelEgresoComun                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  EGRESO_COMUN: En esta tabla se almacenan las cabeceras de los egresos       #
'#  det_egreso_comun: donde se guardan los datos de el detalle del egreso       #
'#  BANCO: De donde se consultan los bancos existentes                          #
'#  Cta_banco: de donde se consultan los números de cuentas bancarias y         #
'#             su cuenta contable                                               #
'#  CTACONTA: Para consultar las cuentas cosntables y sus nombres               #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsEgr As New clsConsulta
Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Private clsDet As New clsConsulta
Private clsSql As New clsConsulta
Private clsCtb As New clsConsulta
Private clsctc As New clsConsulta
Private strSQL As String
Private intDato As Variant
Dim ff As Variant
Dim m As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsEgr = Nothing
    Set clsBan = Nothing
    Set clsCta = Nothing
    Set clsDet = Nothing
    Set clsSql = Nothing
    Set clsCtb = Nothing
    Set clsctc = Nothing
End Sub

Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la seginda columna del grid de todas las filas
    For i = 1 To (VSFG.Rows - 1)
        VSFG.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFG.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
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
    
End Sub

Private Sub cmdAceptar_Click()
   
    Dim i As Long
    Dim j As Long
    Dim a As Long
    
    ' pone en la variable ff la fecha para la bdd
    
    ff = Format(cmbAño + "-" + cmbMes + "-" + cmbDia, "yyyy-mm-dd")

    If (IsDate(ff) = False) Then
        MsgBox "La fecha no es válida", vbInformation, "Egresos Comunes"
        Exit Sub
    End If

    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> txtTotalHaber Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbExclamation, "Comprobante de Egreso Común"
        Exit Sub
    End If

    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd

    a = VSFG.Rows - 1
    For i = 1 To a
        For j = i + 1 To a
            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) Then
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

    If Me.Tag = "N" Then
    
        strSQL = " SELECT Count(egr_com_codigo) as Codigo  " & _
                  " FROM egreso_comun " & _
                  " WHERE egr_com_codigo = '" & txtCodigo & "'"
        clsSql.Ejecutar strSQL
    
        codigo = clsSql.adorec_Def("Codigo")
    
        If codigo <> 0 Then
           MsgBox "El código ya existe", vbInformation, "Egreso Comun"
           txtCodigo.SetFocus
           Exit Sub
        Else
    
'    'Suma los valores de las columnas 3 y 4 de las cuentas que se repitan en el greed para grabar en la bdd
'
'    a = VSFG.Rows - 1
'    For i = 1 To a
'        For j = i + 1 To a
'            If VSFG.TextMatrix(i, 1) = VSFG.TextMatrix(j, 1) Then
'                VSFG.TextMatrix(i, 3) = Val(VSFG.TextMatrix(i, 3)) + Val(VSFG.TextMatrix(j, 3))
'                VSFG.TextMatrix(i, 4) = Val(VSFG.TextMatrix(i, 4)) + Val(VSFG.TextMatrix(j, 4))
'                VSFG.RemoveItem j
'                a = a - 1
'                j = j - 1
'            End If
'            If j >= a Then
'                Exit For
'            End If
'        Next j
'    Next i

        
        
        'Verificar que todos los datos se han llenado para ingresar en la base de datos
        If txtCodigo = "" Or dcmbBanco = "" Or txtDescripcion = "" Then
            MsgBox "No estan ingresados todos los datos", vbInformation, "Ingreso"
            txtCodigo.SetFocus
            Exit Sub
        Else
            'Ingreso de datos en egreso_comun
            strSQL = " INSERT INTO egreso_comun (egr_com_codigo, emp_codigo, cta_ban_numero, ban_codigo, " & _
                 " egr_com_fecha, egr_com_descripcion, egr_com_fechamod, egr_com_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo) & "','" & strEmpresa & "','" & dcmbCuenta.BoundText & "','" & dcmbBanco.Tag & "', " & _
                 " '" & ff & "','" & UCase(txtDescripcion) & "', " & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsSql.Ejecutar (strSQL), "M"
            
            'ingreso de datos en el la tabla det_egreso_comun
        
            With VSFG
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                        strSQL = " INSERT INTO det_egreso_comun (egr_com_codigo, emp_codigo, cta_codigo, det_egr_com_debe, det_egr_com_haber, det_egr_com_fechamod, det_egr_com_usumod) " & _
                                 " VALUES ('" & UCase(txtCodigo) & "','" & strEmpresa & "', '" & .TextMatrix(i, 1) & "', " & _
                                 " '" & Replace(.TextMatrix(i, 3), ",", ".") & "', '" & Replace(.TextMatrix(i, 4), ",", ".") & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
                        clsSql.Ejecutar strSQL, "M"
                    End If
                Next i
            End With
            
            MsgBox " Los datos han sido ingresados", vbInformation, "Ingresos"
        End If

    End If
End If
    If Me.Tag = "M" Then
    
        'actualiza los valores en la cabecera
        strSQL = " UPDATE  egreso_comun" & _
                 " SET cta_ban_numero = '" & UCase(dcmbCuenta.Text) & " ',ban_codigo = '" & UCase(dcmbBanco.Tag) & "', " & _
                 " egr_com_fecha ='" & ff & "', egr_com_descripcion = '" & UCase(txtDescripcion) & "', egr_com_fechamod=CURRENT_TIMESTAMP,egr_com_usumod='" & strUsuario & "' " & _
                 " WHERE egr_com_codigo='" & txtCodigo.Text & "' AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSQL, "M"
        
        'Borra los valores del greed para ingresar los datos modificados
        strSQL = " DELETE FROM det_egreso_comun " & _
                 " WHERE egr_com_codigo = '" & txtCodigo.Text & "' AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSQL, "M"
    
        With VSFG
            For i = 1 To VSFG.Rows - 1
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
             
                    strSQL = " INSERT INTO det_egreso_comun (egr_com_codigo, emp_codigo, cta_codigo, det_egr_com_debe, det_egr_com_haber, det_egr_com_fechamod, det_egr_com_usumod) " & _
                             " VALUES ('" & txtCodigo & "','" & strEmpresa & "', '" & .TextMatrix(i, 1) & "', " & _
                             " '" & Replace(.TextMatrix(i, 3), ",", ".") & "', '" & Replace(.TextMatrix(i, 4), ",", ".") & "', CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    clsSql.Ejecutar strSQL, "M"
                End If
            Next i
        End With
        MsgBox " Los datos han sido modificados", vbInformation, "Modificar"
    End If
Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub
       
Private Sub dcmbBanco_Change()
    dcmbCuenta = ""
    dcmbBanco.Tag = dcmbBanco.BoundText
    strSQL = " SELECT cta_ban_numero,cta_ban_ctaconta " & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSQL

    Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
    dcmbCuenta.ListField = ("cta_ban_numero")
    
End Sub

Private Sub dcmbCuenta_Change()
    If dcmbCuenta <> "" Then
        If VSFG.Row = 1 Then
            strSQL = " SELECT cta_banco.cta_ban_ctaconta,ctaconta.cta_nombre " & _
                     " FROM cta_banco INNER JOIN ctaconta ON cta_banco.cta_ban_ctaconta=ctaconta.cta_codigo " & _
                     "                                    AND cta_banco.emp_codigo=ctaconta.emp_codigo " & _
                     " WHERE cta_banco.emp_codigo = '" & strEmpresa & "' AND cta_ban_numero = '" & dcmbCuenta & "' AND ban_codigo='" & dcmbBanco.BoundText & "'"
            clsctc.Ejecutar strSQL
            If clsctc.adorec_Def.RecordCount > 0 Then
                VSFG.TextMatrix(1, 1) = clsctc.adorec_Def("cta_ban_ctaconta")
                VSFG.TextMatrix(1, 2) = clsctc.adorec_Def("cta_nombre")
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    
    'consulta para saber los  bancos existentes
    strSQL = " SELECT ban_codigo, ban_nombre " & _
             " FROM banco " & _
             " ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL

    Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
    dcmbBanco.ListField = "ban_nombre"
    dcmbBanco.BoundColumn = "ban_codigo"
    
    'consulta para seleccionar las cuentas bancarias y las cuentas contables
    strSQL = " SELECT cta_ban_numero, cta_ban_ctaconta" & _
             " FROM cta_banco " & _
             " WHERE ban_codigo = '" & dcmbBanco.BoundText & "' " & _
             " AND emp_codigo = '" & strEmpresa & "' " & _
             " ORDER BY cta_ban_numero "
    clsCtb.Ejecutar strSQL

    Set dcmbCuenta.RowSource = clsCtb.adorec_Def.DataSource
    dcmbCuenta.ListField = ("cta_ban_numero")
    
    'Pone los combolist en las columnas 1 y 2 despues de la primera fila
    strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSQL
              
     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre")
 
If Me.Tag = "M" Then
' Pone el nombre a la forma
    Me.Caption = "Modificar Egresos Comunes"
 
    'selecciona los valores para el greed
    With VSFG
        strSQL = " SELECT distinct det_egreso_comun.cta_codigo,ctaconta.cta_nombre ,det_egr_com_debe, det_egr_com_haber " & _
                 " FROM ((( egreso_comun INNER JOIN det_egreso_comun ON egreso_comun.egr_com_codigo=det_egreso_comun.egr_com_codigo " & _
                 "                                                   AND egreso_comun.emp_codigo=det_egreso_comun.emp_codigo) " & _
                 "                       INNER JOIN ctaconta ON det_egreso_comun.cta_codigo= ctaconta.cta_codigo " & _
                 "                                           AND det_egreso_comun.emp_codigo= ctaconta.emp_codigo) " & _
                 "                       INNER JOIN cta_banco ON egreso_comun.cta_ban_numero=cta_banco.cta_ban_numero " & _
                 "                                            AND egreso_comun.emp_codigo=cta_banco.emp_codigo " & _
                 "                                            AND egreso_comun.ban_codigo=cta_banco.ban_codigo) " & _
                 " WHERE egreso_comun.egr_com_codigo = '" & txtCodigo & "' AND egreso_comun.emp_codigo = '" & strEmpresa & "'" & _
                " ORDER BY iff(det_egreso_comun.cta_codigo=cta_banco.cta_ban_ctaconta,0,1)"
                 
    clsDet.Ejecutar strSQL
        
    Set VSFG.DataSource = clsDet.adorec_Def.DataSource
    PonerBotones
    'calcula el total de debe y haber
    CalcuTotal
    End With
End If
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
    'Inicializa las clases para hacer distintas consultas
    clsEgr.Inicializar AdoConn, AdoConnMaster
    clsBan.Inicializar AdoConn, AdoConnMaster
    clsCta.Inicializar AdoConn, AdoConnMaster
    clsDet.Inicializar AdoConn, AdoConnMaster
    clsCtb.Inicializar AdoConn, AdoConnMaster
    clsSql.Inicializar AdoConn, AdoConnMaster
    clsctc.Inicializar AdoConn, AdoConnMaster
    'Realiza la consulta que contiene las diferentes cotizaciones en la empresa
    
    d = CStr(Day(HoyDia))
    mm = Month(HoyDia)
    m = Month(HoyDia)
    Y = CStr(Year(HoyDia))
    cmbDia.Text = d
    cmbAño.Text = Y

    For var = 0 To 11
        If cmbMes.ItemData(var) = mm Then
            cmbMes.Text = cmbMes.List(var)
            Exit For
        End If
    Next var
    strSQL = " SELECT ban_codigo, ban_nombre " & _
             " FROM banco " & _
             " ORDER BY ban_codigo"
    clsBan.Ejecutar strSQL

    Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
    dcmbBanco.ListField = "ban_nombre"
    dcmbBanco.BoundColumn = "ban_codigo"
    
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
            'MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
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

    If (c <> 0 Or r = 1) Then Exit Sub

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


Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)
' filtra el nombre y codigo de cuenta para los combos del grid
If Row > 1 Then
  
    
    With VSFG
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                 clsCta.Filtrar ("cta_codigo = '" & .TextMatrix(Row, 1) & "'")
                     .TextMatrix(Row, 2) = clsCta.adorec_Def("cta_nombre")
                 clsCta.QuitarFiltro
             End If
             
             If Col = 2 Then
                 clsCta.Filtrar ("cta_nombre = '" & .TextMatrix(Row, 2) & "'")
                     .TextMatrix(Row, 1) = clsCta.adorec_Def("cta_codigo")
                 clsCta.QuitarFiltro
             End If
         End If
    End With
End If
End Sub


Private Sub VSFG_KeyDown(KeyCode As Integer, Shift As Integer)
'hace que cuando llegue al final del greed, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes
    
    If VSFG.Row = VSFG.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFG.TextMatrix(VSFG.Row, 1) <> "" And (VSFG.TextMatrix(VSFG.Row, 3) <> "" Or VSFG.TextMatrix(VSFG.Row, 4) <> "") Then
            VSFG.AddItem ""
            VSFG.TextMatrix(VSFG.Rows - 1, 0) = VSFG.Rows - 1
            VSFG.Cell(flexcpPicture, (VSFG.Rows - 1), 0) = imgBtnUp
            VSFG.Cell(flexcpPictureAlignment, (VSFG.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
            strSQL = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSQL
              
     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre")
        End If
    End If
End Sub

Private Sub VSFG_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 1 Then
        If KeyCode = vbKeyF2 Then
            frmSelecCtaConta.Tag = "UN"
            frmSelecCtaConta.Show
            Set frmSelecCtaConta.objEscribir = VSFG
        End If
    End If
End Sub
