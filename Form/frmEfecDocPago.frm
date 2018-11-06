VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmEfecDocPago 
   BackColor       =   &H00BAA892&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efectivización de Documentos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmEfecDocPago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10440
   Begin VB.Frame Frame3 
      BackColor       =   &H00BAA892&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   1095
      Left            =   3000
      TabIndex        =   24
      Top             =   2400
      Width           =   4335
      Begin VB.OptionButton optProtestado 
         BackColor       =   &H00BAA892&
         Caption         =   "Protestado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optCobrado 
         BackColor       =   &H00BAA892&
         Caption         =   "Cobrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optGirado 
         BackColor       =   &H00BAA892&
         Caption         =   "Girado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
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
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   5880
      Width           =   1815
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
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00BAA892&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   1815
      Left            =   5400
      TabIndex        =   16
      Top             =   360
      Width           =   4695
      Begin VB.TextBox txtFechaDoc 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFechaPago 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbdia 
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
         ItemData        =   "frmEfecDocPago.frx":0442
         Left            =   3795
         List            =   "frmEfecDocPago.frx":04A3
         TabIndex        =   8
         Text            =   "DIA"
         Top             =   1200
         Width           =   780
      End
      Begin VB.ComboBox cmbmes 
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
         ItemData        =   "frmEfecDocPago.frx":051A
         Left            =   2970
         List            =   "frmEfecDocPago.frx":0545
         TabIndex        =   7
         Text            =   "MES"
         Top             =   1200
         Width           =   780
      End
      Begin VB.ComboBox Cmbaño 
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
         ItemData        =   "frmEfecDocPago.frx":0585
         Left            =   2160
         List            =   "frmEfecDocPago.frx":05E6
         TabIndex        =   6
         Text            =   "AÑO"
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblfechapago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   195
         Left            =   690
         TabIndex        =   23
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C3DBD1&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Efectivización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAA892&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644017&
      Height          =   1815
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   4695
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo dcmbDocumento 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNumero 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbBanco 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   195
         Left            =   1380
         TabIndex        =   27
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad Bancaria:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   210
         Left            =   510
         TabIndex        =   19
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label lbldescripcion1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   210
         Left            =   345
         TabIndex        =   15
         Top             =   1005
         Width           =   1545
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644017&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   270
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFG 
      Height          =   2055
      Left            =   1560
      TabIndex        =   10
      Top             =   3840
      Width           =   7320
      _cx             =   12912
      _cy             =   3625
      _ConvInfo       =   1
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
      FormatString    =   $"frmEfecDocPago.frx":06A4
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
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   1200
      Picture         =   "frmEfecDocPago.frx":075A
      Top             =   4800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   960
      Picture         =   "frmEfecDocPago.frx":0886
      ToolTipText     =   "Elimina una Fila"
      Top             =   4800
      Visible         =   0   'False
      Width           =   225
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
      ForeColor       =   &H00644017&
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   5880
      Width           =   855
   End
End
Attribute VB_Name = "frmEfecDocPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de Efectivización de Documentos                                       #
'#  frmEfecDocPago V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite cambiar el estado de los documentos de cobros de GIRADOS#
'#  a PROTESTADOS o COBRADOS                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  doc_pago: se almacena todos los datos consernientes al documento con que se #
'#            recepta el pago de una cuenta                                     #
'#  tip_asiento: Almacena los tipos de asientos contables                       #
'#  ctaconta: Almacena los asientos contables de los documentos efectivizados   #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsSql As New clsConsulta
Private clsTip As New clsConsulta
Private clsNum As New clsConsulta
Private clsFec As New clsConsulta
Private clsasi As New clsConsulta
Private clsBan As New clsConsulta
Private clsCta As New clsConsulta
Dim strSql As String

Private Sub ponerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
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
    TxtTotalHaber = Format(SumaHaber, "##0.00")
    
End Sub

Private Sub cmdAceptar_Click()
    
    If optGirado.Value = True Then
        MsgBox "El documento no ha cambiado de Estado", vbInformation, "Efectivización de Docuemntos"
        Exit Sub
    End If
    
    ff = Format(Cmbaño + "-" + cmbmes + "-" + cmbdia, "yyyy-mm-dd")
    'verifica que todos los datos esten ingresados
    If txtValor = "" Or VSFG.TextMatrix(1, 3) = "" Then
        MsgBox "No estan ingresados todos los datos", vbExclamation, "Efectivización de Documentos"
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
    
    
    'verifica que el debe y el haber esten cuadrados
    If txtTotalDebe <> TxtTotalHaber Then
        MsgBox "No esta cuadrado el Debe y el Haber", vbInformation, "Comprobante de Egreso"
        txtValor.SetFocus
        Exit Sub
    ElseIf optCobrado.Value = True Then
        'actualizamos tabla de documeto de pago
        strSql = " UPDATE doc_pago " & _
                 " SET doc_pag_fecha_efec = '" & ff & "', doc_pag_estado = 'COBRADO', doc_pag_fechamod = CURRENT_TIMESTAMP, doc_pag_usumod = substring_index(USER(),'@',1) " & _
                 " WHERE emp_codigo = '" & strEmpresa & "' AND tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' AND doc_pag_codigo = '" & dcmbNumero.BoundText & "' "
                clsSql.Ejecutar strSql
        'ingrasamos asientos
        'Busca el código máximo de la tabla ctaconta
        strSql = " Select max(SUBSTRING(asi_numasiento,1,11)) as numAS " & _
                " From asiento " & _
                " WHERE emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar (strSql)
        
        If Not IsNull(clsSql.adorec_Def("numas")) Then
            Maximo = clsSql.adorec_Def("numAS") + 1
            strmaximo = Space(11 - Len(str(Maximo))) & Trim(str(Maximo))
        Else
            Maximo = 1
            strmaximo = Space(11 - Len(str(Maximo))) & Trim(str(Maximo))
        End If
        
        'Ingreso de datos en la tabla asiento
        Set clsIngAsiento = New clsConsulta
        clsIngAsiento.Inicializar AdoConn
        descripcion = "Cheque Cobrado No.:" + " " + dcmbNumero.Text + " " + "valor:" + " " + txtValor
        strSql = " INSERT INTO asiento (asi_numasiento, emp_codigo,asi_fecha, asi_revisado, asi_mayorizado, asi_totaldebe, asi_totalhaber, asi_descripcion, asi_fechamod, asi_usumod) " & _
                 " VALUES ('" & strmaximo & "','" & strEmpresa & "','" & ff & "', '0','0', '" & Replace(txtTotalDebe, ",", ".") & "', '" & Replace(TxtTotalHaber, ",", ".") & "', '" & descripcion & "', CURRENT_TIMESTAMP, substring_index(USER(),'@',1))"
        clsIngAsiento.Ejecutar (strSql)
        
        
        'Ingreso de Detalle de asientos
            Set clsIngDetAsiento = New clsConsulta
            clsIngDetAsiento.Inicializar AdoConn
            With VSFG
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Or Val(.TextMatrix(i, 3)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                        strSql = " INSERT INTO det_asiento ( emp_codigo, asi_numasiento,cta_codigo, det_asi_debe, det_asi_haber, det_asi_fechamod, det_asi_usumod) " & _
                                 " VALUES ('" & strEmpresa & "','" & strmaximo & "','" & .TextMatrix(i, 1) & "', " & _
                                 " '" & Replace(.TextMatrix(i, 3), ",", ".") & "', '" & Replace(.TextMatrix(i, 4), ",", ".") & "', CURRENT_TIMESTAMP, substring_index(USER(),'@',1))"
                        clsIngDetAsiento.Ejecutar (strSql)
                    End If
                Next i
            End With
            
        MsgBox "Los datos han sido ingresados con éxito", vbInformation, "Efectivización de Documentos"
        
        dcmbDocumento.Text = ""
        txtValor = ""
        txtFechaPago = ""
        txtFechaDoc = ""
        VSFG.Clear 1
        VSFG.Rows = 2
        txtTotalDebe = 0
        TxtTotalHaber = 0
        Else
            Exit Sub
            
    End If
        
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dcmbasiento_Change()
    If dcmbasiento.Text = "" Then
        dcmbdescripciona.Text = ""
    Else
        cmdAceptar.Enabled = True
        clsasi.Actualizar
        clsasi.Filtrar "tip_asi_codigo = '" & dcmbasiento.BoundText & "'"
            dcmbdescripciona.Tag = "A"
            dcmbdescripciona = clsasi.adorec_Def("descripcionasi")
        clsasi.QuitarFiltro
        dcmbdescripciona.Tag = ""
    End If
End Sub


Private Sub dcmbBanco_Change()
dcmbNumero = ""
 'consulta el numero de los documentos existentes en el tipo de documeto seleccionado
    strSql = " SELECT doc_pag_codigo,doc_pag_numero, doc_pag_persona, CONCAT(doc_pag_numero , ' ' , doc_pag_persona) as numero, doc_pag_estado " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo = '" & strEmpresa & " ' AND tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' AND doc_pag_estado= 'GIRADO' AND ban_codigo='" & dcmbBanco.BoundText & "' "
    clsNum.Ejecutar strSql
    If Not clsNum.adorec_Def.EOF Then
      
        Set dcmbNumero.RowSource = clsNum.adorec_Def.DataSource
        dcmbNumero.ListField = "numero"
        dcmbNumero.BoundColumn = "doc_pag_codigo"
        dcmbNumero.Tag = clsNum.adorec_Def("doc_pag_persona")
        Frame1.Tag = clsNum.adorec_Def("doc_pag_numero")
        optGirado.Value = True
'        estado = clsNum.adorec_Def("doc_pag_estado")
    Else
        Set dcmbNumero.RowSource = Nothing
        estado = ""
    End If
End Sub

Private Sub dcmbDescripciona_Change()
  'Cambia el valor del codigo para actualizar este y la descripcion
  If dcmbasiento.Tag <> "A" Then
        If dcmbdescripciona.MatchedWithList = True Then
            dcmbasiento.BoundText = dcmbdescripciona.BoundText
        End If
    End If
End Sub


Private Sub dcmbDescripciona_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbasiento.BoundText = dcmbdescripciona.BoundText
    End If
End Sub

Private Sub dcmbdescripciona_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbasiento.BoundText = dcmbdescripciona.BoundText
End Sub


Private Sub dcmbDocumento_Change()
    dcmbBanco.Text = ""
    Set dcmbNumero.RowSource = Nothing
    dcmbNumero.Text = ""
    txtdocumento = ""
    txtValor = ""
    txtFechaDoc = ""
    txtFechaPago = ""
    
'    'consulta el numero de los documentos existentes en el tipo de documeto seleccionado
'    strSql = " SELECT doc_pag_numero, doc_pag_estado " & _
'             " FROM doc_pago " & _
'             " WHERE emp_codigo = '" & strEmpresa & " ' AND tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' AND doc_pag_estado= 'GIRADO' "
'    clsNum.Ejecutar strSql
'    If Not clsNum.adorec_Def.EOF Then
'        Set dcmbNumero.RowSource = clsNum.adorec_Def.DataSource
'        dcmbNumero.ListField = "doc_pag_numero"
'        optGirado.Value = True
''        estado = clsNum.adorec_Def("doc_pag_estado")
'    Else
'        Set dcmbNumero.RowSource = Nothing
'        estado = ""
'    End If
'
'    Select Case estado
'
'        Case "GIRADO"
'            optGirado.Value = True
'        Case "PROTESTADO"
'            optProtestado.Value = True
'        Case "COBRADO"
'            optCobrado.Value = True
'        Case ""
'
'    End Select
    
End Sub


Private Sub dcmbNumero_Change()

If dcmbNumero = "" Then
    txtValor = ""
    txtFechaDoc = ""
    txtFechaPago = ""
    VSFG.Clear 1
    VSFG.Rows = 2
    Exit Sub
Else
    espacio = InStr(Trim(dcmbNumero.Text), " ")
    numero = Mid(dcmbNumero.Text, 1, espacio)
    strSql = " SELECT doc_pag_valor, doc_pag_fecha_recepcion, doc_pag_fecha_doc " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo = '" & strEmpresa & " ' AND tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' AND ban_codigo = '" & dcmbBanco.BoundText & "' AND doc_pag_codigo = '" & dcmbNumero.BoundText & "' "
    clsNum.Ejecutar strSql
    
    If clsNum.adorec_Def.BOF = False And Not IsNull(clsNum.adorec_Def("doc_pag_valor")) Then
        txtValor = clsNum.adorec_Def("doc_pag_valor")
        txtFechaDoc = clsNum.adorec_Def("doc_pag_fecha_doc")
        txtFechaPago = clsNum.adorec_Def("doc_pag_fecha_recepcion")
    Else
        txtValor = ""
        txtFechaDoc = ""
        txtFechaPago = ""
    End If
End If
    
    
End Sub

Private Sub Form_Activate()
   dcmbDocumento.Text = ""
        txtValor = ""
        txtFechaPago = ""
        txtFechaDoc = ""
        VSFG.Clear 1
        VSFG.Rows = 2
        txtTotalDebe = 0
        TxtTotalHaber = 0
        optGirado.Value = True
End Sub

'Private Sub Form_Activate()
'    dcmbDocumento = ""
'    dcmbasiento = ""
'End Sub

Private Sub Form_Load()
    
    'inicializa variables de consulta
    clsSql.Inicializar AdoConn
    clsTip.Inicializar AdoConn
    clsNum.Inicializar AdoConn
    clsFec.Inicializar AdoConn
    clsasi.Inicializar AdoConn
    clsBan.Inicializar AdoConn
    clsCta.Inicializar AdoConn
    
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - mdiPrincipal.Height / 40
    
    'consulta los tipos de documentos
    strSql = " SELECT * " & _
             " from tipo_doc_pago "
    clsTip.Ejecutar strSql
    
    If Not clsTip.adorec_Def.EOF Then
        Set dcmbDocumento.RowSource = clsTip.adorec_Def.DataSource
        dcmbDocumento.ListField = "tip_doc_pag_nombre"
        dcmbDocumento.BoundColumn = "tip_doc_pag_codigo"
    Else
        Set dcmbDocumento.RowSource = Nothing
    End If
    
'    'Consulta el tipo de asiento para hacer el asiento del comprobante de egreso
'    strSql = " SELECT tip_asi_codigo, tip_asi_nombre, CONCAT(SUBSTRING(tip_asi_descripcion,1,20),'...') as descripcionasi " & _
'             " FROM tipo_asiento "
'    clsasi.Ejecutar strSql
'
'    If Not clsasi.adorec_Def.EOF Then
'        Set dcmbasiento.RowSource = clsasi.adorec_Def.DataSource
'        dcmbasiento.ListField = "tip_asi_nombre"
'        dcmbasiento.BoundColumn = "tip_asi_codigo"
'        Set dcmbdescripciona.RowSource = clsasi.adorec_Def.DataSource
'        dcmbdescripciona.ListField = "descripcionasi"
'        dcmbdescripciona.BoundColumn = "tip_asi_codigo"
'    Else
'        Set dcmbasiento.RowSource = Nothing
'        Set dcmbdescripciona.RowSource = Nothing
'    End If
    
    strSql = " SELECT ban_nombre, ban_codigo " & _
             " FROM banco "
    clsBan.Ejecutar strSql
    If Not clsBan.adorec_Def.EOF Then
        Set dcmbBanco.RowSource = clsBan.adorec_Def.DataSource
        dcmbBanco.ListField = "ban_nombre"
        dcmbBanco.BoundColumn = "ban_codigo"
    Else
        Set dcmbBanco.RowSource = Nothing
    End If
    
    d = CStr(Day(Date))
    mm = Month(Date)
    Y = CStr(Year(Date))
    cmbdia.Text = d
    Cmbaño.Text = Y
    
    For var = 1 To 12
        If cmbmes.ItemData(var) = mm Then
            cmbmes.Text = cmbmes.List(var)
            Exit For
        End If
    Next var
    
             
End Sub

Private Sub optCobrado_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub optGirado_Click()
'    VSFG.Clear 1
'    VSFG.Rows = 2
End Sub

Private Sub optProtestado_Click()
    
'    VSFG.Clear 1
'    VSFG.Rows = 2
    
    Dim caracter1 As Long
    Dim caracter2 As Long
    
    ff = Format(Cmbaño + "-" + cmbmes + "-" + cmbdia, "yyyy-mm-yy")
    If MsgBox(" ¿Está seguro que el Documento seleccionado está protestado? ", vbYesNo + vbExclamation, "Efectivizacón de Documentos") = vbYes Then
        
        strSql = " UPDATE doc_pago " & _
                 " SET doc_pag_fecha_efec = '" & ff & "', doc_pag_estado = 'PROTESTADO', doc_pag_fechamod = CURRENT_TIMESTAMP, doc_pag_usumod = substring_index(USER(),'@',1)" & _
                 " WHERE tip_doc_pag_codigo = '" & dcmbDocumento.BoundText & "' AND doc_pag_codigo = '" & dcmbNumero.BoundText & "' "
        clsSql.Ejecutar strSql
        
        If MsgBox(" Debe generar una Cuenta por Cobrar " & vbCrLf & _
                  " por el valor de" + " " + txtValor + " " + "dolares mas el costo del protesto.", vbOKOnly + vbInformation, "Documentos de Protesto") = vbOK Then
            descripcion = UCase("Cheque Protestado No.:" + " " + dcmbNumero.Text + " " + "valor:" + " " + txtValor)
            
            caracter1 = InStr(Trim(dcmbNumero.Text), " ")
            caracter2 = InStr(caracter1 + 1, Trim(dcmbNumero.Text), " ")
            nombre = Mid(Trim(dcmbNumero.Text), caracter1, caracter2 - caracter1)
            apellido = Mid(Trim(dcmbNumero.Text), caracter2)
            strSql = " SELECT cat_p_tipo " & _
                    " FROM persona " & _
                    " WHERE per_nombre= '" & Trim(nombre) & "' and per_apellido = '" & Trim(apellido) & "' "
            clsFec.Ejecutar strSql
            If Not clsFec.adorec_Def.EOF Then
                categoria = clsFec.adorec_Def("cat_p_tipo")
                If categoria = "C" Then
                    
                    frmCtaxc_p.OptCliente = True
                ElseIf categoria = "P" Then
                 
                    frmCtaxc_p.Optproveedores.Value = True
                    
                End If
                nombre = dcmbNumero.Tag
                frmCtaxc_p.DCmbNomPersona = nombre
            Else
                  nombre = ""
                  frmCtaxc_p.DCmbNomPersona = ""
            End If
            frmCtaxc_p.TxtObservacion.Enabled = False
            frmCtaxc_p.TxtObservacion = descripcion
            frmCtaxc_p.Tag = "C"
            frmCtaxc_p.txtdocumento = dcmbNumero.Text
            frmCtaxc_p.Show
        End If
    Else
        optGirado.Value = True
    End If
End Sub
Private Sub txtTotalDebe_Change()
    txtTotalDebe = Format(Val(txtTotalDebe), "##0.00")
End Sub

Private Sub txtTotalHaber_Change()
 TxtTotalHaber = Format(Val(TxtTotalHaber), "##0.00")
End Sub

Private Sub txtValor_Change()
    VSFG.TextMatrix(1, 3) = txtValor
    txtValor = Format(Val(txtValor), "##0.00")
End Sub


'Detecta cuando se ha dado un enter para enviar un tab
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
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
            ponerBotones
        End If
    End If
End Sub


Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = 1 Then
        If Col = 3 Then
            Cancel = True
        End If
        If Col = 4 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    'Verifica que solo se ingresen números tanto en el Debe como en el Haber
    If Col = 3 And VSFG.TextMatrix(Row, 1) = "" And VSFG.TextMatrix(Row, 2) = "" Then
            MsgBox "Ingrese la cuenta contable", vbInformation, "Detalle"
            VSFG.TextMatrix(Row, 3) = 0
            VSFG.TextMatrix(Row, 4) = 0
           
        ElseIf Col = 3 Or Col = 4 Then
        'Verifica que solo se ingresen números en el campo Debe

        If Not IsNumeric(VSFG.TextMatrix(Row, 3)) And VSFG.TextMatrix(Row, 3) <> "" Then
            MsgBox "Ingrese solo números en el Debe.", vbInformation, "Debe"
            VSFG.TextMatrix(Row, 3) = 0
        End If

        If Not IsNumeric(VSFG.TextMatrix(Row, 4)) And VSFG.TextMatrix(Row, 4) <> "" Then
            MsgBox "Ingrese solo números en el Haber.", vbInformation, "Haber"
            VSFG.TextMatrix(Row, 4) = 0
        End If
        CalcuTotal
    End If
End Sub

Private Sub VSFG_CellChanged(ByVal Row As Long, ByVal Col As Long)

    
' filtra el nombre y codigo de cuenta para los combos del greed
'If Row > 0 Then
    'hace la consulta para saber las cuentas contables que no tengan subcuentas
     strSql = " SELECT cta_codigo, cta_nombre" & _
                 " FROM ctaconta " & _
                 " WHERE cta_subcta = '0' AND emp_codigo = '" & strEmpresa & "'" & _
                 " ORDER BY cta_codigo"
     clsCta.Ejecutar strSql
    
     VSFG.ColComboList(1) = VSFG.BuildComboList(clsCta.adorec_Def, "*cta_codigo, cta_nombre", "cta_codigo")
     VSFG.ColComboList(2) = VSFG.BuildComboList(clsCta.adorec_Def, "cta_codigo, *cta_nombre")

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
CalcuTotal
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

Private Sub VSFG_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    ' si el boton es derecho se procede
    If Button <> 1 Then Exit Sub

    ' se obtiene la celda seleccionada
    Dim r&, c&
    r = VSFG.MouseRow
    c = VSFG.MouseCol

    ' si la fila y la columna son menores que 0 no hace nada
    If r < 0 Or c < 0 Then Exit Sub
    
    'si la coluna es diferente de 0 no hace nada
    If (c <> 0) Then Exit Sub

     'If (c <> 0 Or r = (VSFG.Rows - 1)) Then Exit Sub
    
    'Si la fila es mayor que 0 procede a verificación para borrar
    If r > 0 Then
        'Si la columna es mayor que 0 verifica que no tenga el icono
        If c > 1 Then
            If VSFG.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFG.Cell(flexcpLeft, r, c) + VSFG.Cell(flexcpWidth, r, c) - X
        If d > imgBtnDn.Width Then Exit Sub
        'si la fila es mayor que la primera procede a borrar
        If r > 1 Then
        ' click was on a button: do the work
        VSFG.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Egresos Efectivización de Documentos"   ' Define el título.
        Respuesta = MsgBox(Mensaje, Estilo, Título)
    
        'Recorro el FlexGrid para poner números a las filas
    
        If Respuesta = vbYes Then
            Dim i As Integer
            VSFG.RemoveItem (r)
            ponerBotones
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
Private Sub VSFG_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (VSFG.TextMatrix(VSFG.Row, 3) = "") Then
                VSFG.TextMatrix(VSFG.Row, 3) = 0
     ElseIf VSFG.TextMatrix(VSFG.Row, 4) = "" Then
                VSFG.TextMatrix(VSFG.Row, 4) = 0
     End If
End Sub
