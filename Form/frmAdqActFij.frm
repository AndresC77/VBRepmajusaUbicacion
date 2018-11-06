VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmAdqActFij 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adquisición Activo Fijo"
   ClientHeight    =   7320
   ClientLeft      =   3705
   ClientTop       =   2040
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdqActFij.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8010
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4017
      TabIndex        =   17
      Top             =   6800
      Width           =   1700
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2294
      TabIndex        =   16
      Top             =   6800
      Width           =   1700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Proveedor"
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
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   7575
      Begin VB.TextBox txtDirProveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1125
         Width           =   2685
      End
      Begin VB.TextBox txtTelProveedor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5190
         TabIndex        =   9
         Top             =   750
         Width           =   2130
      End
      Begin VB.TextBox txtFaxProveedor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5190
         TabIndex        =   10
         Top             =   1125
         Width           =   2130
      End
      Begin VB.TextBox txtRucProveedor 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5190
         TabIndex        =   8
         Top             =   360
         Width           =   2130
      End
      Begin VB.TextBox txtNomProveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   750
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcmbCodProveedor 
         Height          =   330
         Left            =   1080
         TabIndex        =   5
         Top             =   345
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
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
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "RUC:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1155
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00BAA892&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Height          =   6735
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Detalle Adquisición"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   7575
         Begin VB.TextBox txtObs 
            Height          =   810
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   3120
            Width           =   7320
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   14
            Top             =   2640
            Width           =   1095
         End
         Begin VSFlex7Ctl.VSFlexGrid VSFDetalle 
            Height          =   1455
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   7335
            _cx             =   12938
            _cy             =   2566
            _ConvInfo       =   1
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdqActFij.frx":030A
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
         End
         Begin VB.TextBox TxtIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   13
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   12
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Image imgBtnUp 
            Height          =   210
            Left            =   120
            Picture         =   "frmAdqActFij.frx":03DC
            Top             =   1920
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Image imgBtnDn 
            Height          =   210
            Left            =   480
            Picture         =   "frmAdqActFij.frx":0512
            Top             =   1920
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label Label13 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Left            =   5160
            TabIndex        =   33
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lblIVA 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "IVA :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   5160
            TabIndex        =   32
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label12 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5160
            TabIndex        =   31
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label10 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2880
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Adquisición Activo Fijo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   7575
         Begin VB.TextBox txtNumAdquision 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   0
            Top             =   368
            Width           =   945
         End
         Begin VB.ComboBox cmbAño 
            Height          =   330
            ItemData        =   "frmAdqActFij.frx":063E
            Left            =   5040
            List            =   "frmAdqActFij.frx":069F
            TabIndex        =   2
            Text            =   "cmbAño"
            Top             =   360
            Width           =   810
         End
         Begin VB.ComboBox cmbMes 
            Height          =   330
            ItemData        =   "frmAdqActFij.frx":075D
            Left            =   5880
            List            =   "frmAdqActFij.frx":0788
            TabIndex        =   3
            Text            =   "cmbMes"
            Top             =   360
            Width           =   750
         End
         Begin VB.ComboBox cmbDia 
            Height          =   330
            ItemData        =   "frmAdqActFij.frx":07C8
            Left            =   6720
            List            =   "frmAdqActFij.frx":0829
            TabIndex        =   4
            Text            =   "cmbDia"
            Top             =   360
            Width           =   630
         End
         Begin VB.TextBox txtNumdoc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3000
            TabIndex        =   1
            Top             =   368
            Width           =   1125
         End
         Begin VB.Label Label9 
            BackColor       =   &H00DDDDDD&
            Caption         =   "# Orden:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2280
            TabIndex        =   34
            Top             =   398
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   4440
            TabIndex        =   28
            Top             =   398
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00BAA892&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   398
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmAdqActFij"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso de adquisiciones de Activos Fijos concepto de compra  #
'#  frmAdqActFij  V1.0                                                          #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana que permite ingresar los adquisiciones de activos fijos             #
'#  de la compañía por concepto de compra.                                      #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    adquisicion: En esta tabla se almacenan las nuevas adquisiciones          #
'#    det_adquisicion_af: En esta tabla se almacena los detalles de la          #
'#                       adquisiciòn                                            #
'#    persona    : Se consulta los proveedores de la empresa                    #
'#    activo Fijo: Se consulta los activos fijos de la empresa                  #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#               limpiarFxGD()   Permite borrar los datos que se encuentran     #
'#                               en el flexGrid para realizar un nuevo ingreso  #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsConsu clsConsulta: Objeto para consultar a la base de datos            #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsConsu As New clsConsulta
Private clsCon_Def As New clsConsulta
Private clsCon_Aux As New clsConsulta
Private clsCon_Prove As New clsConsulta
Private clsCon_Act As New clsConsulta
Private clsCon_Iva As New clsConsulta
Private strSql As String
Private tipo_ingreso As String
Private tipo_asiento As String
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsConsu = Nothing
    Set clsCon_Def = Nothing
    Set clsCon_Aux = Nothing
    Set clsCon_Prove = Nothing
    Set clsCon_Act = Nothing
    Set clsCon_Iva = Nothing
End Sub
Private Sub PonerBotones(Optional conBot As Boolean = True)
    'Agrega un botón de eliminar en la primera columna del grid de todas las filas
    For i = 1 To (VSFDetalle.Rows - 1)
        VSFDetalle.TextMatrix(i, 0) = i
        If conBot = True Then
            'Coloca los botones de elimniar fila en el grid
            VSFDetalle.Cell(flexcpPicture, i, 0) = imgBtnUp
            VSFDetalle.Cell(flexcpPictureAlignment, i, 0) = flexAlignRightCenter
        End If
    Next i
End Sub

Private Sub CalcuTotal()
   'Calcula totales
    Dim SubTotal As Double
    'Calcula Subtotal
    SubTotal = 0
    For i = 1 To VSFDetalle.Rows - 1
        SubTotal = SubTotal + Val(VSFDetalle.TextMatrix(i, 4))
    Next i
    txtSubTotal = FormatoD2(SubTotal)
    txtIVA = FormatoD2(Val(txtSubTotal) * (txtIVA.Tag / 100))
    txtTotal = FormatoD2(Val(txtSubTotal.Text) + Val(txtIVA.Text))
End Sub

Private Sub cmdAceptar_Click()
Dim X As Date
Dim i, j  As Integer
Dim cod_asiento As String
Dim f As Date
Dim d As String
Dim m As String
Dim Y As String
Dim ff As Variant
Dim ff1 As Variant


    'Valido los datos de # de Adquisicion, Proveedor, fecha de Adquisiion, etc.
    
    If (txtNumAdquision.Text = "") Then
        MsgBox "Número de Adquisición incorrecto", vbExclamation, "SisAdmi"
        txtNumAdquision.SetFocus
        Exit Sub
    End If
    If (txtNumdoc.Text = "") Then
        MsgBox "Número de Documento no Ingresado", vbExclamation, "SisAdmi"
        txtNumdoc.SetFocus
        Exit Sub
    End If
    If (txtIVA.Text = "") Then
        MsgBox "El Valor del Iva es incorrecto", vbExclamation, "SisAdmi"
        Exit Sub
    End If
    If (dcmbCodProveedor.Text = "" Or txtNomProveedor.Text = "" Or txtRucProveedor.Text = "") Then
        MsgBox "Datos del Proveedor incorrectos, verifíquelos", vbExclamation, "SisAdmi "
        dcmbCodProveedor.SetFocus
        Exit Sub
    End If

    ff = cmbDia.Text + "-" + cmbMes.Text + "-" + cmbAño.Text
    'Verifica si la fecha ingresada si es correcta
    If (IsDate(ff)) = False Then
        MsgBox "La fecha de Ingreso no es correcta", vbExclamation, "SisAdmi "
        Exit Sub
    End If
    
    'valido que no haga filas vacias
    band = 0
    For i = 1 To VSFDetalle.Rows - 1
        For j = 1 To VSFDetalle.Cols - 1
            If VSFDetalle.TextMatrix(i, j) = "" Then band = band + 1
        Next j
        If band = VSFDetalle.Cols - 1 Then VSFDetalle.RemoveItem (i)
    Next i

    'para buscar en adquisicion si se quiere ingresar un detalle
    'Verifica que existan datos en el FlexGrid
        For i = 1 To VSFDetalle.Rows - 1
         If (VSFDetalle.TextMatrix(i, 1) = "" And VSFDetalle.TextMatrix(i, 2) = "") Then
                If i = 1 Then
                    MsgBox "La Adquisicion no tiene detalle", vbExclamation, "SisAdmi "
                End If
                i = i - 1
                Exit For
         Else
            For j = 1 To VSFDetalle.Cols - 1
                If (VSFDetalle.TextMatrix(i, j) = "") Then
                    MsgBox "Dato incorrecto en: " & VSFDetalle.TextMatrix(i, j) & " ,fila: " & i, vbExclamation, "SisAdmi "
                    Exit Sub
                End If
            Next j
         End If
    Next i
   
    If (txtSubTotal.Text = "") Then
        Exit Sub
    End If
    
    If (i - 1 <> 0) Then ' Si existen detalles, almaceno.

        Mensaje = "Existen " & i - 1 & " detalle(s) en la Adquisición, desea guardar?" ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi "   ' Define el título.
        Respuesta = MsgBox(Mensaje, Estilo, Título)

    'Recorro el FlexGrid para almacenar los detalles de adquisicion
    '    Else
    If Respuesta = vbYes Then
        Dim aux As Integer
        clsCon_Def.Inicializar AdoConn
        strSql = " SELECT COALESCE(MAX(adq_codigo),0) " & _
                 " FROM adquisicion " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'"
        clsConsu.Ejecutar (strSql)
            If clsConsu.adorec_Def.EOF Then
                aux = 1
            Else
                aux = clsConsu.adorec_Def.Fields(0).Value + 1
            End If
            If (aux <> CInt(txtNumAdquision.Text)) Then
            MsgBox "El número de adquisicion ha cambiado a: " & aux, vbExclamation, " SisAdmi "
            txtNumAdquision.Text = aux
            End If
        strSql = " INSERT INTO adquisicion(emp_codigo,adq_subtotal,adq_impuesto,adq_codigo, " & _
                 "                         adq_numdoc,per_codigo, adq_fecha, adq_observacion," & _
                 "                         adq_fechamod, adq_usumod) " & _
                 "                   VALUES('" & strEmpresa & "'," & txtSubTotal.Text & "," & txtIVA.Text & ", " & _
                 "                           " & CInt(txtNumAdquision.Text) & ",'" & UCase(txtNumdoc.Text) & "','" & dcmbCodProveedor.Text & "', '" & Format(ff, "yyyy-mm-dd") & "', '" & UCase(txtObs.Text) & "' ," & _
                 "                          CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsConsu.Ejecutar (strSql)
    
        For aux = 1 To i - 1
        strSql = " INSERT INTO det_adquisicion_af (emp_codigo, adq_codigo,act_fij_codigo," & _
                 "                               det_adq_af_fechamod,det_adq_af_usumod)" & _
                 "                       VALUES ('" & strEmpresa & "' ,'" & txtNumAdquision & "', " & _
                 "                          '" & VSFDetalle.TextMatrix(aux, 1) & "', " & _
                 "                              CURRENT_TIMESTAMP, '" & strUsuario & "')"
        clsCon_Def.Ejecutar (strSql)
        Next
        End If
     End If
     Unload Me
     frmVerAdqActFij.Show
End Sub

Private Sub CmdSalir_Click()
   Unload Me
   frmVerAdqActFij.Show

End Sub

Private Sub dcmbCodProveedor_Change()
On Error GoTo errhandler
 'Muestra el nombre del proveedor relacionado con el código seleccionado
' o ingresado en el combo Proveedores al momento de hacer un cambio en el combo
If dcmbCodProveedor.Text = "" Then
Exit Sub
End If
    clsCon_Prove.adorec_Def.MoveFirst
    clsCon_Prove.adorec_Def.Find "per_codigo = '" & dcmbCodProveedor & "'", , adSearchForward

    If clsCon_Prove.adorec_Def.EOF = False Then
        'Muestra los datos del proveedor tales como: Nombres, Apellidos, Dirección, etc.
        txtNomProveedor.Text = clsCon_Prove.adorec_Def("nombre")
        txtRucProveedor.Text = clsCon_Prove.adorec_Def("per_ruc")
        txtDirProveedor.Text = clsCon_Prove.adorec_Def("per_direccion")
        txtTelProveedor.Text = clsCon_Prove.adorec_Def("per_telf")
        txtFaxProveedor.Text = clsCon_Prove.adorec_Def("per_fax")
        txtNomProveedor.Tag = clsCon_Prove.adorec_Def("par_texto")
    Else
        'MsgBox "No existe el Proveedor ingresado", vbInformation, "SisAdmi - PROVEEDOR"
        borrar_datos
    End If

    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal sql_server_connect and " & vbCrLf & _
                   " not a sql_server_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub Form_Load()

Dim d As String
Dim m As Integer
Dim Y As String
Dim ff As Variant
Dim var As Integer
'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6) + 500
    'Inicializa la clase con la conexión activa a la base de datos
    clsConsu.Inicializar AdoConn
    clsCon_Aux.Inicializar AdoConn
    clsCon_Act.Inicializar AdoConn
    clsCon_Iva.Inicializar AdoConn
    clsCon_Prove.Inicializar AdoConn

    'Descompone la fecha actual  en día, mes y año
    d = CStr(Day(Date))
    m = Month(Date)
    Y = CStr(Year(Date))

    cmbDia.Text = d
    cmbAño.Text = Y
    For var = 0 To 11
        If (cmbMes.ItemData(var) = m) Then
            cmbMes.Text = cmbMes.List(var)
            Exit For
        End If
    Next var

    'Busco el valor del IVA para la compra
    strSql = " SELECT par_numero,par_texto " & _
             " FROM  parametro " & _
             " WHERE par_codigo ='IVAC' AND emp_codigo='" & strEmpresa & "' "
    clsCon_Iva.Ejecutar (strSql)
    If clsCon_Iva.adorec_Def.EOF Then
        lblIva.Caption = "IVA 0%"
        txtIVA.Tag = " 0 "
    Else
        lblIva.Caption = "IVA " & Format(clsCon_Iva.adorec_Def.Fields("par_numero").Value, "###0.00") & "%"
        txtIVA.Tag = clsCon_Iva.adorec_Def.Fields("par_numero")
    End If

    'Consulta del nùmero de adquisicion último, se agrega uno para el nueva adquisicion
    strSql = " SELECT COALESCE(max(adq_codigo),0) as num " & _
             " FROM adquisicion " & _
             " WHERE emp_codigo = '" & strEmpresa & "' "
    clsConsu.Ejecutar (strSql)
    
    If clsConsu.adorec_Def.EOF Then
        txtNumAdquision.Text = "1"
    Else
        txtNumAdquision.Text = clsConsu.adorec_Def.Fields(0).Value + 1
    End If
    'Consulto los activos fijos de la empresa
    strSql = " CREATE TEMPORARY TABLE act_fij " & _
             " SELECT act_fij_codigo " & _
             " FROM det_adquisicion_af " & _
             " WHERE emp_codigo = '" & strEmpresa & "'" & _
             " ORDER BY act_fij_codigo "
    clsCon_Act.Ejecutar (strSql)
    
    strSql = " SELECT activo_fijo.act_fij_codigo, act_fij_nombre " & _
             " FROM activo_fijo LEFT JOIN act_fij ON activo_fijo.act_fij_codigo=act_fij.act_fij_codigo " & _
             " WHERE activo_fijo.emp_codigo = '" & strEmpresa & "'" & _
             " AND act_fij.act_fij_codigo IS NULL " & _
             " ORDER BY activo_fijo.act_fij_codigo "
    clsCon_Act.Ejecutar (strSql)
    strSql = " DROP TABLE act_fij"
    clsCon_Prove.Ejecutar (strSql)
    'Ejecuta un SQL par ver los datos de proveedores en la base de datos
     strSql = " SELECT per_codigo, concat(per_apellido,' ',per_nombre) as nombre,  " & _
              " per_direccion,per_ruc, per_telf, per_fax, par_texto " & _
              " FROM persona " & _
              " INNER JOIN parametro  " & _
              " ON persona.emp_codigo = parametro.emp_codigo " & _
              " WHERE persona.emp_codigo= '" & strEmpresa & "' " & _
              " AND cat_p_tipo = 'P'" & _
              " AND par_codigo='CXP'" & _
              " ORDER BY per_codigo,per_apellido,per_nombre "
    clsCon_Prove.Ejecutar (strSql)
    'Muestra los códigos de los proveedores en el combobox de códigos de proveedores

    If (clsCon_Prove.adorec_Def.RecordCount = 0) Then
        MsgBox "No existen Proveedores ingresados en el Sistema", vbInformation, "SisAdmi"
        Exit Sub
    Else
        Set dcmbCodProveedor.RowSource = clsCon_Prove.adorec_Def.DataSource
        dcmbCodProveedor.ListField = "per_codigo"
    End If
    
    If Not clsCon_Act.adorec_Def.EOF Then
        VSFDetalle.ColComboList(1) = VSFDetalle.BuildComboList(clsCon_Act.adorec_Def, "*act_fij_codigo,act_fij_nombre", "act_fij_codigo")
        VSFDetalle.ColComboList(2) = VSFDetalle.BuildComboList(clsCon_Act.adorec_Def, "act_fij_codigo,*act_fij_nombre", "act_fij_nombre")
   
     Else
        VSFDetalle.Clear 1
        VSFDetalle.Rows = 2
    End If
        

    'Insertamos el botón de eliminar en cada una de las filas

    'Inizializa el flexgrid
    VSFDetalle.Editable = flexEDKbdMouse
    VSFDetalle.AllowUserResizing = flexResizeBoth

    ' Agrega un botón en el grid
    VSFDetalle.Cell(flexcpPicture, 1, 0) = imgBtnUp
    VSFDetalle.Cell(flexcpPictureAlignment, 1, 0) = flexAlignRightCenter
    cmdAceptar.Enabled = False
End Sub
Private Sub txtTotal_Change()
If txtTotal <> "" Then
cmdAceptar.Enabled = True
End If
End Sub
Private Sub VSFDetalle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 3 Then
Cancel = True
End If
If Col = 4 Then
Cancel = True
End If
End Sub
Private Sub VSFDetalle_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    ' get cell that was clicked
    Dim r&, c&
    r = VSFDetalle.MouseRow
    c = VSFDetalle.MouseCol
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    If (c <> 0 Or r = (VSFDetalle.Rows) Or VSFDetalle.Rows = 2) Then Exit Sub
    ' make sure the click was on a cell with a button
    If r > 0 Then
        If c > 1 Then
            If VSFDetalle.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
        End If
        ' make sure the click was on the button (not just on the cell)
        ' note: this works for right-aligned buttons
        Dim d!
        d = VSFDetalle.Cell(flexcpLeft, r, c) + VSFDetalle.Cell(flexcpWidth, r, c) - X
        If d > imgBtnDn.Width Then Exit Sub
        If r > 0 Then
        ' click was on a button: do the work
        VSFDetalle.Cell(flexcpPicture, r, c) = imgBtnDn
        Mensaje = "Desea eliminar la fila " & r & " ?"    ' Define el mensaje.
        Estilo = vbYesNo + vbInformation + vbDefaultButton2   ' Define los botones.
        Título = "SisAdmi - Proyecto de Ventas"   ' Define el título.
        Respuesta = MsgBox(Mensaje, Estilo, Título)
        'Recorro el FlexGrid para poner números a las filas
        If Respuesta = vbYes Then
            Dim i As Integer
            VSFDetalle.RemoveItem (r)
            PonerBotones
            CalcuTotal
        Else
            VSFDetalle.Cell(flexcpPicture, r, c) = imgBtnUp
        End If
    End If
End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub
'Private Sub VSFDetalle_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
''para que no se pueda escribir en las columnas que se indica
'  If Col = 3 Or Col = 4 Then
'            SendKeys "{TAB}"
'            Else
'            SendKeys "+{TAB}"
'            Cancel = True
'  End If
'End Sub
Private Sub VSFDetalle_CellChanged(ByVal Row As Long, ByVal Col As Long)

    With VSFDetalle
        If .TextMatrix(Row, Col) <> "" Then
            If Col = 1 Then
                 clsCon_Act.Filtrar ("act_fij_codigo = '" & .TextMatrix(Row, 1) & "'")
                     .TextMatrix(Row, 2) = clsCon_Act.adorec_Def("act_fij_nombre")
                 clsCon_Act.QuitarFiltro
             End If
             If Col = 2 Then
                 clsCon_Act.Filtrar ("act_fij_nombre = '" & .TextMatrix(Row, 2) & "'")
                     .TextMatrix(Row, 1) = clsCon_Act.adorec_Def("act_fij_codigo")
                 clsCon_Act.QuitarFiltro
             End If
         End If
    End With
    If Row > 0 And Col = 1 Or Col = 2 Then
        strSql = " SELECT act_fij_descripcion, act_fij_valor " & _
                 " FROM activo_fijo " & _
                 " WHERE emp_codigo = '" & strEmpresa & "'AND act_fij_codigo = '" & VSFDetalle.TextMatrix(Row, 1) & "'"
        clsConsu.Ejecutar (strSql)
    
        If Not clsConsu.adorec_Def.EOF Then
            VSFDetalle.TextMatrix(Row, 3) = clsConsu.adorec_Def("act_fij_descripcion")
            VSFDetalle.TextMatrix(Row, 4) = clsConsu.adorec_Def("act_fij_valor")
            CalcuTotal
        Else
            VSFDetalle.TextMatrix(Row, 3) = ""
            VSFDetalle.TextMatrix(Row, 4) = ""
        End If
    End If
    'Verifica que no se ingresen dos activos Fijos iguales en el grid
    If Row > 1 And Col = 1 Or Col = 2 Then
        With VSFDetalle
            For i = 1 To .Rows - 1
                For j = i + 1 To .Rows - 1
                    If .TextMatrix(i, 1) = .TextMatrix(j, 1) Then
                        MsgBox "El Activo Fijo ya ha sido ingresado, ingrese un Activo Fijo diferente", vbExclamation, "SisAdmi"
                        .TextMatrix(Row, 1) = ""
                        .TextMatrix(Row, 2) = ""
                        .TextMatrix(Row, 3) = ""
                        .TextMatrix(Row, 4) = ""
                    End If
                    If j >= .Rows - 1 Then
                        Exit For
                    End If
                Next j
            Next i
        End With
    End If
End Sub

Private Sub limpiarFxGD()
'función que recorre el flexGrid y limpia los campos
    Dim X, Y  As Integer
    VSFDetalle.Tag = "N"
    VSFDetalle.Clear 1
    VSFDetalle.Rows = 2
    VSFDetalle.Tag = "T"
End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    'Se da un tab al presionar enter para que al ingresar un dato pase al siguiente campo
'    If KeyCode = vbKeyReturn Then
'        KeyCode = 0
'        SendKeys "{TAB}"
'    End If
'
'End Sub
Private Sub VSFDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
'Hace que cuando llegue al final del grid, presiona las teclas: enter, tab, izquierda y abajo , se cree otra fila y ponga los botones correspondientes
    
    If VSFDetalle.Row = VSFDetalle.Rows - 1 And (KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight) Then
       If VSFDetalle.TextMatrix(VSFDetalle.Row, 1) <> "" And (VSFDetalle.TextMatrix(VSFDetalle.Row, 3) <> "" Or VSFDetalle.TextMatrix(VSFDetalle.Row, 4) <> "") Then
            VSFDetalle.AddItem ""
            VSFDetalle.TextMatrix(VSFDetalle.Rows - 1, 0) = VSFDetalle.Rows - 1
            VSFDetalle.Cell(flexcpPicture, (VSFDetalle.Rows - 1), 0) = imgBtnUp
            VSFDetalle.Cell(flexcpPictureAlignment, (VSFDetalle.Rows - 1), 0) = flexAlignRightCenter
            PonerBotones
        End If
    End If
End Sub


Private Sub VSFDetalle_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'Valido que solo se pueda dar enter en el campo Desc. Activo fijo
    If (Col = 3) Or (Col = 4) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub

'Private Sub VSFDetalle_RowColChange()
'    'Se envía un espacio en blanco al recorrer el flexGrid para desplegar los combos que existan
'    SendKeys " "
'    TxtTotal.Tag = ""
'End Sub

Public Sub borrar_datos()

    txtNomProveedor.Text = ""
    txtDirProveedor.Text = ""
    txtRucProveedor.Text = ""
    txtTelProveedor.Text = ""
    txtFaxProveedor.Text = ""
    txtSubTotal.Text = ""
    lblIva.Caption = "IVA 0 %"
    txtIVA.Text = ""
    txtTotal.Text = ""
    txtObs.Text = ""
    txtNumdoc.Text = ""
    dcmbCodProveedor.Text = ""

End Sub

