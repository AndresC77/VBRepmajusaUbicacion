VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmSelProducto 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   9510
   Begin VB.CommandButton cmdMant 
      Caption         =   "Listado"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   128
      TabIndex        =   12
      Top             =   120
      Width           =   9255
      Begin VB.CheckBox chkBaja 
         BackColor       =   &H00DDDDDD&
         Enabled         =   0   'False
         Height          =   240
         Left            =   5880
         TabIndex        =   7
         Top             =   2040
         Width           =   300
      End
      Begin VB.TextBox txtLinea 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   3225
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   3225
      End
      Begin VB.TextBox txtCosto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox lstGrupo 
         Height          =   1320
         Left            =   5880
         TabIndex        =   6
         Top             =   720
         Width           =   3225
      End
      Begin VB.TextBox txtUnidad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   3225
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcmbNombre 
         Height          =   330
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Producto:"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4200
         TabIndex        =   19
         Top             =   420
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Línea:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1125
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupos de Producto:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   15
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5040
         TabIndex        =   14
         Top             =   2055
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6368
      TabIndex        =   11
      Top             =   2655
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4808
      TabIndex        =   10
      Top             =   2655
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3248
      TabIndex        =   9
      Top             =   2655
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1688
      TabIndex        =   8
      Top             =   2655
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion del Producto y poder modificarlo o                  #
'#  crear o eliminar paises                                                     #
'#  frmSelPais V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar los productos que al momento estan                   #
'#  ingresadas en el sistema. Desde esta ventana se puede crear un nuevo        #
'#  producto o modificar o eliminar los productos ya creados.                   #
'#  Desde esta ventana se llama a la ventana frmProducto en la que se crea      #
'#  y modifica los productos                                                    #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#     producto: En esta tabla se almacenan los nuevos productos y se           #
'#               modifican los datos de estos.                                  #
'#        grupo: En esta tabla se sacan los grupos a los que puede pertenecer   #
'#               los diferentes productos en sus diferentes niveles.            #
'#       marca : En esta tabla se sacan las marcas a las que se puede asignar a #
'#               los diferentes productos.                                      #
'#       linea : En esta tabla se sacan las lineas a las que se puede asignar a #
'#               los diferentes productos.                                      #
'#      unidad : En esta tabla se sacan las unidades de medida que se puede     #
'#               asignar los productos.                                         #
'#                                                                              #
'#  Procedimientos INTERNOS:                                                    #
'#     LlenarListaGrupo(strCod As String, intNiv As Integer)                    #
'#               Proceso para llenar la lista el grupo y sub grupos a los       #
'#               que pertenece el producto                                      #
'#                                                                              #
'#  Procedimientos EXTERNOS:                                                    #
'#                                                                              #
'#  Objetos de la forma:                                                        #
'#    clsCon_Def clsConsulta: Objeto para consultar a la base de datos          #
'#    strGrupo String: Variable que tiene el codio del grupo del producto       #
'#    intNivel Integer: Variable que tiene el nivel máximo de los grupos        #
'#                                                                              #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private clsCon_Def As clsConsulta
Private strGrupo As String
Private intNivel As Integer

Private Sub cmdMant_Click()
    frmCambioProductos.Show vbModal
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

Private Sub cmdEliminar_Click()
    Dim strSql As String
    ' Consulta para conocer si hay existencia del produto
    strSql = " SELECT count(*) As Num " & _
             " FROM existencia " & _
             " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
             " AND exi_cantidad > 0 " & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar (strSql)
    ' Si existen existencias no se elimina
    If clsCon_Def.adorec_Def("Num") > 0 Then
        MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
    ' Si no existen productos se elimina
    Else
        ' Consulta para conocer si hay productos en el detalle de ingresos
        strSql = " SELECT count(*) As Ing " & _
                 " FROM det_ingreso " & _
                 " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar (strSql)
        ' Si existen ingresos no se elimina
        If clsCon_Def.adorec_Def("Ing") > 0 Then
            MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
        ' Si no existen productos se elimina
        Else
            ' Consulta para conocer si hay productos en el detalle de egresos
            strSql = " SELECT count(*) As Egr " & _
                     " FROM det_egreso " & _
                     " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsCon_Def.Ejecutar (strSql)
            ' Si existen egresos no se elimina
            If clsCon_Def.adorec_Def("Egr") > 0 Then
                MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
            ' Si no existen productos se elimina
            Else
                ' Consulta para conocer si hay productos en detalle de pedidos
                strSql = " SELECT count(*) As Ped " & _
                         " FROM det_pedido " & _
                         " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                         " AND emp_codigo='" & strEmpresa & "'"
                clsCon_Def.Ejecutar (strSql)
                ' Si existen pedidos no se elimina
                If clsCon_Def.adorec_Def("Ped") > 0 Then
                    MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
                ' Si no existen productos se elimina
                Else
                    ' Consulta para conocer si hay productos en el detalle productos compuestos
                    strSql = " SELECT count(*) As Com " & _
                             " FROM det_prd_com " & _
                             " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                             " AND emp_codigo='" & strEmpresa & "'"
                    clsCon_Def.Ejecutar (strSql)
                    ' Si existen productos compuestos no se elimina
                    If clsCon_Def.adorec_Def("Com") > 0 Then
                        MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
                    ' Si no existen productos se elimina
                    Else
                        ' Consulta para conocer si hay productos en el detalle de cotización
                        strSql = " SELECT count(*) As Cot " & _
                                 " FROM det_cotizacion " & _
                                 " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                                 " AND emp_codigo='" & strEmpresa & "'"
                        clsCon_Def.Ejecutar (strSql)
                        ' Si existen cotizaciones no se elimina
                        If clsCon_Def.adorec_Def("Cot") > 0 Then
                            MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
                        ' Si no existen productos se elimina
                        Else
                            ' Consulta para conocer si hay productos en el detalle de egresos
                            strSql = " SELECT count(*) As Bck " & _
                                     " FROM det_backorder " & _
                                     " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                                     " AND emp_codigo='" & strEmpresa & "'"
                            clsCon_Def.Ejecutar (strSql)
                            ' Si existen backorders no se elimina
                            If clsCon_Def.adorec_Def("Bck") > 0 Then
                                MsgBox "No Puede eliminar este producto", vbInformation, "Eliminación"
                            ' Si no existen productos se elimina
                            Else
                                'Se elimina de existencias
                                strSql = " DELETE " & _
                                         " FROM existencia " & _
                                         " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                                         " AND emp_codigo='" & strEmpresa & "'"
                                clsCon_Def.Ejecutar (strSql)
                                
                                'Se elimina de lista de precios
                                strSql = " DELETE " & _
                                         " FROM lista_precio_p " & _
                                         " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                                         " AND emp_codigo='" & strEmpresa & "'"
                                clsCon_Def.Ejecutar (strSql)
                                
                                'Se elimina de lista de producto
                                strSql = " DELETE " & _
                                         " FROM producto " & _
                                         " WHERE prd_codigo = '" & dcmbCodigo.Text & "' " & _
                                         " AND emp_codigo='" & strEmpresa & "'"
                                clsCon_Def.Ejecutar (strSql)
                                MsgBox "Producto eliminado", vbInformation, "Eliminación"
                            End If
                        
                        End If
                    End If
                End If
            End If
        End If
    End If
    'Consulta para actualizar los combos
    strSql = " SELECT producto.prd_codigo, producto.prd_nombre, producto.prd_costo, " & _
             " lin_nombre, producto.prd_baja , unidad.uni_nombre, marca.mar_nombre, " & _
             " grupo.gru_nivel, producto.gru_codigo " & _
             " FROM ((((producto INNER JOIN unidad " & _
             " ON producto.uni_codigo = unidad.uni_codigo AND producto.emp_codigo = unidad.emp_codigo) " & _
             " INNER JOIN grupo ON producto.emp_codigo = grupo.emp_codigo " & _
             " AND producto.gru_codigo = grupo.gru_codigo) " & _
             " INNER JOIN linea ON producto.emp_codigo = linea.emp_codigo " & _
             " AND producto.lin_codigo = linea.lin_codigo) " & _
             " INNER JOIN marca ON producto.emp_codigo = marca.emp_codigo " & _
             " AND producto.mar_codigo = marca.mar_codigo) " & _
             " WHERE producto.emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar (strSql)
    dcmbCodigo.ListField = "prd_codigo"
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "prd_nombre"
    dcmbNombre.BoundColumn = "prd_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbCodigo.Text = ""
    
End Sub

Private Sub cmdModificar_Click()
' Modifica los datos de un producto, se manda a la variable Tag del formulario una bandera para
' conocer que se esta modificando y ademas se envia el código del producto que se modificará
    Dim i As Integer
    Dim intPos As Integer
    Dim strCodAux As String
    frmProducto.Show
'    frmProducto.txtCodigo.Text = Me.dcmbCodigo.Text
'    frmProducto.txtNombre.Text = Me.dcmbNombre.Text
'    frmProducto.chkBaja.Value = Me.chkBaja.Value
'    frmProducto.dcmbUnidad.Text = Me.txtUnidad.Text
'    frmProducto.dcmbLinea.Text = Me.txtLinea.Text
'    frmProducto.dcmbMarca.Text = Me.txtMarca.Text
'    frmProducto.txtCosto.Text = Me.txtCosto.Text
'    ' Ingresa el numero de filas del vsflexgrid
'    frmProducto.vsfgGrupo.Rows = intNivel
    strCodAux = strGrupo & "."
    ' Llena el vsflexgrid
    For i = 1 To intNivel
        intPos = InStrRev(strCodAux, ".") - 1
        If intPos <> -1 Then
            strCodAux = Left(strCodAux, intPos)
        Else
            Exit For
        End If
'        frmProducto.vsfgGrupo.TextMatrix(intNivel - i, 1) = strCodAux
'        frmProducto.vsfgGrupo.TextMatrix(intNivel - i, 0) = Me.lstGrupo.List(intNivel - i)
    Next i
'    frmProducto.Tag = "M"
End Sub

Private Sub cmdNuevo_Click()
' Crea un nuevo producto, se manda a la variable Tag del formulario una bandera para
' conocer que se esta ingresará un nuevo producto
    frmProducto.Show
'    frmProducto.vsfgGrupo.Rows = intNivel
'    frmProducto.Tag = "N"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub dcmbCodigo_Change()
' Chequea el producto seleccionado y escribe su nombre en el combo
    Dim strComparar As String
    On Error GoTo errhandler
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "prd_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("prd_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            chkBaja.Value = clsCon_Def.adorec_Def("prd_baja")
            txtUnidad.Text = clsCon_Def.adorec_Def("uni_nombre")
            txtLinea.Text = clsCon_Def.adorec_Def("lin_nombre")
            txtMarca.Text = clsCon_Def.adorec_Def("mar_nombre")
            txtCosto.Text = clsCon_Def.adorec_Def("prd_costo")
            strGrupo = clsCon_Def.adorec_Def("gru_codigo")
            ' Llena el listbox
            LlenarListaGrupo strGrupo, intNivel
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Else
            dcmbNombre.Text = ""
            dcmbNombre.BoundText = ""
            chkBaja.Value = 0
            txtUnidad.Text = ""
            txtLinea.Text = ""
            txtMarca.Text = ""
            txtCosto.Text = ""
            strGrupo = ""
            lstGrupo.Clear
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
        End If
        dcmbCodigo.Tag = ""
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

Private Sub dcmbNombre_Change()
'Cambia el valor del codigo para actualizar este y la descripcion
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        End If
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Cambia el valor del codigo para actualizar este y la descripcion
    dcmbCodigo.Text = dcmbNombre.BoundText
End Sub

Private Sub dcmbNombre_KeyUp(KeyCode As Integer, Shift As Integer)
'Cambia el valor del codigo para actualizar este y la descripcion
     If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dcmbCodigo.Text = dcmbNombre.BoundText
    End If
End Sub
Private Sub Form_Activate()
' Actualiza la lista de productos al volver al formulario
    clsCon_Def.Actualizar
    dcmbCodigo.ListField = "prd_codigo"
    Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
    dcmbNombre.ListField = "prd_nombre"
    dcmbNombre.BoundColumn = "prd_codigo"
    Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource

End Sub

Private Sub Form_Load()
    Dim strSql As String
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (Me.Height / 6)
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn
        strSql = " SELECT COALESCE(max(gru_nivel),0) as num " & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' "
        clsCon_Def.Ejecutar (strSql)
        'Si exsiten grupos de productos
        If Not clsCon_Def.adorec_Def.EOF Then
            intNivel = clsCon_Def.adorec_Def("num")
        'Consulta los productos que estan disponibles
            strSql = " SELECT producto.prd_codigo, producto.prd_nombre, producto.prd_costo, " & _
                     " lin_nombre, producto.prd_baja , unidad.uni_nombre, marca.mar_nombre, " & _
                     " grupo.gru_nivel, producto.gru_codigo " & _
                     " FROM ((((producto INNER JOIN unidad " & _
                     " ON producto.uni_codigo = unidad.uni_codigo AND producto.emp_codigo = unidad.emp_codigo) " & _
                     " INNER JOIN grupo ON producto.emp_codigo = grupo.emp_codigo " & _
                     " AND producto.gru_codigo = grupo.gru_codigo) " & _
                     " INNER JOIN linea ON producto.emp_codigo = linea.emp_codigo " & _
                     " AND producto.lin_codigo = linea.lin_codigo) " & _
                     " INNER JOIN marca ON producto.emp_codigo = marca.emp_codigo " & _
                     " AND producto.mar_codigo = marca.mar_codigo) " & _
                     " WHERE producto.emp_codigo='" & strEmpresa & "' ORDER BY producto.prd_codigo"
            clsCon_Def.Ejecutar (strSql)
            dcmbCodigo.ListField = "prd_codigo"
            Set dcmbCodigo.RowSource = clsCon_Def.adorec_Def.DataSource
            dcmbNombre.ListField = "prd_nombre"
            dcmbNombre.BoundColumn = "prd_codigo"
            Set dcmbNombre.RowSource = clsCon_Def.adorec_Def.DataSource
        Else
            cmdNuevo.Enabled = False
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

Private Sub txtCosto_Change()
    ' Pone los decimales en el txt del costo
    txtCosto.Text = FormatoD4(txtCosto.Text)
End Sub
Private Sub LlenarListaGrupo(strCod As String, intNiv As Integer)
' Proceso para llenar la lista el grupo y sub grupos a los que pertenece el producto
    Dim clsCon_Aux As clsConsulta
    Dim strSql As String
    Dim strCodAux As String
    Dim intPos As Integer
    Dim i As Integer
    Set clsCon_Aux = New clsConsulta
    clsCon_Aux.Inicializar AdoConn
    lstGrupo.Clear
    ' Genera el SQL para consultar todos los grupos y subgrupos del producto
    strSql = " SELECT CONCAT(SPACE((gru_nivel-1)*4),gru_nombre) as gru_nombre " & _
             " FROM grupo " & _
             " WHERE (gru_codigo='" & strCod & "' "
    intPos = InStrRev(strCod, ".") - 1
    If intPos <> -1 Then
        strCodAux = Left(strCod, intPos)
        For i = 1 To intNiv - 1
            strSql = strSql & " OR gru_codigo='" & strCodAux & "' "
            intPos = InStrRev(strCodAux, ".") - 1
            If intPos <> -1 Then
                strCodAux = Left(strCodAux, intPos)
            Else
                Exit For
            End If
        Next i
    End If
    strSql = strSql & ") AND emp_codigo='" & strEmpresa & "' ORDER BY gru_codigo"
    clsCon_Aux.Ejecutar (strSql)
    clsCon_Aux.adorec_Def.MoveFirst
    ' Llena el listbox
    While Not clsCon_Aux.adorec_Def.EOF
        lstGrupo.AddItem clsCon_Aux.adorec_Def("gru_nombre")
        clsCon_Aux.adorec_Def.MoveNext
    Wend
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub
