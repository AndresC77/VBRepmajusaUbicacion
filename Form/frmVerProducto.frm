VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVerProducto 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Productos"
   ClientHeight    =   4710
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
   Icon            =   "frmVerProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   9510
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3968
      TabIndex        =   29
      Top             =   4200
      Width           =   1575
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
      Height          =   3975
      Left            =   128
      TabIndex        =   6
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtColeccion 
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   3225
      End
      Begin VB.CheckBox chkPrecio 
         BackColor       =   &H00DDDDDD&
         Enabled         =   0   'False
         Height          =   240
         Left            =   8400
         TabIndex        =   25
         Top             =   1680
         Width           =   300
      End
      Begin VB.ListBox lstPromo 
         Height          =   1530
         Left            =   3120
         TabIndex        =   24
         Top             =   2280
         Width           =   2985
      End
      Begin VB.ListBox lstGrupo 
         Height          =   1530
         Left            =   6240
         TabIndex        =   23
         Top             =   2280
         Width           =   2865
      End
      Begin VB.ListBox lstPrecios 
         Height          =   1530
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   2865
      End
      Begin VB.TextBox txtColor 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   3225
      End
      Begin VB.TextBox txtTalla 
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   3225
      End
      Begin VB.TextBox txtCosto 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.CheckBox chkBaja 
         BackColor       =   &H00DDDDDD&
         Enabled         =   0   'False
         Height          =   240
         Left            =   5880
         TabIndex        =   5
         Top             =   1680
         Width           =   300
      End
      Begin VB.TextBox txtLinea 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   3225
      End
      Begin VB.TextBox txtMarca 
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   3225
      End
      Begin VB.TextBox txtUnidad 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   3225
      End
      Begin MSDataListLib.DataCombo dcmbCodigo 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   240
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
         TabIndex        =   4
         Top             =   240
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colección:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4995
         TabIndex        =   28
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambio de Precio de Venta"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6240
         TabIndex        =   26
         Top             =   1695
         Width           =   1950
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promociones:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3120
         TabIndex        =   21
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupos de Producto:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6240
         TabIndex        =   20
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Precios:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   1005
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Talla:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5370
         TabIndex        =   16
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   195
         TabIndex        =   14
         Top             =   1725
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Producto:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4200
         TabIndex        =   11
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Línea:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   225
         TabIndex        =   10
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5250
         TabIndex        =   9
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De Baja:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5145
         TabIndex        =   8
         Top             =   1695
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmVerProducto"
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
Public strGrupo As String
Private intNivel As Integer
Public isNoChange As Boolean

Private Sub dcmbNombre_Validate(Cancel As Boolean)
    If dcmbNombre.MatchedWithList = False Then
        dcmbCodigo.Text = ""
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

Private Sub dcmbCodigo_Change()
' Chequea el producto seleccionado y escribe su nombre en el combo
    Dim strSql As String
    Dim strComparar As String
    If isNoChange = True Then Exit Sub 'Edu
    On Error GoTo errhandler
        clsCon_Def.adorec_Def.MoveFirst
        strComparar = "prd_codigo = '" & dcmbCodigo.Text & "'"
        clsCon_Def.adorec_Def.Find strComparar
        dcmbCodigo.Tag = "A"
        If clsCon_Def.adorec_Def.EOF = False Then
            dcmbNombre.Text = clsCon_Def.adorec_Def("prd_nombre")
            dcmbNombre.BoundText = dcmbCodigo.Text
            txtCosto.Text = clsCon_Def.adorec_Def("prd_costo")
            chkBaja.Value = clsCon_Def.adorec_Def("prd_baja")
            chkPrecio.Value = clsCon_Def.adorec_Def("prd_cambia_precio")
            txtUnidad.Text = clsCon_Def.adorec_Def("uni_nombre")
            txtLinea.Text = clsCon_Def.adorec_Def("lin_nombre")
            txtMarca.Text = clsCon_Def.adorec_Def("mar_nombre")
            txtColor.Text = IIf(Not IsNull(clsCon_Def.adorec_Def("col_nombre")), clsCon_Def.adorec_Def("col_nombre"), "")
            'txtCosto.Text = clsCon_Def.adorec_Def("prd_costo")
            txtColeccion.Text = IIf(Not IsNull(clsCon_Def.adorec_Def("clc_nombre")), clsCon_Def.adorec_Def("clc_nombre"), "")
            txtTalla.Text = IIf(Not IsNull(clsCon_Def.adorec_Def("tal_nombre")), clsCon_Def.adorec_Def("tal_nombre"), "")
            strGrupo = clsCon_Def.adorec_Def("gru_codigo")
            ' Llena el listbox
            'LlenarListaGrupo strGrupo, intNivel
            LlenarList (strGrupo) 'Edu: En vez del la funcion LlenarListaGrupo
            LlenarPrecio dcmbCodigo.Text
            LlenarPromo dcmbCodigo.Text
        Else
            dcmbNombre.Text = ""
            dcmbNombre.BoundText = ""
            chkBaja.Value = 0
            chkPrecio.Value = 0
            txtCosto.Text = "" 'Edu
            txtUnidad.Text = ""
            txtLinea.Text = ""
            txtMarca.Text = ""
            txtColor.Text = ""
            'txtCosto.Text = ""
            txtColeccion.Text = ""
            txtTalla.Text = ""
            strGrupo = ""
            lstGrupo.Clear
            lstPrecios.Clear
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
    If isNoChange = True Then Exit Sub 'Edu
    If dcmbCodigo.Tag <> "A" Then
        If dcmbNombre.MatchedWithList = True Then
            dcmbCodigo.Text = dcmbNombre.BoundText
        Else
            dcmbCodigo.Text = ""
        End If
    End If
End Sub

Private Sub dcmbNombre_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT max(gru_nivel) as num " & _
                 " FROM grupo " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " GROUP BY emp_codigo"
        clsCon_Def.Ejecutar (strSql)
        'Si exsiten grupos de productos
        If Not clsCon_Def.adorec_Def.EOF Then
            intNivel = clsCon_Def.adorec_Def("num")
        'Consulta los productos que estan disponibles
            strSql = " SELECT producto.prd_codigo, producto.prd_costo, producto.prd_nombre, producto.prd_costo, " & _
                     " lin_nombre, producto.prd_baja , unidad.uni_nombre, marca.mar_nombre, " & _
                     " grupo.gru_nivel, producto.gru_codigo,tal_nombre,col_nombre,prd_cambia_precio,clc_nombre " & _
                     " FROM ((((producto INNER JOIN unidad " & _
                     " ON producto.uni_codigo = unidad.uni_codigo AND producto.emp_codigo = unidad.emp_codigo) " & _
                     " INNER JOIN grupo ON producto.emp_codigo = grupo.emp_codigo " & _
                     " AND producto.gru_codigo = grupo.gru_codigo) " & _
                     " INNER JOIN linea ON producto.emp_codigo = linea.emp_codigo " & _
                     " AND producto.lin_codigo = linea.lin_codigo) " & _
                     " INNER JOIN marca ON producto.emp_codigo = marca.emp_codigo " & _
                     " AND producto.mar_codigo = marca.mar_codigo) " & _
                     " INNER JOIN talla ON producto.emp_codigo = talla.emp_codigo " & _
                     " AND producto.tal_codigo = talla.tal_codigo " & _
                     " INNER JOIN color ON producto.emp_codigo = color.emp_codigo " & _
                     " AND producto.col_codigo = color.col_codigo " & _
                     " LEFT JOIN coleccion ON producto.emp_codigo = coleccion.emp_codigo " & _
                     " AND producto.clc_codigo = coleccion.clc_codigo " & _
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

Sub LlenarList(codigo As String, Optional EnModific As Boolean = False)
'Llena lstGrupo con informacion de grupos
  Dim clsCon_Aux As New clsConsulta
  Dim i As Integer
  Dim st As String
  If EnModific = False Then lstGrupo.Clear
  clsCon_Aux.Inicializar AdoConn, AdoConnMaster
  For i = 0 To GetNumNiv(codigo)
    clsCon_Aux.Ejecutar "SELECT * " & _
                         "FROM grupo " & _
                         "WHERE gru_codigo='" & GetCod(codigo, i) & "' AND gru_nivel=" & i + 1
    If EnModific = False Then
      lstGrupo.AddItem st & clsCon_Aux.adorec_Def("gru_nombre")
      st = st & "  "
    Else
      'frmProducto.vsfgGrupo.TextMatrix(i, 0) = clsCon_Aux.adorec_Def("gru_nombre")
      'frmProducto.vsfgGrupo.TextMatrix(i, 1) = GetCod(Codigo, i)
    End If
  Next i
End Sub

Sub LlenarPrecio(codigo As String)
'Llena lstGrupo con informacion de grupos
  Dim clsCon_Aux As New clsConsulta
  Dim i As Integer
  Dim st As String
  lstPrecios.Clear
  clsCon_Aux.Inicializar AdoConn, AdoConnMaster
  clsCon_Aux.Ejecutar "SELECT CONCAT(lis_pre_descripcion, ' - ' ,lis_pre_p_precio) as val " & _
                    "FROM lista_precio_p INNER JOIN lista_precio ON lista_precio_p.lis_pre_codigo=lista_precio.lis_pre_codigo AND lista_precio_p.emp_codigo=lista_precio.emp_codigo " & _
                    "WHERE lista_precio_p.emp_codigo='" & strEmpresa & "' " & _
                    "AND prd_codigo='" & codigo & "' " & _
                    " ORDER BY lis_pre_descripcion "
  While Not clsCon_Aux.adorec_Def.EOF
      lstPrecios.AddItem clsCon_Aux.adorec_Def("val")
      clsCon_Aux.adorec_Def.MoveNext
  Wend
End Sub

Sub LlenarPromo(codigo As String)
'Llena lstGrupo con informacion de grupos
  Dim clsCon_Aux As New clsConsulta
  Dim i As Integer
  Dim st As String
  lstPromo.Clear
  clsCon_Aux.Inicializar AdoConn, AdoConnMaster
  clsCon_Aux.Ejecutar "SELECT CONCAT(LEFT(prd_pro_fechaini,10),' al ',LEFT(prd_pro_fechafin,10),' : ',prd_pro_porcentaje,'%') as val " & _
                    "FROM producto_promo " & _
                    "WHERE prd_codigo='" & codigo & "' " & _
                    " ORDER BY prd_pro_fechaini,prd_pro_fechafin "
  While Not clsCon_Aux.adorec_Def.EOF
      lstPromo.AddItem clsCon_Aux.adorec_Def("val")
      clsCon_Aux.adorec_Def.MoveNext
  Wend
End Sub
Private Function GetCod(codigo As String, Nivel As Integer) As String
'Devuelve codigo a partir de una linea de codigo. Ej: Codigo: 01.02 , Nivel:1 -> Devuelve 01
  Dim i As Integer
  Dim j As Integer
  For i = 1 To Len(codigo)
    If Mid(codigo, i, 1) = "." Then j = j + 1
    If j > Nivel Then Exit Function
    GetCod = GetCod & Mid(codigo, i, 1)
  Next i
End Function
Private Function GetNumNiv(codigo As String) As Integer
'Retorna el nuvero de niveles de acuerdo al codigo especificado
  Dim i As Integer
  For i = 1 To Len(codigo)
    If Mid(codigo, i, 1) = "." Then GetNumNiv = GetNumNiv + 1
  Next i
End Function


Private Sub LlenarListaGrupo(strCod As String, intNiv As Integer)
' Proceso para llenar la lista el grupo y sub grupos a los que pertenece el producto
    Dim clsCon_Aux As clsConsulta
    Dim strSql As String
    Dim strCodAux As String
    Dim intPos As Integer
    Dim i As Integer
    Set clsCon_Aux = New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
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
    For i = 1 To intNiv
        lstGrupo.AddItem clsCon_Aux.adorec_Def("gru_nombre")
        clsCon_Aux.adorec_Def.MoveNext
    Next i
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
