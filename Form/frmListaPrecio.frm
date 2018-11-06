VERSION 5.00
Begin VB.Form frmListaPrecio 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Lista de Precio"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListaPrecio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3510
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Lista de Precio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   128
      TabIndex        =   6
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txtPolitica 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   960
         Width           =   1920
      End
      Begin VB.TextBox txtDesPolitica 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Política:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1005
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Política Fijada:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   1365
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1808
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   248
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmListaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma de ingreso y modificación de Listas de Precios de productos, para     #
'#  definir los precios de acuerdo a las categorias.                            #
'#  frmListaPrecio V1.0                                                         #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para la creación y modificación de las listas de precios y sus      #
'#  políticas. Permitirá almacenar en la base de datos nuevas listas y modficar #
'#  sus descripciones y políticas, dependiendo de la propiedad Tag,             #
'#  la cual se cambiará en la ventana frmSelListaPrecio y desde esta se llamará #
'#  a esta ventana.                                                             #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#  lista_precio:En esta tabla se almacenan las listas de precios y se          #
'#               modifican los datos de estas.                                  #
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

Private clsCon_Def As clsConsulta
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCon_Def = Nothing
End Sub

Private Sub cmbAceptar_Click()
    Dim strSQL As String
    ' Si se esta ingresando una nueva lista
    If Me.Tag = "N" Then
    ' Almacenamiento de los datos de la ueva lista
        strSQL = " INSERT INTO lista_precio(lis_pre_codigo,emp_codigo,lis_pre_descripcion,lis_pre_politica,lis_pre_fijada,lis_pre_fechamod,lis_pre_usumod) " & _
                 " VALUES ('" & UCase(txtCodigo.Text) & "','" & strEmpresa & "','" & UCase(txtDescripcion.Text) & "','" & txtPolitica.Text & "',0, " & _
                 " CURRENT_TIMESTAMP, '" & strUsuario & "')"
        On Error GoTo errhandler
            clsCon_Def.Ejecutar (strSQL), "M"
    ' Almacenamiento de los datos de la nueva lista en las listas de precios de productos
            strSQL = " INSERT INTO lista_precio_p " & _
                     " SELECT '" & UCase(txtCodigo.Text) & "', prd_codigo, emp_codigo, " & _
                     " prd_costo/(1-" & txtPolitica.Text & "/100.00), " & _
                     " '" & txtPolitica.Text & "',0,0,CURRENT_TIMESTAMP, '" & strUsuario & "' " & _
                     " FROM producto WHERE emp_codigo='" & strEmpresa & "' "
    ' Si se esta modificando la Lista
    ElseIf Me.Tag = "M" Then
    'Almacenamiento de los cambios realizados a la lista
        strSQL = " UPDATE lista_precio " & _
                 " SET lis_pre_descripcion='" & UCase(txtDescripcion.Text) & "',lis_pre_fijada=IF(lis_pre_politica=" & txtPolitica.Text & ",lis_pre_fijada,0)" & _
                 ",lis_pre_politica='" & txtPolitica.Text & _
                 "',lis_pre_fechamod=CURRENT_TIMESTAMP,lis_pre_usumod='" & strUsuario & "' " & _
                 " WHERE lis_pre_codigo='" & txtCodigo.Text & "' AND emp_codigo='" & strEmpresa & "'"
    End If
    On Error GoTo errhandler
        clsCon_Def.Ejecutar (strSQL), "M"
        Unload Me
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

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' De acuerdo a la propiedad Tag escribe el título de la ventana
    If Me.Tag = "M" Then
        Me.Caption = "Modificar datos de la Lista"
        txtCodigo.Enabled = False
    ElseIf Me.Tag = "N" Then
        Me.Caption = "Ingreso de Nueva Lista"
        txtCodigo.Enabled = False
        'Consulta el codigo de la lisya que tocaría ingresar (autonumerico)
        strSQL = " SELECT COALESCE(max(lis_pre_codigo),0) as num " & _
                 " FROM lista_precio " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " GROUP BY emp_codigo"
        clsCon_Def.Ejecutar (strSQL)
        If clsCon_Def.adorec_Def.EOF Then
            txtCodigo.Text = 1
        Else
            txtCodigo.Text = Val(clsCon_Def.adorec_Def("num")) + 1
        End If
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

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    If txtCodigo.Text = "" Or txtDescripcion.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub


Private Sub txtPolitica_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtPolitica_Validate(Cancel As Boolean)
' Verifica si el dato uçingresado es numérico
    If IsNumeric(txtPolitica.Text) = False Then
        MsgBox "Solo se permiten valores numéricos", vbOKOnly + vbInformation, "ERROR"
        Cancel = True
    Else
        ' Pone dos decimales al valor
        txtPolitica.Text = FormatoD2(txtPolitica.Text)
        Cancel = False
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

Private Sub txtDescripcion_Change()
    If txtCodigo.Text = "" Or txtDescripcion.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Contenido
End Sub

Private Sub txtCodigo_Change()
    If txtCodigo.Text = "" Or txtDescripcion.Text = "" Then
        cmbAceptar.Enabled = False
    Else
        cmbAceptar.Enabled = True
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Contenido
End Sub
