VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de datos del sistema"
   ClientHeight    =   4815
   ClientLeft      =   5970
   ClientTop       =   4650
   ClientWidth     =   9480
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9480
   Begin VB.ComboBox cmbPuertos 
      Height          =   315
      ItemData        =   "frmConfig.frx":030A
      Left            =   6720
      List            =   "frmConfig.frx":032C
      TabIndex        =   42
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Impresoras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   33
      Top             =   240
      Width           =   4575
      Begin VB.TextBox txtImpresoraPorDefecto 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   1920
      End
      Begin VB.CommandButton cmdImpresotaPorDefecto 
         Caption         =   "..."
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdImpresotaEtiqueta 
         Caption         =   "..."
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdImpresoraTicket 
         Caption         =   "..."
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtImpresoraEtiqueta 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txtImpresoraTicket 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1920
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinterAUX 
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
         _cx             =   7646
         _cy             =   661
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   -2.58236865538736
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   3
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
         AutoLinkNavigate=   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin VB.Label Label12 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Impresora por Defecto"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Impresora de Etiqueta:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Impresora de Ticket:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   285
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos Conexión SQL "
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
      Left            =   4680
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txt_instanciasql 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txt_catalogo 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txt_passwordbd 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1320
         Width           =   1920
      End
      Begin VB.TextBox txt_userbd 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Nombre Instancia SQL"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Nombre Catálogo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Password"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   1365
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "User"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1005
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Datos del Sistema"
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
      TabIndex        =   18
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtServBDDLocal 
         Height          =   315
         Left            =   2190
         TabIndex        =   2
         Top             =   960
         Width           =   1920
      End
      Begin VB.TextBox txtPuertoBDDLocal 
         Height          =   315
         Left            =   2190
         TabIndex        =   3
         Top             =   1320
         Width           =   1920
      End
      Begin VB.TextBox txtAutorFactura 
         Height          =   315
         Left            =   2190
         TabIndex        =   7
         Top             =   3240
         Width           =   1920
      End
      Begin VB.TextBox txtPtoFactura 
         Height          =   315
         Left            =   2190
         TabIndex        =   6
         Top             =   2880
         Width           =   1920
      End
      Begin VB.TextBox txtNombreBDD 
         Height          =   315
         Left            =   2190
         TabIndex        =   4
         Top             =   1680
         Width           =   1920
      End
      Begin VB.TextBox txtServBDDMaster 
         Height          =   315
         Left            =   2190
         TabIndex        =   0
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtPuertoBDDMaster 
         Height          =   315
         Left            =   2190
         TabIndex        =   1
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox txtServWEB 
         Height          =   315
         Left            =   2190
         TabIndex        =   5
         Top             =   2040
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker dtpCaducaFactura 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2190
         TabIndex        =   8
         Top             =   3600
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/yyyy"
         Format          =   66256899
         CurrentDate     =   37463
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor de BDD Local"
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
         Left            =   165
         TabIndex        =   27
         Top             =   1005
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto de Conexión Local"
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
         Left            =   165
         TabIndex        =   26
         Top             =   1365
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caducidad de Facturas"
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
         TabIndex        =   25
         Top             =   3645
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autorización Facturas"
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
         TabIndex        =   24
         Top             =   3285
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Punto de Facturacion"
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
         TabIndex        =   23
         Top             =   2925
         Width           =   1530
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la BDD"
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
         Left            =   165
         TabIndex        =   22
         Top             =   1725
         Width           =   1305
      End
      Begin VB.Label lblCodio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor de BDD Master"
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
         Left            =   165
         TabIndex        =   21
         Top             =   285
         Width           =   1740
      End
      Begin VB.Label lbldireccion 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto de Conexión Master"
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
         Left            =   165
         TabIndex        =   20
         Top             =   645
         Width           =   1950
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor Web"
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
         Left            =   165
         TabIndex        =   19
         Top             =   2085
         Width           =   990
      End
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Puerto Balanza:"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   1965
      Width           =   1815
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para el ingreso y modificación de Bancos                              #
'#  frmBanco V1.0                                                               #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para el ingreso y modificación de Bancos.                           #
'#  Permitirá almacenar en la base de datos nuevos bancos y modificar sus       #
'#  nombres, dependiendo de la propiedad Tag, la cual se cambiará en la         #
'#  ventana frmSelBanco y desde esta se llamará a esta ventana.                 #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    Banco: En esta tabla se almacenan los nuevos bancos y se modifican        #
'#               los datos de estos.                                            #
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

Private Sub cmbAceptar_Click()
    Dim strPath As String
    Dim strLinea As String
    
    ImpresoraEtiqueta = txtImpresoraEtiqueta.Text
    ImpresoraTicket = txtImpresoraTicket.Text
    ImpresoraPorDefecto = txtImpresoraPorDefecto.Text
    PuertoBalanza = FormatoD0(Replace(cmbPuertos.Text, "COM", ""))
    
    GuardarImpresoras
    GuardarPuertoBalanza
    
    'LLenar variables para conexion con SQL Server
    s_instanciaSQL = UCase(txt_instanciasql.Text)
    s_catalogoSQL = UCase(txt_catalogo.Text)
    s_userSql = UCase(txt_userbd.Text)
    s_passwordSql = UCase(txt_passwordbd.Text)
    
    
    If Trim(txtServBDDMaster.Text) <> "" Or Trim(txtServBDDLocal.Text) <> "" Or Trim(txtNombreBDD.Text) <> "" Or Trim(txtServWEB.Text) <> "" Then
        strPath = Trim(App.Path)
        Open strPath & "\config" For Output As #1
            strLinea = "servidorbddmaster = " & Trim(txtServBDDMaster.Text)
            Print #1, strLinea
            If Trim(txtPuertoBDDMaster.Text) <> "" Then
                strLinea = "puertomaster = " & Trim(txtPuertoBDDMaster.Text)
                Print #1, strLinea
            End If
            strLinea = "servidorbddlocal = " & Trim(txtServBDDLocal.Text)
            Print #1, strLinea
            If Trim(txtPuertoBDDLocal.Text) <> "" Then
                strLinea = "puertolocal = " & Trim(txtPuertoBDDLocal.Text)
                Print #1, strLinea
            End If
            strLinea = "bdd = " & Trim(txtNombreBDD.Text)
            Print #1, strLinea
            strLinea = "servidorweb = " & Trim(txtServWEB.Text)
            Print #1, strLinea
            strLinea = "ptofactura = " & Trim(txtPtoFactura.Text)
            Print #1, strLinea
            strLinea = "autorfactura = " & Trim(txtAutorFactura.Text)
            Print #1, strLinea
            strLinea = "caducafactura = " & Format(dtpCaducaFactura.Value, "mm\/yyyy")
            Print #1, strLinea
            strLinea = "mssql = " & Trim(txt_instanciasql.Text)
            Print #1, strLinea
            strLinea = "mssqlcatalogo = " & Trim(txt_catalogo.Text)
            Print #1, strLinea
            strLinea = "mssqlusuario = " & Trim(txt_userbd.Text)
            Print #1, strLinea
            strLinea = "mssqlclave = " & Trim(txt_passwordbd.Text)
            Print #1, strLinea
        Close #1
        Unload Me
        frmConexion.Show
    Else
        MsgBox "Llene los campos de:" & vbNewLine & vbNewLine & _
               "· Servidor de Base de Datos" & vbNewLine & _
               "· Nombre de la Base de Datos" & vbNewLine & _
               "· Servidor Web", vbInformation, "Configuración"
    End If
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub cmdImpresoraTicket_Click()
    
    VSPrinterAUX.PrintDialog pdPrint
        
    txtImpresoraTicket.Text = VSPrinterAUX.Device
    
End Sub

Private Sub cmdImpresotaEtiqueta_Click()
    VSPrinterAUX.PrintDialog pdPrint
        
    txtImpresoraEtiqueta.Text = VSPrinterAUX.Device
    
End Sub

Private Sub cmdImpresotaPorDefecto_Click()
    VSPrinterAUX.PrintDialog pdPrint
        
    txtImpresoraPorDefecto.Text = VSPrinterAUX.Device

End Sub

Private Sub Form_Load()
    Dim strPath As String
    
    ' Para seleccionar el primer puerto encontrado:
    cmbPuertos.Text = "COM" & PuertoBalanza
    
    txtImpresoraEtiqueta.Text = ImpresoraEtiqueta
    txtImpresoraTicket.Text = ImpresoraTicket
    txtImpresoraPorDefecto.Text = ImpresoraPorDefecto
'    txt_catalogo = "Intermediate"
'    txt_instanciasql = "APLICACIONES-PC"
'    txt_passwordbd = "sa5%"
'    txt_userbd = "sa"
    strPath = Trim(App.Path)
    If Right(strpad, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    If Dir(strPath & "config") <> "" Then
        Open strPath & "config" For Input As #1
            For i = 0 To 12
                Input #1, strLinea
                strVar = LCase(Trim(Left(strLinea, InStr(strLinea, "=") - 1)))
                strVal = Trim(Right(strLinea, Len(strLinea) - InStr(strLinea, "=")))
                If Len(strVal) <> 0 Then
                    Select Case strVar
                        Case "servidorbddmaster"
                            txtServBDDMaster.Text = strVal
                        Case "puertomaster"
                            txtPuertoBDDMaster.Text = strVal
                        Case "servidorbddlocal"
                            txtServBDDLocal.Text = strVal
                        Case "puertolocal"
                            txtPuertoBDDLocal.Text = strVal
                        Case "bdd"
                            txtNombreBDD.Text = strVal
                        Case "servidorweb"
                            txtServWEB.Text = strVal
                        Case "ptofactura"
                            txtPtoFactura.Text = strVal
                        Case "autorfactura"
                            txtAutorFactura.Text = strVal
                        Case "caducafactura"
                            dtpCaducaFactura.Value = strVal
                        Case "mssql"
                            txt_instanciasql.Text = strVal
                        Case "mssqlcatalogo"
                            txt_catalogo.Text = strVal
                        Case "mssqlusuario"
                            txt_userbd.Text = strVal
                        Case "mssqlclave"
                            txt_passwordbd.Text = strVal
                    End Select
                End If
                If EOF(1) Then
                    Exit For
                End If
            Next
        Close #1
    End If
End Sub

Private Sub txtServBDDMaster_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtServBDDLocal_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtNombreBDD_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtPuertoBDDMaster_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtPuertoBDDLocal_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtServWEB_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtPtoFactura_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub txtAutorFactura_GotFocus()
    Seleccionar_Contenido
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub
