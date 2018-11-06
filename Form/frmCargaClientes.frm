VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCargaClientes 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Clientes"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargaClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   13185
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5025
      TabIndex        =   25
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Migracion"
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
      TabIndex        =   28
      Top             =   0
      Width           =   12960
      Begin VB.OptionButton optN9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   3158
         Width           =   255
      End
      Begin VB.OptionButton optN8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   2798
         Width           =   255
      End
      Begin VB.OptionButton optN7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   2438
         Width           =   255
      End
      Begin VB.OptionButton optN6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   2078
         Width           =   255
      End
      Begin VB.OptionButton optN5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1718
         Width           =   255
      End
      Begin VB.OptionButton optN4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1358
         Width           =   255
      End
      Begin VB.OptionButton optN3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   998
         Width           =   255
      End
      Begin VB.OptionButton optN2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   638
         Width           =   255
      End
      Begin VB.OptionButton optN1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   0
         Top             =   278
         Width           =   255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3015
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   12495
         _cx             =   89740824
         _cy             =   89724102
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
         FormatString    =   $"frmCargaClientes.frx":030A
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
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   9480
         TabIndex        =   27
         Top             =   240
         Width           =   2640
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   315
         Left            =   12120
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin MSDataListLib.DataCombo cmbCategoria 
         Height          =   330
         Left            =   9600
         TabIndex        =   19
         Top             =   960
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSComDlg.CommonDialog cdArchivo 
         Left            =   11640
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Archivo de Backup"
         InitDir         =   "C:\"
      End
      Begin MSDataListLib.DataCombo cmbCanal 
         Height          =   330
         Left            =   9600
         TabIndex        =   20
         Top             =   1320
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNegocio 
         Height          =   330
         Left            =   9600
         TabIndex        =   21
         Top             =   1680
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbFEntrega 
         Height          =   330
         Left            =   9600
         TabIndex        =   23
         Top             =   2400
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDirector 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEmprendedor 
         Height          =   330
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEjecutivo 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbVendedor 
         Height          =   330
         Left            =   9600
         TabIndex        =   22
         Top             =   2040
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCiudad 
         Height          =   330
         Left            =   9600
         TabIndex        =   24
         Top             =   2760
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   330
         Left            =   1560
         TabIndex        =   9
         Top             =   1680
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN6 
         Height          =   330
         Left            =   1560
         TabIndex        =   11
         Top             =   2040
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN7 
         Height          =   330
         Left            =   1560
         TabIndex        =   13
         Top             =   2400
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN8 
         Height          =   330
         Left            =   1560
         TabIndex        =   15
         Top             =   2760
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbN9 
         Height          =   330
         Left            =   1560
         TabIndex        =   17
         Top             =   3120
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   735
         TabIndex        =   46
         Top             =   3180
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   735
         TabIndex        =   45
         Top             =   2820
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   735
         TabIndex        =   44
         Top             =   2460
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   735
         TabIndex        =   43
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   735
         TabIndex        =   42
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.Zona N1:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   41
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dir N2:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   495
         TabIndex        =   40
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empren N3:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   39
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eje.E. N4:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   300
         TabIndex        =   38
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   9000
         TabIndex        =   37
         Top             =   2820
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8730
         TabIndex        =   36
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Entrega:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8760
         TabIndex        =   34
         Top             =   2460
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negocio:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8865
         TabIndex        =   33
         Top             =   1740
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Canal:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   9045
         TabIndex        =   32
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8760
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8760
         TabIndex        =   29
         Top             =   1020
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6705
      TabIndex        =   26
      Top             =   7080
      Width           =   1455
   End
End
Attribute VB_Name = "frmCargaClientes"
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

Private Sub cmdAplicar_Click()
    Dim i As Long
    Me.MousePointer = 11
    
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    
    For i = 1 To VSFG.Rows - 1
        If Val(VSFG.TextMatrix(i, VSFG.Cols - 1)) = 1 Then
            
            strSql = " SELECT CONCAT('C',LPAD(ROUND(COALESCE(MAX(RIGHT(per_codigo,LEN(per_codigo)-1)),0)+1,0),5,'0')) as cod " & _
                     " FROM persona " & _
                     " WHERE cat_p_tipo='C'" & _
                     " AND emp_codigo='" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
            clsCon_Def.Ejecutar strSql
            VSFG.TextMatrix(i, 3) = clsCon_Def.adorec_Def("cod")
'            strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
'                     " cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
'                     " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
'                     " per_credito,per_dcto,for_pag_codigo,ven_codigo,tip_ped_codigo,per_codigo_ref,per_codigo_ref2," & _
'                     " per_observacion,per_fac_flete,per_sec_publico,per_siniva, " & _
'                     " per_fechamod,per_usumod,per_fechaing,per_usuing,per_perdesde,per_inactivo,per_aplica_nc) " & _
'                     " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 4) & "','" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "', " & _
'                     " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & VSFG.TextMatrix(i, 17) & "','" & VSFG.TextMatrix(i, 18) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & VSFG.TextMatrix(i, 20) & "','" & FormatoD2(VSFG.TextMatrix(i, 21)) & "','" & FormatoD4(VSFG.TextMatrix(i, 22)) & "','" & VSFG.TextMatrix(i, 23) & "','" & _
'                     VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "','" & VSFG.TextMatrix(i, 26) & "','" & VSFG.TextMatrix(i, 27) & "','" & UCase(VSFG.TextMatrix(i, 28)) & "'," & _
'                     " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 30))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "'," & _
'                     " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_DATE,'" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "')"
'            clsCon_Def.Ejecutar strSql, "M"
        Else
            Exit For
        End If
    Next i
    
'    For i = 1 To VSFG.Rows - 1
'        If Val(VSFG.TextMatrix(i, VSFG.Cols - 1)) = 1 Then
'            strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido,per_nombre," & _
'                     " cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
'                     " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea,per_direccion2,for_ent_codigo," & _
'                     " per_credito,per_dcto,for_pag_codigo,ven_codigo,tip_ped_codigo,per_codigo_ref,per_codigo_ref2," & _
'                     " per_observacion,per_fac_flete,per_sec_publico,per_siniva, " & _
'                     " per_fechamod,per_usumod,per_fechaing,per_usuing,per_perdesde,per_inactivo,per_aplica_nc) " & _
'                     " VALUES ('" & strEmpresa & "','" & VSFG.TextMatrix(i, 1) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 2))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 3))) & "','C','" & VSFG.TextMatrix(i, 4) & "','" & UCase(VSFG.TextMatrix(i, 5)) & "','" & UCase(VSFG.TextMatrix(i, 6)) & "', " & _
'                     " '" & UCase(VSFG.TextMatrix(i, 7)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & UCase(Trim(VSFG.TextMatrix(i, 9))) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 11)) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 14)) & "','" & UCase(VSFG.TextMatrix(i, 15)) & "','" & UCase(VSFG.TextMatrix(i, 16)) & "','" & VSFG.TextMatrix(i, 17) & "','" & VSFG.TextMatrix(i, 18) & "'," & _
'                     " '" & UCase(VSFG.TextMatrix(i, 19)) & "','" & VSFG.TextMatrix(i, 20) & "','" & FormatoD2(VSFG.TextMatrix(i, 21)) & "','" & FormatoD4(VSFG.TextMatrix(i, 22)) & "','" & VSFG.TextMatrix(i, 23) & "','" & _
'                     VSFG.TextMatrix(i, 24) & "','" & VSFG.TextMatrix(i, 25) & "','" & VSFG.TextMatrix(i, 26) & "','" & VSFG.TextMatrix(i, 27) & "','" & UCase(VSFG.TextMatrix(i, 28)) & "'," & _
'                     " '" & Abs(FormatoD0(VSFG.TextMatrix(i, 30))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 33))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 34))) & "'," & _
'                     " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_DATE,'" & Abs(FormatoD0(VSFG.TextMatrix(i, 35))) & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 36))) & "')"
'            clsCon_Def.Ejecutar strSql, "M"
'        Else
'            Exit For
'        End If
'    Next i
    Me.MousePointer = 0
    MsgBox "Carga de clintes", vbInformation, "Clientes"
    Unload Me
End Sub

Private Sub cmbN9_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN9.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref8")
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN8_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN8.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN7_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN7.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN6_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN6.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbN5_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN5.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEjecutivo_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEjecutivo.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbEmprendedor_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEmprendedor.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbDirector_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3,COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6,COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbDirector.BoundText & "'" & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
    End If
End Sub

Private Sub cmbGerente_Validate(Cancel As Boolean)
        cmbN9.BoundText = ""
        cmbN8.BoundText = ""
        cmbN7.BoundText = ""
        cmbN6.BoundText = ""
        cmbN5.BoundText = ""
        cmbEjecutivo.BoundText = ""
        cmbEmprendedor.BoundText = ""
        cmbDirector.BoundText = ""
End Sub

Private Sub cmdAceptar_Click()

    Dim i As Long
    Dim j As Long
    j = 0
    VSFG.Select 1, VSFG.Cols - 1
    VSFG.Sort = flexSortGenericDescending
    For i = 1 To VSFG.Rows - 1
        If VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
            If VSFG.TextMatrix(i, 1) <> "" Then
            j = j + 1
            strSql = " SELECT CONCAT('C',LPAD(ROUND(COALESCE(MAX(REPLACE(per_codigo,'C','0')+0),0)+1,0),6,'0')) as cod " & _
                     " FROM persona " & _
                     " WHERE cat_p_tipo='C'" & _
                     " AND emp_codigo='" & strEmpresa & "'" & _
                     " GROUP BY emp_codigo"
            clsCon_Def.Ejecutar strSql
            
            'controla que no se repita el código
            
            
            strSql = " INSERT INTO persona(emp_codigo,per_codigo,per_cm,per_rcm,cat_p_tipo,per_tipo,per_apellido," & _
                    " per_nombre,cat_p_codigo,can_codigo,per_ruc,ciu_codigo,zon_codigo," & _
                    " per_direccion,per_ubicacion,per_telf,per_fax,per_celular,per_email,per_fechacumplea," & _
                    " per_direccion2,for_ent_codigo,per_credito,per_dcto,for_pag_codigo,for_pag_codigo_imp," & _
                    " ven_codigo,tip_ped_codigo," & _
                    " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4," & _
                    " per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                    " per_observacion," & _
                    " per_fac_flete,per_especial,per_bloqueado,per_sec_publico,per_siniva, " & _
                    " per_inactivo,per_perdesde,per_es_gz,per_es_di,per_es_em,per_es_ee," & _
                    " sac_codigo,cob_codigo,per_aplica_nc," & _
                    " per_fechamod,per_usumod,per_fechaing,per_usuing) " & _
                    " VALUES ('" & strEmpresa & "','" & clsCon_Def.adorec_Def("cod") & "','1','1','C','" & IIf(VSFG.TextMatrix(i, 0) = "C", "Natural", "Jurídico") & "','" & UCase(VSFG.TextMatrix(i, 3)) & "'," & _
                    " '" & UCase(VSFG.TextMatrix(i, 2)) & "','" & cmbCategoria.BoundText & "','" & cmbCanal.BoundText & "','" & UCase(VSFG.TextMatrix(i, 1)) & "','" & cmbCiudad.BoundText & "','NA'," & _
                    " '" & UCase(VSFG.TextMatrix(i, 13)) & "','" & UCase(VSFG.TextMatrix(i, 12)) & "','" & UCase(VSFG.TextMatrix(i, 9)) & "','" & UCase(VSFG.TextMatrix(i, 10)) & "','" & UCase(VSFG.TextMatrix(i, 8)) & "','" & VSFG.TextMatrix(i, 7) & "','" & Right(VSFG.TextMatrix(i, 6), 4) & "-" & Right(Left(VSFG.TextMatrix(i, 6), 4), 2) & "-" & Left(VSFG.TextMatrix(i, 6), 2) & "'," & _
                    " '" & UCase(VSFG.TextMatrix(i, 16)) & "','" & cmbFEntrega.BoundText & "','0','0','CONT','CONT'," & _
                    " '" & cmbVendedor.BoundText & "','" & cmbNegocio.BoundText & "'," & _
                    " '" & cmbGerente.BoundText & "','" & cmbDirector.BoundText & "','" & cmbEmprendedor.BoundText & "','" & cmbEjecutivo.BoundText & "'," & _
                    " '" & cmbN5.BoundText & "','" & cmbN6.BoundText & "','" & cmbN7.BoundText & "','" & cmbN8.BoundText & "','" & cmbN9.BoundText & "'," & _
                    " '" & UCase(VSFG.TextMatrix(i, 17)) & "'," & _
                    " '0','0','0','0','0'," & _
                    " '0','" & Right(VSFG.TextMatrix(i, 23), 4) & "-" & Right(Left(VSFG.TextMatrix(i, 23), 4), 2) & "-" & Left(VSFG.TextMatrix(i, 23), 2) & "','0','0','0','0'," & _
                    " '','','0'," & _
                    " CURRENT_TIMESTAMP, '" & strUsuario & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
                    
            clsCon_Def.Ejecutar strSql, "M"
            End If
        Else
        Exit For
        End If
    Next i
    
    Me.cmdAceptar.Enabled = False
    MsgBox "Termino la migracion " & Chr(13) & _
    "Registros Amarillos Migrados  : " & CStr(j) & Chr(13) & _
    "Registros Blancos No Migrados : " & CStr(VSFG.Rows - j - 2) & Chr(13) _
    , vbInformation, "Migracion de Clientes"
End Sub

Private Sub cmdExplorar_Click()
    Dim sDir As String
    Dim i As Long
    If cmbNegocio.MatchedWithList = False Then
        MsgBox "SELECCIONE UN NEGOCIO"
        Exit Sub
    End If
    sDir = CurDir
    txtArchivo.Tag = sDir
    cdArchivo.ShowOpen
    txtArchivo = cdArchivo.FileName
    ChDir sDir
    If (txtArchivo.Text <> "") Then
        Me.cmdAceptar.Enabled = True
        Me.MousePointer = 11
        VSFG.ClipSeparators = ";" & vbCr
        VSFG.FixedRows = 0
        VSFG.Rows = 0
        VSFG.LoadGrid txtArchivo.Text, flexFileTabText
        'Call LeerExcel(txtArchivo.Text)
        VSFG.FixedRows = 1
        VSFG.Cols = VSFG.Cols + 1
        
        For i = 1 To VSFG.Rows - 1
            strSql = " SELECT COALESCE(count(*),0) as n " & _
                     " FROM persona " & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND per_ruc='" & VSFG.TextMatrix(i, 1) & "' AND tip_ped_codigo='" & cmbNegocio.BoundText & "' "

            clsCon_Def.Ejecutar strSql
            VSFG.ShowCell i, 1
            If clsCon_Def.adorec_Def("n") > 0 Then
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 0
                VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbWhite
            Else
                VSFG.TextMatrix(i, VSFG.Cols - 1) = 2
                VSFG.Cell(flexcpBackColor, i, 0, i, VSFG.Cols - 1) = vbYellow
            End If
        Next i
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
'Private Sub LocalizaCmb(combo As DataCombo, Codigo As String)
'Dim i As Long
'combo.Enabled = True
'For i = 0 To combo.Container - 1
'    combo.se = i
'    If combo.DataMember = Codigo Then
'        combo.Enabled = False
'    End If
'Next i
'
'End Sub

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set ucrtVSFG.VSFGControl = VSFG
    ucrtVSFG.Inicializar False, False, False
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        strSql = " SELECT cat_p_codigo ,cat_p_nombre " & _
                 " FROM categoria_p " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' and cat_p_codigo = 'EJE' " & _
                 " ORDER BY cat_p_codigo "
        clsCon_Def.Ejecutar strSql
        Set cmbCategoria.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbCategoria.ListField = "cat_p_nombre"
        cmbCategoria.BoundColumn = "cat_p_codigo"
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            cmbCategoria.BoundText = clsCon_Def.adorec_Def(0)
        End If
        cmbCategoria.Enabled = False
        
        strSql = " SELECT can_codigo ,can_nombre " & _
                 " FROM canal " & _
                 " WHERE emp_codigo='" & strEmpresa & "' and can_codigo = 'VPC' " & _
                 " ORDER BY can_codigo "
        clsCon_Def.Ejecutar strSql
        Set cmbCanal.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbCanal.ListField = "can_nombre"
        cmbCanal.BoundColumn = "can_codigo"
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            cmbCanal.BoundText = clsCon_Def.adorec_Def(0)
        End If
        cmbCanal.Enabled = False
        
        Set cmbNegocio.RowSource = ComboNegocioDataSource.DataSource
        cmbNegocio.ListField = "tip_ped_nombre"
        cmbNegocio.BoundColumn = "tip_ped_codigo"
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            cmbNegocio.BoundText = clsCon_Def.adorec_Def(0)
        End If
        cmbNegocio.Enabled = False
        
        strSql = " SELECT for_ent_codigo ,for_ent_nombre " & _
                 " FROM forma_entrega " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY for_ent_codigo "
        clsCon_Def.Ejecutar strSql
        Set cmbFEntrega.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbFEntrega.ListField = "for_ent_nombre"
        cmbFEntrega.BoundColumn = "for_ent_codigo"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_gz=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbGerente.RowSource = clsCon_Def.adorec_Def
        cmbGerente.BoundColumn = "codigo"
        cmbGerente.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_di=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbDirector.RowSource = clsCon_Def.adorec_Def
        cmbDirector.BoundColumn = "codigo"
        cmbDirector.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_em=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbEmprendedor.RowSource = clsCon_Def.adorec_Def
        cmbEmprendedor.BoundColumn = "codigo"
        cmbEmprendedor.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_ee=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbEjecutivo.RowSource = clsCon_Def.adorec_Def
        cmbEjecutivo.BoundColumn = "codigo"
        cmbEjecutivo.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n5=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN5.RowSource = clsCon_Def.adorec_Def
        cmbN5.BoundColumn = "codigo"
        cmbN5.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n6=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN6.RowSource = clsCon_Def.adorec_Def
        cmbN6.BoundColumn = "codigo"
        cmbN6.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n7=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN7.RowSource = clsCon_Def.adorec_Def
        cmbN7.BoundColumn = "codigo"
        cmbN7.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n8=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN8.RowSource = clsCon_Def.adorec_Def
        cmbN8.BoundColumn = "codigo"
        cmbN8.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n9=1 AND p1.tip_ped_codigo='JON'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN9.RowSource = clsCon_Def.adorec_Def
        cmbN9.BoundColumn = "codigo"
        cmbN9.ListField = "nombre"
        
        strSql = " SELECT DISTINCT ven_codigo as codigo, CONCAT(ven_apellido,' ',ven_nombre) AS nombre " & _
                 " FROM vendedor " & _
                " WHERE emp_codigo='" & strEmpresa & "' and ven_codigo ='JSE'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbVendedor.RowSource = clsCon_Def.adorec_Def
        cmbVendedor.BoundColumn = "codigo"
        cmbVendedor.ListField = "nombre"
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            cmbVendedor.BoundText = clsCon_Def.adorec_Def(0)
        End If
        cmbVendedor.Enabled = False
        
        strSql = " SELECT DISTINCT ciu_codigo as codigo, ciu_nombre AS nombre " & _
                 " FROM ciudad " & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbCiudad.RowSource = clsCon_Def.adorec_Def
        cmbCiudad.BoundColumn = "codigo"
        cmbCiudad.ListField = "nombre"
        
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

Function LeerExcel(Archivo As String) As Boolean
On Error GoTo SalirExcel
'dimensiones
LeerExcel = False
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim lngUltimaFila As Long, Fil As Long, Col As Long, FilXl As Long, mc As Boolean

Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

'abrir programa Excel
Set xlApp = New Excel.Application
xlApp.Visible = False

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(Archivo, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)

'2. Si no conoces el rango
lngUltimaFila = Columns("A:A").Range("A65536").End(xlUp).Row

'lngUltimaFila = 17

If MsgBox("Serán Migrados #" & CStr(lngUltimaFila - 1) & " Desea Continuar?", vbYesNo) = vbNo Then Exit Function
VSFG.Rows = 1
VSFG.Cols = 54
For FilXl = 1 To lngUltimaFila
    VSFG.Rows = Fil + 1
    mc = False
    For Col = 1 To VSFG.Cols - 1
        If Col = 1 And Len(xlHoja.Range(xlHoja.Cells(FilXl, 2), xlHoja.Cells(FilXl, 2))) > 0 Then
            mc = True
            Fil = Fil + 1
        End If
        If mc = True Then
            VSFG.TextMatrix(Fil - 1, Col - 1) = xlHoja.Range(xlHoja.Cells(FilXl, Col), xlHoja.Cells(FilXl, Col))
        End If
    Next Col
Next FilXl
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit


'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing
LeerExcel = True
Exit Function
SalirExcel:
    LeerExcel = False
    MsgBox "El Formato del Archivo Excel no es el Correcto."
End Function

Private Sub optN1_Click()
    ActivarCombo
End Sub

Private Sub optN2_Click()
    ActivarCombo
End Sub

Private Sub optN3_Click()
    ActivarCombo
End Sub

Private Sub optN4_Click()
    ActivarCombo
End Sub

Private Sub optN5_Click()
    ActivarCombo
End Sub

Private Sub optN6_Click()
    ActivarCombo
End Sub

Private Sub optN7_Click()
    ActivarCombo
End Sub

Private Sub optN8_Click()
    ActivarCombo
End Sub

Private Sub optN9_Click()
    ActivarCombo
End Sub

Private Sub ActivarCombo()
    cmbGerente.Locked = True
    cmbDirector.Locked = True
    cmbEmprendedor.Locked = True
    cmbEjecutivo.Locked = True
    cmbN5.Locked = True
    cmbN6.Locked = True
    cmbN7.Locked = True
    cmbN8.Locked = True
    cmbN9.Locked = True
    If optN1.Value = True Then
        cmbGerente.Locked = False
    ElseIf optN2.Value = True Then
        cmbDirector.Locked = False
    ElseIf optN3.Value = True Then
        cmbEmprendedor.Locked = False
    ElseIf optN4.Value = True Then
        cmbEjecutivo.Locked = False
    ElseIf optN5.Value = True Then
        cmbN5.Locked = False
    ElseIf optN6.Value = True Then
        cmbN6.Locked = False
    ElseIf optN7.Value = True Then
        cmbN7.Locked = False
    ElseIf optN8.Value = True Then
        cmbN8.Locked = False
    ElseIf optN9.Value = True Then
        cmbN9.Locked = False
    End If
End Sub
