VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRed 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Red"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4208
      TabIndex        =   24
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.OptionButton optN10 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   705
         TabIndex        =   21
         Top             =   3525
         Width           =   255
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtCI 
         Height          =   285
         Left            =   2265
         TabIndex        =   1
         Top             =   3960
         Width           =   2415
      End
      Begin VB.OptionButton optN1 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   278
         Width           =   255
      End
      Begin VB.OptionButton optN2 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   638
         Width           =   255
      End
      Begin VB.OptionButton optN3 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   998
         Width           =   255
      End
      Begin VB.OptionButton optN4 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   1358
         Width           =   255
      End
      Begin VB.OptionButton optN5 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1718
         Width           =   255
      End
      Begin VB.OptionButton optN6 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   2078
         Width           =   255
      End
      Begin VB.OptionButton optN7 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   2438
         Width           =   255
      End
      Begin VB.OptionButton optN8 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   2798
         Width           =   255
      End
      Begin VB.OptionButton optN9 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Option1"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   3158
         Width           =   255
      End
      Begin MSDataListLib.DataCombo cmbGerente 
         Height          =   330
         Left            =   1080
         TabIndex        =   4
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
         Left            =   1080
         TabIndex        =   6
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
         Left            =   1080
         TabIndex        =   8
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
         Left            =   1080
         TabIndex        =   10
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
      Begin MSDataListLib.DataCombo cmbN5 
         Height          =   330
         Left            =   1080
         TabIndex        =   12
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
         Left            =   1080
         TabIndex        =   14
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
         Left            =   1080
         TabIndex        =   16
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
         Left            =   1080
         TabIndex        =   18
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
         Left            =   1080
         TabIndex        =   20
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
      Begin MSDataListLib.DataCombo cmbN10 
         Height          =   330
         Left            =   1065
         TabIndex        =   22
         Top             =   3480
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N10:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   3540
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Lider:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1080
         TabIndex        =   34
         Top             =   3990
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N4:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   33
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N3:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   32
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N2:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   31
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N1:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   30
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N5:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   29
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N6:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   28
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N7:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   27
         Top             =   2460
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N8:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   26
         Top             =   2820
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N9:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   255
         TabIndex        =   25
         Top             =   3180
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmRed"
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
Public Negocio  As String
Public CodPer  As String
Public Linea As Long


Private Sub cmbN10_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6," & _
             " COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8,COALESCE(per_codigo_ref9,'') as per_codigo_ref9 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN10.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref9")
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

Private Sub cmbN9_Validate(Cancel As Boolean)
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6," & _
             " COALESCE(per_codigo_ref7,'') as per_codigo_ref7,COALESCE(per_codigo_ref8,'') as per_codigo_ref8 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN9.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6," & _
             " COALESCE(per_codigo_ref7,'') as per_codigo_ref7 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN8.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5,COALESCE(per_codigo_ref6,'') as per_codigo_ref6 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN7.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4,COALESCE(per_codigo_ref5,'') as per_codigo_ref5 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN6.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3," & _
             " COALESCE(per_codigo_ref4,'') as per_codigo_ref4 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbN5.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2,COALESCE(per_codigo_ref3,'') as per_codigo_ref3 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEjecutivo.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref,COALESCE(per_codigo_ref2,'') as per_codigo_ref2 " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbEmprendedor.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
    strSql = " SELECT COALESCE(per_codigo_ref,'') as per_codigo_ref " & _
             " FROM persona " & _
             " WHERE cat_p_tipo='C'" & _
             " AND emp_codigo='" & strEmpresa & "'" & _
             " AND per_codigo='" & cmbDirector.BoundText & "'" & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = ""
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
        cmbN10.BoundText = ""
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
    Dim Cli As String
    Dim esContado As Boolean
    esContado = False
    strSql = " SELECT per_red_contado " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " AND per_codigo='" & cmbGerente.BoundText & "' " & _
             " AND tip_ped_codigo='" & Negocio & "'"
    
    clsCon_Def.Ejecutar strSql
    If FormatoD0(clsCon_Def.adorec_Def("per_red_contado")) = 1 Then esContado = True
    
    strSql = " SELECT TOP 1 COALESCE(ven_codigo,'') as ven_codigo,tip_ped_codigo,COALESCE(sac_codigo,'') as sac_codigo,COALESCE(cob_codigo,'') as cob_codigo," & _
             " COALESCE(cat_p_codigo,'') as cat_p_codigo,COALESCE(can_codigo,'') as can_codigo,for_pag_codigo,for_pag_codigo_imp,COALESCE(per_codigo_resp,'') as per_codigo_resp," & _
             " COALESCE(per_codigo_postal,'') as per_codigo_postal,count(*) as moda " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' AND tip_ped_codigo='" & Negocio & "'" & _
             " AND per_codigo_ref='" & cmbGerente.BoundText & "' " & _
             " AND per_codigo_ref2='" & cmbDirector.BoundText & "' " & _
             " AND per_codigo_ref3='" & cmbEmprendedor.BoundText & "' " & _
             " AND per_codigo_ref4='" & cmbEjecutivo.BoundText & "' " & _
             " AND per_codigo_ref5='" & cmbN5.BoundText & "' " & _
             " AND per_codigo_ref6='" & cmbN6.BoundText & "' " & _
             " AND per_codigo_ref7='" & cmbN7.BoundText & "' " & _
             " AND per_codigo_ref8='" & cmbN8.BoundText & "' " & _
             " AND per_codigo_ref9='" & cmbN9.BoundText & "' " & _
             " GROUP BY ven_codigo,tip_ped_codigo,sac_codigo,cob_codigo,cat_p_codigo,can_codigo,for_pag_codigo,for_pag_codigo_imp,per_codigo_resp,per_codigo_postal " & _
             " ORDER BY moda DESC"
    
    clsCon_Def.Ejecutar strSql
    
    If clsCon_Def.adorec_Def.RecordCount <= 0 Then
        Cli = ""
        If optN10.Value = True Then
            If cmbN9.BoundText <> "" Then
                Cli = cmbN9.BoundText
            ElseIf cmbN8.BoundText <> "" Then
                Cli = cmbN8.BoundText
            ElseIf cmbN7.BoundText <> "" Then
                Cli = cmbN7.BoundText
            ElseIf cmbN6.BoundText <> "" Then
                Cli = cmbN6.BoundText
            ElseIf cmbN5.BoundText <> "" Then
                Cli = cmbN5.BoundText
            ElseIf cmbN4.BoundText <> "" Then
                Cli = cmbN4.BoundText
            ElseIf cmbN3.BoundText <> "" Then
                Cli = cmbN3.BoundText
            ElseIf cmbN2.BoundText <> "" Then
                Cli = cmbN2.BoundText
            ElseIf cmbN1.BoundText <> "" Then
                Cli = cmbN1.BoundText
            End If
        End If
        If optN9.Value = True Then
            Cli = cmbN9.BoundText
        ElseIf optN8.Value = True Then
            Cli = cmbN8.BoundText
        ElseIf optN7.Value = True Then
            Cli = cmbN7.BoundText
        ElseIf optN6.Value = True Then
            Cli = cmbN6.BoundText
        ElseIf optN5.Value = True Then
            Cli = cmbN5.BoundText
        ElseIf optN4.Value = True Then
            Cli = cmbEjecutivo.BoundText
        ElseIf optN3.Value = True Then
            Cli = cmbEmprendedor.BoundText
        ElseIf optN2.Value = True Then
            Cli = cmbDirector.BoundText
        ElseIf optN1.Value = True Then
            Cli = cmbGerente.BoundText
        End If
        
        strSql = " SELECT TOP 1 COALESCE(ven_codigo,'') as ven_codigo,tip_ped_codigo,COALESCE(sac_codigo,'') as sac_codigo,COALESCE(cob_codigo,'') as cob_codigo," & _
                 " COALESCE(cat_p_codigo,'') as cat_p_codigo,COALESCE(can_codigo,'') as can_codigo,for_pag_codigo,for_pag_codigo_imp,COALESCE(per_codigo_resp,'') as per_codigo_resp," & _
                 " COALESCE(per_codigo_postal,'') as per_codigo_postal,count(*) as moda " & _
                 " FROM persona " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND cat_p_tipo='C' AND tip_ped_codigo='" & Negocio & "'" & _
                 " AND per_codigo='" & Cli & "' " & _
                 " GROUP BY ven_codigo,tip_ped_codigo,sac_codigo,cob_codigo,cat_p_codigo,can_codigo,for_pag_codigo,for_pag_codigo_imp,per_codigo_resp,ven_codigo,tip_ped_codigo," & _
                 " sac_codigo,cob_codigo,cat_p_codigo,can_codigo,for_pag_codigo,for_pag_codigo_imp,per_codigo_resp,per_codigo_postal " & _
                 " ORDER BY moda DESC"
        
        clsCon_Def.Ejecutar strSql
    
    End If
    
    frmClienteMod.VSFG.TextMatrix(Linea, 40) = cmbGerente.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 41) = cmbDirector.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 42) = cmbEmprendedor.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 43) = cmbEjecutivo.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 44) = cmbN5.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 45) = cmbN6.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 46) = cmbN7.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 47) = cmbN8.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 48) = cmbN9.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, 49) = cmbN10.BoundText
    frmClienteMod.VSFG.TextMatrix(Linea, frmClienteMod.VSFG.Cols - 1) = Abs(frmClienteMod.VSFG.TextMatrix(Linea, frmClienteMod.VSFG.Cols - 1))
    
    frmClienteMod.VSFG.TextMatrix(Linea, 28) = 0
    frmClienteMod.VSFG.TextMatrix(Linea, 29) = 0
    
    
    frmClienteMod.VSFG.TextMatrix(Linea, 17) = clsCon_Def.adorec_Def("per_codigo_postal")
    If esContado = False Then
        'forma de pago de la mayoria de hermanos
        frmClienteMod.VSFG.TextMatrix(Linea, 30) = clsCon_Def.adorec_Def("for_pag_codigo")
        frmClienteMod.VSFG.TextMatrix(Linea, 31) = clsCon_Def.adorec_Def("for_pag_codigo_imp")
    Else
        frmClienteMod.VSFG.TextMatrix(Linea, 30) = "CONT"
        frmClienteMod.VSFG.TextMatrix(Linea, 31) = "CONT"
    End If
    'deudor
    frmClienteMod.VSFG.TextMatrix(Linea, 37) = clsCon_Def.adorec_Def("per_codigo_resp")
    'vendedor
    frmClienteMod.VSFG.TextMatrix(Linea, 38) = clsCon_Def.adorec_Def("ven_codigo")
    'tipo de negocio
    frmClienteMod.VSFG.TextMatrix(Linea, 39) = clsCon_Def.adorec_Def("tip_ped_codigo")
    'sac
    frmClienteMod.VSFG.TextMatrix(Linea, 70) = clsCon_Def.adorec_Def("sac_codigo")
    'cobrador
    frmClienteMod.VSFG.TextMatrix(Linea, 71) = clsCon_Def.adorec_Def("cob_codigo")
    'categoria 11
    frmClienteMod.VSFG.TextMatrix(Linea, 11) = clsCon_Def.adorec_Def("cat_p_codigo")
    'canal 12
    frmClienteMod.VSFG.TextMatrix(Linea, 13) = clsCon_Def.adorec_Def("can_codigo")
    
    strSql = " SELECT per_direccion2,for_ent_codigo " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " AND per_codigo='" & IIf(cmbN9.BoundText <> "", cmbN9.BoundText, IIf(cmbN8.BoundText <> "", cmbN8.BoundText, IIf(cmbN7.BoundText <> "", cmbN7.BoundText, IIf(cmbN6.BoundText <> "", cmbN6.BoundText, IIf(cmbN5.BoundText <> "", cmbN5.BoundText, IIf(cmbEjecutivo.BoundText <> "", cmbEjecutivo.BoundText, IIf(cmbEmprendedor.BoundText <> "", cmbEmprendedor.BoundText, IIf(cmbDirector.BoundText <> "", cmbDirector.BoundText, cmbGerente.BoundText)))))))) & "' " & _
             " AND tip_ped_codigo='" & Negocio & "'"
    clsCon_Def.Ejecutar strSql
    'direccion de entrega papa
    frmClienteMod.VSFG.TextMatrix(Linea, 26) = clsCon_Def.adorec_Def("per_direccion2")
    'forma de entrega papa
    frmClienteMod.VSFG.TextMatrix(Linea, 27) = clsCon_Def.adorec_Def("for_ent_codigo")
    
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    strSql = " SELECT per_codigo,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9,per_codigo_ref10 " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cat_p_tipo='C' " & _
             " AND per_ruc='" & txtCI.Text & "' AND tip_ped_codigo='" & Negocio & "'" & _
             " AND (per_es_gz=1 " & _
             " OR per_es_di=1 OR per_es_em=1 OR per_es_ee=1 OR per_es_n5=1 " & _
             " OR per_es_n6=1 OR per_es_n7=1 OR per_es_n8=1 OR per_es_n9=1 " & _
             " OR per_es_n10=1) " & _
             " AND (per_codigo=per_codigo_ref OR per_codigo=per_codigo_ref2 OR per_codigo=per_codigo_ref3 " & _
             " OR per_codigo=per_codigo_ref4 OR per_codigo=per_codigo_ref5 OR per_codigo=per_codigo_ref6 " & _
             " OR per_codigo=per_codigo_ref7 OR per_codigo=per_codigo_ref8 OR per_codigo=per_codigo_ref9 " & _
             " OR per_codigo=per_codigo_ref10) "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        cmbN10.BoundText = clsCon_Def.adorec_Def("per_codigo_ref10")
        cmbN9.BoundText = clsCon_Def.adorec_Def("per_codigo_ref9")
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
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_gz=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_di=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_em=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_ee=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_n5=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_n6=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_n7=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_n8=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
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
                " AND p1.per_es_n9=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN9.RowSource = clsCon_Def.adorec_Def
        cmbN9.BoundColumn = "codigo"
        cmbN9.ListField = "nombre"
        
        strSql = " SELECT DISTINCT p1.per_codigo as codigo, CONCAT(p1.per_apellido,' ',p1.per_nombre,' (', tip_ped_nombre ,')') AS nombre " & _
                 " FROM persona as p1 INNER JOIN tipo_pedido ON p1.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND p1.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE p1.emp_codigo='" & strEmpresa & "' " & _
                " AND p1.cat_p_tipo='C'" & _
                " AND p1.per_es_n10=1 AND p1.tip_ped_codigo='" & Negocio & "'" & _
                " ORDER BY nombre "
        clsCon_Def.Ejecutar strSql
        Set cmbN10.RowSource = clsCon_Def.adorec_Def
        cmbN10.BoundColumn = "codigo"
        cmbN10.ListField = "nombre"
                
        If CodPer <> "" Then
            strSql = " SELECT per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5," & _
                     " per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9,per_codigo_ref10 " & _
                     " FROM persona " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND cat_p_tipo='C' AND tip_ped_codigo='" & Negocio & "'" & _
                    " AND per_codigo='" & CodPer & "'"
            clsCon_Def.Ejecutar strSql
            
            cmbN10.BoundText = clsCon_Def.adorec_Def("per_codigo_ref10")
            cmbN9.BoundText = clsCon_Def.adorec_Def("per_codigo_ref9")
            cmbN8.BoundText = clsCon_Def.adorec_Def("per_codigo_ref8")
            cmbN7.BoundText = clsCon_Def.adorec_Def("per_codigo_ref7")
            cmbN6.BoundText = clsCon_Def.adorec_Def("per_codigo_ref6")
            cmbN5.BoundText = clsCon_Def.adorec_Def("per_codigo_ref5")
            cmbEjecutivo.BoundText = clsCon_Def.adorec_Def("per_codigo_ref4")
            cmbEmprendedor.BoundText = clsCon_Def.adorec_Def("per_codigo_ref3")
            cmbDirector.BoundText = clsCon_Def.adorec_Def("per_codigo_ref2")
            cmbGerente.BoundText = clsCon_Def.adorec_Def("per_codigo_ref")
            
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

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

Private Sub optN10_Click()
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
    cmbN10.Locked = True
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
    ElseIf optN10.Value = True Then
        cmbN10.Locked = False
    End If
End Sub
