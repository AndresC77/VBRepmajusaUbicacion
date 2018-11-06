VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelCaracteristica 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Característica"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelCaracteristica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5685
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2895
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1335
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dcmbNombre 
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      Top             =   975
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin VB.Label lblDato 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   5400
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5400
   End
   Begin VB.Label lblCodigo 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5385
   End
   Begin VB.Label lblCaracteristica 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Zona:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1035
      Width           =   1680
   End
End
Attribute VB_Name = "frmSelCaracteristica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################'
'#  Forma para la seleccion de Zonas, y poder modificar o                       #
'#  crear o eliminar zonas                                                      #
'#  frmSelZona V1.0                                                             #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Ventana para consultar las zonas que al momento estan                       #
'#  ingresadas en el sistema. Desde esta ventana se puede crear una nueva       #
'#  zona o modificar o eliminar las zonas ya creadas.                           #
'#  Desde esta ventana se llama a la ventana frmZona en la que se crea          #
'#  y modifica las zonas                                                        #
'#                                                                              #
'#  Tablas que se maneja:                                                       #
'#    documento: En esta tabla se almacenan las nuevas zonas, se                #
'#               modifican los datos de las zonas y se eliminan.                #
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

Public clsCaracteristica As New clsConsulta

Private Sub cmdAceptar_Click()
    frmCambioProductos.Tag = dcmbNombre.BoundText
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    frmCambioProductos.Tag = frmSelCaracteristica.Tag
    Unload Me
End Sub

Private Sub dcmbNombre_Change()
    If dcmbNombre.MatchedWithList = True Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Set dcmbNombre.RowSource = clsCaracteristica.adorec_Def.DataSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsCaracteristica = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub

