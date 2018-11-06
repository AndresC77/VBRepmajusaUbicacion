VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPermisos 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmPermisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11055
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10815
      Begin VB.CheckBox chkFiltroTipo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Tipo de Permiso"
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
         Left            =   3600
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltroFecha 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar por fecha"
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
         Left            =   7200
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame fraFecha 
         BackColor       =   &H00DDDDDD&
         Height          =   1500
         Left            =   7200
         TabIndex        =   16
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option1"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   210
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Rango de Fechas"
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   585
            Width           =   1815
         End
         Begin VB.ComboBox cmbMesI 
            Height          =   315
            ItemData        =   "frmPermisos.frx":030A
            Left            =   1320
            List            =   "frmPermisos.frx":0335
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDDD&
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   255
         End
         Begin MSComCtl2.DTPicker Fecha1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   480
            TabIndex        =   21
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   60882947
            CurrentDate     =   37463
         End
         Begin MSComCtl2.DTPicker Fecha2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   22
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   60882947
            CurrentDate     =   37463
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000050&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Final"
            Enabled         =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1920
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblMes 
            BackColor       =   &H002F1905&
            BackStyle       =   0  'Transparent
            Caption         =   "Por mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CheckBox chkFiltroCodigo 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Filtrar Descripción"
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
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar / Recargar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   720
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo cmbTipoPermiso 
         Height          =   315
         Left            =   3600
         TabIndex        =   30
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Permiso"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00000050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   495
         Width           =   3255
      End
   End
   Begin VB.Frame fraBotones 
      BackColor       =   &H00DDDDDD&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   7320
      Width           =   10815
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Cerrar"
         Height          =   360
         Left            =   5600
         TabIndex        =   7
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   3755
         TabIndex        =   6
         Top             =   240
         Width           =   1700
      End
   End
   Begin VB.Frame fraDatos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Permisos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   10815
      Begin VB.OptionButton optPendientes 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Pendientes"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8640
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optEjecutados 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Ejecutados"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8640
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Mostrar Todos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8640
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   3375
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   10335
         _cx             =   18230
         _cy             =   5953
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
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPermisos.frx":039E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         ExplorerBar     =   3
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
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60882945
            CurrentDate     =   39449
         End
         Begin MSComCtl2.DTPicker dtpHora 
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   60882945
            CurrentDate     =   39449
         End
      End
      Begin MSDataListLib.DataCombo cmbEmpleado 
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin NEED2.uctrVSFG ucrtVSFG 
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   4695
         _extentx        =   8281
         _extenty        =   661
      End
      Begin VB.Label lblFacultad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   480
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private clsSql As New clsConsulta
Private strSql As String
Private FechaI As Variant
Private FechaF As Variant


Private Sub btnAceptar_Click()
    Dim i As Long, control As Integer, codigo As String
    control = 0
    If VSFG.Rows > 1 Then
        If cmbEmpleado.BoundText <> "" Then
            VSFG.Select 1, VSFG.Cols - 1
            VSFG.Sort = flexSortGenericDescending
            For i = 1 To VSFG.Rows - 1
                If VSFG.TextMatrix(i, VSFG.Cols - 1) = 3 Then
                    If Trim(VSFG.TextMatrix(i, 3)) = "" Then
                        MsgBox "No puede modificar Permiso, falta la Descripción", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 4)) = "" Then
                        MsgBox "No puede modificar Permiso, falta la Fecha de Inicio", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 5)) = "" Or Format(VSFG.TextMatrix(i, 5), "HH:mm") = "00:00" Then
                        MsgBox "No puede modificar Permiso, falta la Hora de Inicio", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 6)) = "" Then
                        MsgBox "No puede modificar Permiso, falta la Fecha de Finalización", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 7)) = "" Or Format(VSFG.TextMatrix(i, 7), "HH:mm") = "00:00" Then
                        MsgBox "No puede modificar Permiso, falta la Hora de Finalización", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf FormatoFecha(VSFG.TextMatrix(i, 4)) > FormatoFecha(VSFG.TextMatrix(i, 6)) Then
                        MsgBox "No puede modificar Permiso, la Fecha de Finalización debe ser mayor que Fecha de Inicio", vbCritical, "Modificación de Permiso"
                        control = 1
                    ElseIf FormatoFecha(VSFG.TextMatrix(i, 4)) = FormatoFecha(VSFG.TextMatrix(i, 6)) And FormatoHora(VSFG.TextMatrix(i, 5)) > FormatoHora(VSFG.TextMatrix(i, 7)) Then
                        MsgBox "No puede modificar Permiso, la Hora de Finalización debe ser mayor que Hora de Inicio", vbCritical, "Modificación de Permiso"
                        control = 1
                    Else
                        strSql = " UPDATE permiso SET " & _
                                 " tip_per_codigo='" & VSFG.TextMatrix(i, 2) & "', " & _
                                 " per_inicio='" & FormatoFecha(VSFG.TextMatrix(i, 4)) & "', " & _
                                 " per_horainicio='" & FormatoHora(VSFG.TextMatrix(i, 5)) & "', " & _
                                 " per_fin='" & FormatoFecha(VSFG.TextMatrix(i, 6)) & "', " & _
                                 " per_horafin='" & FormatoHora(VSFG.TextMatrix(i, 7)) & "', " & _
                                 " per_descripcion='" & Trim(VSFG.TextMatrix(i, 3)) & "', " & _
                                 " per_justificado='" & Abs(FormatoD0(VSFG.TextMatrix(i, 8))) & "', " & _
                                 " per_fechamod=CURRENT_TIMESTAMP, " & _
                                 " per_usumod='" & strUsuario & "' " & _
                                 " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                                 " AND tip_per_codigo='" & VSFG.TextMatrix(i, 9) & "' " & _
                                 " AND emp_codigo='" & strEmpresa & "' " & _
                                 " AND epl_codigo='" & cmbEmpleado.BoundText & "' "
                        clsSql.Ejecutar strSql, "M"
                    End If
                ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 2 Then
                    If Trim(VSFG.TextMatrix(i, 3)) = "" Then
                        MsgBox "No puede ingresar Permiso, falta la Descripción", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 4)) = "" Then
                        MsgBox "No puede ingresar Permiso, falta la Fecha de Inicio", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 5)) = "" Or Format(VSFG.TextMatrix(i, 5), "HH:mm") = "00:00" Then
                        MsgBox "No puede ingresar Permiso, falta la Hora de Inicio", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 6)) = "" Then
                        MsgBox "No puede ingresar Permiso, falta la Fecha de Finalización", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf Trim(VSFG.TextMatrix(i, 7)) = "" Or Format(VSFG.TextMatrix(i, 7), "HH:mm") = "00:00" Then
                        MsgBox "No puede ingresar Permiso, falta la Hora de Finalización", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf FormatoFecha(VSFG.TextMatrix(i, 4)) > FormatoFecha(VSFG.TextMatrix(i, 6)) Then
                        MsgBox "No puede ingresar Permiso, la Fecha de Finalización debe ser mayor que Fecha de Inicio", vbCritical, "Ingreso de Permiso"
                        control = 1
                    ElseIf FormatoFecha(VSFG.TextMatrix(i, 4)) = FormatoFecha(VSFG.TextMatrix(i, 6)) And FormatoHora(VSFG.TextMatrix(i, 5)) > FormatoHora(VSFG.TextMatrix(i, 7)) Then
                        MsgBox "No puede ingresar Permiso, la Hora de Finalización debe ser mayor que Hora de Inicio", vbCritical, "Ingreso de Permiso"
                        control = 1
                    Else
                        strSql = " SELECT COALESCE(max(per_codigo),0)+1 " & _
                                 " FROM permiso " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND epl_codigo = '" & cmbEmpleado.BoundText & "' " & _
                                 " AND tip_per_codigo='" & VSFG.TextMatrix(i, 2) & "' "
                        clsSql.Ejecutar strSql
                        
                        codigo = "1"
                        If clsSql.adorec_Def.RecordCount > 0 Then
                            codigo = FormatoD0(clsSql.adorec_Def(0))
                        End If
                        
                        'Busca codigo existente
                        strSql = " SELECT per_codigo" & _
                                " FROM permiso " & _
                                " WHERE emp_codigo='" & strEmpresa & "' " & _
                                " AND epl_codigo = '" & cmbEmpleado.BoundText & "' " & _
                                " AND tip_per_codigo='" & VSFG.TextMatrix(i, 2) & "' " & _
                                " AND per_codigo='" & codigo & "' "
                        clsSql.Ejecutar strSql
                        
                        If clsSql.adorec_Def.RecordCount = 0 Then
                            strSql = " INSERT INTO permiso(emp_codigo,epl_codigo,per_codigo,tip_per_codigo,per_inicio,per_horainicio,per_fin,per_horafin,per_descripcion,per_fechamod,per_usumod,per_justificado) " & _
                                     " VALUES('" & strEmpresa & "','" & cmbEmpleado.BoundText & "','" & codigo & "','" & VSFG.TextMatrix(i, 2) & "','" & FormatoFecha(VSFG.TextMatrix(i, 4)) & "','" & FormatoHora(VSFG.TextMatrix(i, 5)) & "'," & _
                                     "'" & FormatoFecha(VSFG.TextMatrix(i, 6)) & "','" & FormatoHora(VSFG.TextMatrix(i, 7)) & "','" & Trim(VSFG.TextMatrix(i, 3)) & "',CURRENT_TIMESTAMP,'" & strUsuario & "','" & Abs(FormatoD0(VSFG.TextMatrix(i, 8))) & "')"
                            clsSql.Ejecutar strSql, "M"
                        Else
                            MsgBox "El Código de Permiso ya existe, ingrese otro diferente", vbCritical, "Ingreso de Permiso"
                            control = 1
                        End If
                    End If
                ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) = 1 Then
                    strSql = " DELETE FROM permiso " & _
                             " WHERE per_codigo='" & VSFG.TextMatrix(i, 1) & "' " & _
                             " AND emp_codigo='" & strEmpresa & "' " & _
                             " AND tip_per_codigo='" & VSFG.TextMatrix(i, 2) & "' " & _
                             " AND epl_codigo='" & cmbEmpleado.BoundText & "' "
                    clsSql.Ejecutar strSql, "M"
                ElseIf VSFG.TextMatrix(i, VSFG.Cols - 1) <= 0 Then
                    Exit For
                End If
            Next i
            If control = 0 Then
                Limpiar
            End If
        Else
            MsgBox "Seleccione primero un Empleado", vbCritical, "Ingreso de Permisos"
        End If
    End If
End Sub


Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub chkFechas_Click()
    If chkFechas.value = 1 Then
        Label22.Caption = "Fecha Inicial"
        Label23.Enabled = True
        Fecha2.Enabled = True
    Else
        Fecha2 = Fecha1
        Label22.Caption = "Fecha"
        Label23.Enabled = False
        Fecha2.Enabled = False
    End If
End Sub

Private Sub chkFiltroCodigo_Click()
    If chkFiltroCodigo.value = 1 Then
        txtCodigo.Enabled = True
    Else
        txtCodigo.Enabled = False
    End If
End Sub

Private Sub chkFiltroTipo_Click()
    If chkFiltroTipo.value = 1 Then
        cmbTipoPermiso.Enabled = True
    Else
        cmbTipoPermiso.Enabled = False
    End If
End Sub

Private Sub cmbEmpleado_Change()
    If cmbEmpleado.Text <> "" Then
        CargarDatos
    End If
End Sub

Private Sub cmbMesI_Click()
    CambiarFecha
End Sub

Private Sub CambiarFecha()
    'If HacerFecha = False Then Exit Sub
    Dim DiaFinal As Integer
        
    FechaI = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-1", "yyyy-MM-dd")
    FechaF = ""
    DiaFinal = 31
    While (IsDate(FechaF) = False)
        FechaF = Format(Year(HoyDia) & "-" & cmbMesI.ListIndex + 1 & "-" & DiaFinal, "yyyy-MM-dd")
        DiaFinal = DiaFinal - 1
    Wend
End Sub

Private Sub chkFiltroFecha_Click()
    If chkFiltroFecha.value = 1 Then
        fraFecha.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        
        If Option1.value = True Then
            lblMes.Enabled = True
            cmbMesI.Enabled = True
        ElseIf Option2.value = True Then
            Fecha1.Enabled = True
            Label22.Enabled = True
            Fecha1.Enabled = True
            chkFechas.Enabled = True
            If chkFechas.value = 1 Then
                Label23.Enabled = True
                Fecha2.Enabled = True
            End If
        End If
    Else
        fraFecha.Enabled = False
        
        Fecha2.Enabled = False
        Label22.Enabled = False
        Fecha1.Enabled = False
        Label23.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
        
        Option1.Enabled = False
        Option2.Enabled = False
        lblMes.Enabled = False
        cmbMesI.Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click()
    CargarDatos
End Sub

Private Sub dtpFecha_Change()
    If Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "2" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "2"
    ElseIf Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "3" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "3"
    End If
    
    If VSFG.Col = 4 Or VSFG.Col = 6 Then
        'If FormatoFecha(VSFG.TextMatrix(VSFG.Row, 3)) <> "" And FormatoFecha(VSFG.TextMatrix(VSFG.Row, 5)) <> "" _
        'And FormatoFecha(VSFG.TextMatrix(VSFG.Row, 3)) > FormatoFecha(VSFG.TextMatrix(VSFG.Row, 5)) Then
            'MsgBox "La Fecha de Finalización debe ser mayor a la Fecha de Inicio", vbCritical, "Fecha de Finalización"
            'VSFG.TextMatrix(VSFG.Row, 5) = ""
        'Else
            VSFG.Text = FormatoFecha(dtpFecha.value)
        'End If
    End If
End Sub

Private Sub dtpHora_Change()
    If Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "2" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "2"
    ElseIf Right(VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1), 1) = "3" Then
        VSFG.TextMatrix(VSFG.Row, VSFG.Cols - 1) = "3"
    End If
    VSFG.Text = dtpHora.value
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            VSFG = dtpFecha.Tag
            dtpFecha.Visible = False
        Case vbKeyReturn
            dtpFecha.Visible = False
    End Select
End Sub

Private Sub dtpHora_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            VSFG = dtpHora.Tag
            dtpHora.Visible = False
        Case vbKeyReturn
            dtpHora.Visible = False
    End Select
End Sub

Private Sub dtpFecha_LostFocus()
    dtpFecha.Visible = False
End Sub

Private Sub dtpHora_LostFocus()
    dtpHora.Visible = False
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    clsSql.Inicializar AdoConn, AdoConnMaster
    Set ucrtVSFG.VSFGControl = VSFG
    dtpFecha.Format = dtpCustom
    dtpFecha.CustomFormat = "yyyy-MM-dd"
    dtpFecha.Visible = False
    
    dtpHora.Format = dtpCustom
    dtpHora.CustomFormat = "HH:mm"
    dtpHora.UpDown = True
    dtpHora.Hour = "00"
    dtpHora.Minute = "00"
    dtpHora.Second = "00"
    dtpHora.Visible = False
    
    CargaEmpleados
    
    
    strSql = " SELECT tip_per_codigo as codigo,tip_per_nombre as nombre " & _
             " FROM tipo_permiso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 2 "
    clsSql.Ejecutar strSql
    Set cmbTipoPermiso.RowSource = clsSql.adorec_Def.DataSource
    cmbTipoPermiso.ListField = "nombre"
    cmbTipoPermiso.BoundColumn = "codigo"
        
    
    chkFiltroFecha.value = 0
    optTodos.value = True
    Dim i As Integer
    Fecha1 = Format(HoyDia, "yyyy-MM-dd")
    Fecha2 = Format(HoyDia, "yyyy-MM-dd")
    For i = 0 To 11
        If (cmbMesI.ItemData(i) = Month(HoyDia)) Then
            cmbMesI.ListIndex = i
            Exit For
        End If
    Next i
    
End Sub

Private Sub Limpiar()
    CargarDatos
End Sub

Private Sub CargarDatos()
    VSFG.Rows = 1
    VSFG.Clear 1
    If cmbEmpleado.Text <> "" Then
        strSql = " SELECT per_codigo,tip_per_codigo,per_descripcion,per_inicio,TIME_FORMAT(per_horainicio,'%H:%i') as per_horainicio,per_fin,TIME_FORMAT(per_horafin,'%H:%i') as per_horafin,per_justificado,tip_per_codigo,'0' as modo " & _
                 " FROM permiso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND epl_codigo='" & cmbEmpleado.BoundText & "' "
        If optPendientes.value = True Then
            strSql = strSql & " AND CURRENT_DATE BETWEEN per_inicio AND per_fin "
        ElseIf optEjecutados.value = True Then
            strSql = strSql & " AND per_fin < CURRENT_DATE "
        End If
        
        If chkFiltroCodigo.value = 1 Then
            strSql = strSql & " AND per_descripcion LIKE '%" & txtCodigo.Text & "%' "
        End If
        
        If chkFiltroTipo.value = 1 Then
            strSql = strSql & " AND tip_per_codigo = '" & cmbTipoPermiso.BoundText & "' "
        End If
        
        If chkFiltroFecha.value = 1 Then
            If Option1.value = True Then
                strSql = strSql & " AND per_inicio BETWEEN '" & FechaI & "' AND '" & FechaF & "' "
            ElseIf Option2.value = True Then
               If chkFechas.value = 0 Then
                    strSql = strSql & " AND per_inicio BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
                Else
                    strSql = strSql & " AND per_inicio BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' AND per_fin BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "' "
                End If
            End If
        End If
        
        
        
        strSql = strSql & " ORDER BY per_inicio,per_horainicio,per_fin,per_horafin DESC "
        clsSql.Ejecutar strSql
        Set VSFG.DataSource = clsSql.adorec_Def.DataSource
        ucrtVSFG.PonerNum
        
        VSFG.ColComboList(4) = "Dummy"
        VSFG.ColComboList(6) = "Dummy"
        
        
        
    End If
    strSql = " SELECT tip_per_codigo,tip_per_nombre " & _
                 " FROM tipo_permiso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY 2 "
        clsSql.Ejecutar strSql
        VSFG.ColComboList(2) = VSFG.BuildComboList(clsSql.adorec_Def, "*tip_per_nombre", "tip_per_codigo")
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
End Sub

Private Sub CargaEmpleados()
    strSql = " SELECT CONCAT(epl_apellidos,' ',epl_nombres) as nombre,epl_codigo as codigo " & _
             " FROM empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 1 "
    clsSql.Ejecutar strSql
    Set cmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    cmbEmpleado.ListField = "nombre"
    cmbEmpleado.BoundColumn = "codigo"
End Sub


Private Sub optEjecutados_Click()
    CargarDatos
End Sub

Private Sub Option1_Click()
    If Option1.value = True Then
        lblMes.Enabled = True
        cmbMesI.Enabled = True
        
        Fecha2.Enabled = False
        Label22.Enabled = False
        Fecha1.Enabled = False
        Label23.Enabled = False
        Fecha2.Enabled = False
        chkFechas.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.value = True Then
        lblMes.Enabled = False
        cmbMesI.Enabled = False
        
        Fecha1.Enabled = True
        Label22.Enabled = True
        Fecha1.Enabled = True
        chkFechas.Enabled = True
        If chkFechas.value = 1 Then
            Label23.Enabled = True
            Fecha2.Enabled = True
        End If
    End If
End Sub

Private Sub optPendientes_Click()
    CargarDatos
End Sub

Private Sub optTodos_Click()
    CargarDatos
End Sub

Private Sub VSFG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 2
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        VSFG.TextMatrix(Row, VSFG.Cols - 1) = 3
    End If
End Sub

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 0 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 1 Then
        Cancel = True
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 2 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -2 Then
        If Col >= VSFG.Cols - 1 Then
            Cancel = True
        Else
            If Col = 4 Or Col = 6 Then
                If VSFG.TextMatrix(Row, Col) = "" Then
                    VSFG.TextMatrix(Row, Col) = Format(Date, "yyyy-MM-dd")
                End If
            ElseIf Col = 5 Or Col = 7 Then
                If VSFG.TextMatrix(Row, Col) = "" Then
                    VSFG.TextMatrix(Row, Col) = Format(Time, "HH:00")
                End If
            End If
        End If
    ElseIf Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = 3 Or Val(VSFG.TextMatrix(Row, VSFG.Cols - 1)) = -3 Then
        If Col = 1 Or Col >= VSFG.Cols - 1 Then
            Cancel = True
        Else
            If Col = 4 Or Col = 6 Then
                If VSFG.TextMatrix(Row, Col) = "" Then
                    VSFG.TextMatrix(Row, Col) = Format(Date, "yyyy-MM-dd")
                End If
            ElseIf Col = 5 Or Col = 7 Then
                If VSFG.TextMatrix(Row, Col) = "" Then
                    VSFG.TextMatrix(Row, Col) = Format(Time, "HH:00")
                End If
            End If
        End If
    End If
End Sub

Private Sub VSFG_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpFecha.Visible Then Cancel = True
    If dtpHora.Visible Then Cancel = True
End Sub

Private Sub VSFG_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If dtpFecha.Visible Then Cancel = True
    If dtpHora.Visible Then Cancel = True
End Sub

Private Sub VSFG_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 4 Or Col = 6 Then
        If VSFG.ColDataType(Col) = flexDTDate Then
            Cancel = True
            dtpFecha.Move VSFG.CellLeft, VSFG.CellTop, VSFG.CellWidth, VSFG.CellHeight
            dtpFecha.value = VSFG
            dtpFecha.Tag = VSFG
            
            dtpFecha.Visible = True
            dtpFecha.SetFocus
            
            SendKeys vbKeyF4
        End If
    End If
    
    If Col = 5 Or Col = 7 Then
        If VSFG.ColDataType(Col) = flexDTDate Then
            Cancel = True
            dtpHora.Move VSFG.CellLeft, VSFG.CellTop, VSFG.CellWidth, VSFG.CellHeight
            dtpHora.value = VSFG
            dtpHora.Tag = VSFG
            
            dtpHora.Visible = True
            dtpHora.SetFocus
            
            SendKeys vbKeyF4
        End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Sub VSFG_KeyPress(KeyAscii As Integer)
    ucrtVSFG.Editar KeyAscii
End Sub

Private Sub VSFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And VSFG.MouseRow > 0 Then
        ucrtVSFG.VerMenu
    End If
End Sub


