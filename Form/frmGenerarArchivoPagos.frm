VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGenerarArchivoPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos Automáticos"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerarArchivoPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13395
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Aplicación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   13200
      Begin VB.CommandButton cmdGenerarArchivo 
         Caption         =   "&Generar Archivo"
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsultaCartera 
         Caption         =   "Consulta Cartera a Pagar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtFormato 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   10695
      End
      Begin VB.CommandButton cmdAplicarFormato 
         Caption         =   "Aplicar Formato"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtEncabezado 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2760
         Width           =   10695
      End
      Begin VB.TextBox txtTotalACobrar 
         Height          =   375
         Left            =   11280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtArchivo 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2400
         Width           =   4695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   12975
         _cx             =   2088786278
         _cy             =   2088766382
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
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmGenerarArchivoPagos.frx":030A
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
         TabBehavior     =   0
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
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG2 
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   3480
         Width           =   12975
         _cx             =   2088786278
         _cy             =   2088767863
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
         FormatString    =   $"frmGenerarArchivoPagos.frx":04FC
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
         TabBehavior     =   0
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
      End
      Begin NEED2.uctrVSFG ucrtVSFG1 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin NEED2.dtpFecha FechaValidez 
         Height          =   315
         Left            =   9360
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Validez:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8760
         TabIndex        =   15
         Top             =   2452
         Width           =   585
      End
   End
   Begin MSDataListLib.DataCombo cmbBanco 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   12720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Archivo de Backup"
      InitDir         =   "C:\"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   510
   End
End
Attribute VB_Name = "frmGenerarArchivoPagos"
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
Private strSQL As String

Private Sub cmbBanco_Change()
    strSQL = " SELECT ban_cas_encabezado,ban_cas_formato,ban_cas_archivo,ban_cas_formato_archivo " & _
             " FROM banco_cash  " & _
             " WHERE ban_codigo='" & cmbBanco.BoundText & "' AND ban_cas_tipo='P' "
    clsCon_Def.Ejecutar strSQL
    txtEncabezado.Text = clsCon_Def.adorec_Def("ban_cas_encabezado")
    txtFormato.Text = clsCon_Def.adorec_Def("ban_cas_formato")
    txtFormato.Tag = clsCon_Def.adorec_Def("ban_cas_formato_archivo")
    txtArchivo.Text = clsCon_Def.adorec_Def("ban_cas_archivo")
End Sub

Private Sub cmdAplicarFormato_Click()
    Dim i As Long
    Dim j As Long
    Dim maxj As Long
    Dim Formato() As String
    Dim Encabezado() As String
    Dim CeldaFormato() As String
    Dim Linea As String
    Dim Celda As String
    Dim aux As String
    Dim NEntero As String
    Dim NDecimal As String
    Dim Separador As String
    If UCase(txtFormato.Tag) = "TAB" Then
        Separador = vbTab
    Else
        Separador = txtFormato.Tag
    End If
    maxj = 0
    VSFG2.Rows = 0
    If Trim(txtEncabezado.Text) <> "" Then
        Encabezado = Split(txtEncabezado.Text, ";")
        maxj = UBound(Encabezado)
        If UCase(txtFormato.Tag) = "TAB" Then
            VSFG2.Cols = maxj + 1
        Else
            VSFG2.Cols = 1
            VSFG2.ColWidth(0) = 1000
            VSFG2.ColWidthMax = 1000
            
            
        End If
        Linea = ""
        For j = 0 To maxj
            If Left(Encabezado(j), 1) = "(" Then
                aux = Mid(Encabezado(j), 2, Len(Encabezado(j)) - 2)
                CeldaFormato = Split(aux, ",")
            '01;REC;00017;00111;01;(YYYYMMDD);(YYYYMMDD);(NUMREG);(TOTALENVIO);;
                If CeldaFormato(0) = "HOYDIA" Then
                    Linea = Linea & Format(HoyDia, CeldaFormato(1)) & Separador
                ElseIf CeldaFormato(0) = "NUMREG" Then
                    Linea = Linea & Format(VSFG.Rows - 1, CeldaFormato(1)) & Separador
                ElseIf CeldaFormato(0) = "TOTALENVIO" Then
                    NEntero = Format(Int(FormatoD2(txtTotalACobrar.Text)), "#####0")
                    NDecimal = Format(FormatoD0(FormatoD2((FormatoD2(txtTotalACobrar.Text) - FormatoD2(NEntero))) * 100#), "00")
                    Linea = Linea & Format(NEntero & NDecimal, CeldaFormato(1)) & Separador
                End If
            ElseIf Left(Encabezado(j), 1) = """" Then
                aux = Mid(Encabezado(j), 2, Len(Encabezado(j)) - 2)
                Linea = Linea & aux & Separador
            End If
        Next j
        VSFG2.AddItem Linea
    End If
    
    Formato = Split(txtFormato.Text, ";")
    If maxj < UBound(Formato) Then
        If UCase(txtFormato.Tag) = "TAB" Then
            VSFG2.Cols = UBound(Formato) + 1
        Else
            VSFG2.Cols = 1
            VSFG2.ColWidth(0) = 10000
            VSFG2.ColWidthMax = 10000
            VSFG2.ColWidthMin = 100000
            
        End If
    End If
    maxj = UBound(Formato)
    For i = 1 To VSFG.Rows - 1
        
        Linea = ""
        For j = 0 To maxj
            If Left(Formato(j), 1) = "(" Then
                aux = Mid(Formato(j), 2, Len(Formato(j)) - 2)
                CeldaFormato = Split(aux, ",")
                
                If CeldaFormato(0) = "HOYDIA" Then
                    Celda = Format(HoyDia, CeldaFormato(1))
                ElseIf CeldaFormato(0) = "SECUENCIAL" Then
                    Celda = Format(i, CeldaFormato(1))
                ElseIf CeldaFormato(0) = "VALIDEZ" Then
                    Celda = Format(FechaValidez.Value, CeldaFormato(1))
                ElseIf IsNumeric(CeldaFormato(0)) = True Then
                    If Left(CeldaFormato(1), 4) = "LSET" Then
                        Celda = JustificarIzquierdaConEspacios(VSFG.TextMatrix(i, CeldaFormato(0)), Mid(CeldaFormato(1), 6, Len(CeldaFormato(1)) - 6))
                    Else
                        Celda = Format(VSFG.TextMatrix(i, CeldaFormato(0)), CeldaFormato(1))
                    End If
                End If
                Linea = Linea & Celda & Separador
            ElseIf Left(Formato(j), 1) = "[" Then
                aux = Mid(Formato(j), 2, Len(Formato(j)) - 2)
                CeldaFormato = Split(aux, ",")
                Celda = VSFG.TextMatrix(i, CeldaFormato(0))
                NEntero = Format(Int(FormatoD2(Celda)), "#####0")
                NDecimal = Format(FormatoD0(FormatoD2((FormatoD2(Celda) - FormatoD2(NEntero))) * 100#), "00")
                Linea = Linea & Format(NEntero & NDecimal, CeldaFormato(1)) & Separador
            ElseIf Left(Formato(j), 1) = """" Then
                aux = Mid(Formato(j), 2, Len(Formato(j)) - 2)
                Linea = Linea & aux & Separador
            End If
        Next j
        VSFG2.AddItem Linea
    Next i
End Sub

Private Function JustificarIzquierdaConEspacios(cadena As String, largo As Long) As String
    Dim i As Long
    If largo < Len(cadena) Then
        JustificarIzquierdaConEspacios = Left(cadena, largo)
    Else
        JustificarIzquierdaConEspacios = cadena
        For i = 1 To largo - Len(cadena)
            JustificarIzquierdaConEspacios = JustificarIzquierdaConEspacios & " "
        Next i
    End If
End Function

Private Sub cmdConsultaCartera_Click()
    Dim i As Long
    strSQL = " SELECT comp_egreso.com_egr_codigo,CONCAT(per_apellido, ' ',per_nombre) as pro," & _
             " IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, com_egr_ch_valor," & _
             " com_egr_descripcion,ban_codigo_interbancario,per_cue_tipo,per_cue_numero,per_direccion,ciu_nombre,per_telf,per_email,comp_egreso.asi_numasiento " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.emp_codigo=persona.emp_codigo" & _
                 " AND comp_egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN persona_cuenta ON persona.emp_codigo=persona_cuenta.emp_codigo" & _
                 " AND persona.per_codigo=persona_cuenta.per_codigo " & _
                 " INNER JOIN banco ON " & _
                 " persona_cuenta.ban_codigo=banco.ban_codigo " & _
                 " INNER JOIN ciudad ON " & _
                 " persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' AND com_egr_proceso_cash=2 AND com_egr_ch_valor!=0 " & _
                 " ORDER BY CONCAT(per_apellido, ' ',per_nombre) "
    clsCon_Def.Ejecutar strSQL
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    txtTotalACobrar.Text = FormatoD2(0)
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 1) = QuitarCaracteresEspeciales(VSFG.TextMatrix(i, 1))
        VSFG.TextMatrix(i, 5) = QuitarCaracteresEspeciales(VSFG.TextMatrix(i, 5))
        txtTotalACobrar.Text = FormatoD2(FormatoD2(txtTotalACobrar.Text) + FormatoD2(VSFG.TextMatrix(i, 4)))
    Next i
End Sub

Private Function QuitarCaracteresEspeciales(cadena As String) As String
    Dim i As Long
    Dim CadenaFinal As String
    Dim Caracter As String
    CadenaFinal = ""
    For i = 1 To Len(cadena)
        Caracter = Mid(cadena, i, 1)
        If Caracter <> " " Then
            '65-90 A-Z 165 Ñ
            Caracter = Replace(Replace(Caracter, "Ñ", "N"), vbNewLine, " ")
            If Not (Asc(Caracter) >= 65 And Asc(Caracter) <= 90) Or IsNumeric(Caracter) = True Then
                Caracter = ""
            End If
        End If
        CadenaFinal = CadenaFinal & Caracter
    Next i
    QuitarCaracteresEspeciales = CadenaFinal
End Function

Private Sub cmdGenerarArchivo_Click()
    Dim sDir As String
    Dim ArchivoDefecto As String
    ArchivoDefecto = ""
    If txtArchivo.Text <> "" Then
        If InStr(1, txtArchivo.Text, "[") > 0 Then
            If InStr(1, txtArchivo.Text, "]") > 0 Then
                ArchivoDefecto = Left(txtArchivo.Text, InStr(1, txtArchivo.Text, "[") - 1) & Format(HoyDia, Right(Left(txtArchivo.Text, InStr(1, txtArchivo.Text, "]") - 1), Len(Left(txtArchivo.Text, InStr(1, txtArchivo.Text, "]") - 1)) - (InStr(1, txtArchivo.Text, "[")))) & Right(txtArchivo.Text, Len(txtArchivo.Text) - InStr(1, txtArchivo.Text, "]"))
            End If
        End If
    End If
    sDir = CurDir
    cdArchivo.Filter = "Todos los Archivos|*.*|Archivos de texto .txt|*.txt"
    cdArchivo.FileName = ArchivoDefecto
    cdArchivo.ShowSave
    ChDir sDir
    If (cdArchivo.FileName <> "") Then
        Me.MousePointer = 11
        VSFG2.SaveGrid cdArchivo.FileName, flexFileTabText
        For i = 1 To VSFG.Rows - 1
            strSQL = " UPDATE comp_egreso  " & _
                     " SET com_egr_proceso_cash=3 " & _
                     " WHERE comp_egreso.emp_codigo = '" & strEmpresa & "' " & _
                     " AND comp_egreso.com_egr_codigo='" & VSFG.TextMatrix(i, 0) & "' " & _
                     " AND com_egr_proceso_cash=2 AND com_egr_ch_valor!=0 "
            clsCon_Def.Ejecutar strSQL
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

Private Sub Form_Load()
    'Centra esta forma dentro de la forma MDI
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Set ucrtVSFG1.VSFGControl = VSFG
    ucrtVSFG1.Inicializar False, False, False
    FechaValidez.Value = HoyDia
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
        strSQL = " SELECT banco_cash.ban_codigo, ban_nombre " & _
                 " FROM banco_cash INNER JOIN banco ON banco_cash.ban_codigo=banco.ban_codigo " & _
                 " WHERE ban_cas_tipo='P' ORDER BY ban_nombre "
        clsCon_Def.Ejecutar strSQL
        Set cmbBanco.RowSource = clsCon_Def.adorec_Def.DataSource
        cmbBanco.ListField = "ban_nombre"
        cmbBanco.BoundColumn = "ban_codigo"
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

Private Sub VSFG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

Private Sub VSFG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True
End Sub

