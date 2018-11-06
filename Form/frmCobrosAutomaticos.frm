VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCobrosAutomaticos 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros Automáticos"
   ClientHeight    =   7455
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
   Icon            =   "frmCobrosAutomaticos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   13395
   Begin VB.CheckBox chkNoBajarPedidos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "No Bajar Pedidos"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox chkNoEnviarSMS 
      BackColor       =   &H00DDDDDD&
      Caption         =   "No Enviar SMS"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkNoCarteraPedidos 
      BackColor       =   &H00DDDDDD&
      Caption         =   "No Generar Cartera de Pedidos"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   600
      Width           =   2895
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFGNegocio 
      Height          =   1575
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   6615
      _cx             =   11668
      _cy             =   2778
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCobrosAutomaticos.frx":030A
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
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   13200
      Begin VB.CommandButton cmdGenerarArchivo 
         Caption         =   "&Generar Archivo"
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsultaCartera 
         Caption         =   "Consulta Cartera a Cobrar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtFormato 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3000
         Width           =   10695
      End
      Begin VB.CommandButton cmdAplicarFormato 
         Caption         =   "Aplicar Formato"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txtEncabezado 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   10695
      End
      Begin VB.TextBox txtTotalACobrar 
         Height          =   375
         Left            =   11280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtArchivo 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   4695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFG 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   600
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCobrosAutomaticos.frx":0372
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
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   12975
         _cx             =   22886
         _cy             =   2778
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
         FormatString    =   $"frmCobrosAutomaticos.frx":04C3
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
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
      End
      Begin NEED2.dtpFecha FechaValidez 
         Height          =   315
         Left            =   9360
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Value           =   41836.4460300926
      End
      Begin NEED2.dtpFecha dtpFechaAl 
         Height          =   315
         Left            =   11760
         TabIndex        =   18
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Value           =   41836.4460300926
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   11520
         TabIndex        =   17
         Top             =   262
         Width           =   195
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   4965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Validez:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8760
         TabIndex        =   15
         Top             =   2325
         Width           =   585
      End
   End
   Begin MSDataListLib.DataCombo cmbBanco 
      Height          =   315
      Left            =   7920
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
      Left            =   7200
      TabIndex        =   1
      Top             =   165
      Width           =   510
   End
End
Attribute VB_Name = "frmCobrosAutomaticos"
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
Private strTipoPedidos As String

Private Sub cmbBanco_Change()
    strSql = " SELECT ban_cas_encabezado,ban_cas_formato,ban_cas_archivo,ban_cas_formato_archivo " & _
             " FROM banco_cash  " & _
             " WHERE ban_codigo='" & cmbBanco.BoundText & "' AND ban_cas_tipo='C' "
    clsCon_Def.Ejecutar strSql
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
        Separador = Trim(txtFormato.Tag)
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
                    NDecimal = Format(FormatoD0(FormatoD2((FormatoD2(txtTotalACobrar.Text) - FormatoD2(NEntero))) * 100), "00")
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
                NDecimal = Format(FormatoD0(FormatoD2((FormatoD2(Celda) - FormatoD2(NEntero))) * 100), "00")
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
    
    strTipoPedidos = ""
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            strTipoPedidos = strTipoPedidos & "'" & VSFGNegocio.TextMatrix(i, 1) & "',"
        End If
    Next i
    strTipoPedidos = strTipoPedidos & "'---'"
    
    LiberarYBajarPedidos IIf(Me.chkNoBajarPedidos.Value = 1, True, False), strTipoPedidos

'************** INCENTIVOS
    'pbGeneral.Value = 1
    'pbGeneral.Refresh
    lblEstado.Caption = "Promo x Combo"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optPromoCombo.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    'pbGeneral.Value = 2
    'pbGeneral.Refresh
    lblEstado.Caption = "Promo Combo Pedido"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optPromoComboPedido.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    'pbGeneral.Value = 3
    'pbGeneral.Refresh
    lblEstado.Caption = "Dcto x Combo"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optDctoCombo.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    'pbGeneral.Value = 4
    'pbGeneral.Refresh optNPrendasAY.Value
    lblEstado.Caption = "n Prendas a $x.xx"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optNPrendasAY.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    'dtpFechaFinAplicar.Value = dtpEjecutar.Value
    'pbGeneral.Value = 5
    'pbGeneral.Refresh
    lblEstado.Caption = "Premio por monto Marca"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optPromoPremioPorMontoMarca.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    
    lblEstado.Caption = "Dcto por Fecha"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optDctoFecha.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    
    
    'pbGeneral.Value = 7
    'pbGeneral.Refresh
    lblEstado.Caption = "Premio x Monto"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optPremio.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    
    lblEstado.Caption = "Cargando Incentivos"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmIncentivos.Show
            frmIncentivos.cmbNegocioAplicar.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmIncentivos.dtpFechaInicioAplicar.Value = Format(DateAdd("d", -7, dtpFechaAl.Value), "yyyy-mm-dd") & " 00:00:00"
            frmIncentivos.dtpFechaFinAplicar.Value = Format(dtpFechaAl.Value, "yyyy-mm-dd") & " 23:59:59"
            frmIncentivos.optIncentivo.Value = True
            frmIncentivos.cmdActualizar_Click
            frmIncentivos.cmdAplicarAplicar_Click
        End If
    Next i
    
    
'    frmIncentivos.CmdSalir_Click
'*****************FIN INCENTIVOS
    
'***************** APLICAR NC A PEDIDOS

    lblEstado.Caption = "Aplicando NC a PEDIDOS"
    lblEstado.Refresh
    For i = 1 To Me.VSFGNegocio.Rows - 1
        If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
            frmAplicarNC.Show
            frmAplicarNC.cmbNegocio.BoundText = VSFGNegocio.TextMatrix(i, 1)
            frmAplicarNC.dtpFechaNCIni.Value = Format(DateAdd("m", -6, dtpFechaAl.Value), "yyyy-mm-dd")
            frmAplicarNC.dtpFechaNCFin.Value = dtpFechaAl.Value
            frmAplicarNC.chkSelTodo.Value = 1
            frmAplicarNC.cmdMostrarNC_Click
            frmAplicarNC.chkIncluirPedidos.Value = 1
            frmAplicarNC.dtpFechaFacIni.Value = Format(DateAdd("m", -3, dtpFechaAl.Value), "yyyy-mm-dd")
            frmAplicarNC.dtpFechaFacFin.Value = dtpFechaAl.Value
            frmAplicarNC.cmdMostrarAplicacion_Click
            frmAplicarNC.cmdAplicar_Click
            'guardar excel
            frmAplicarNC.CmdSalir_Click
        End If
    Next i
    
'***************** FIN APLICAR NC A PEDIDOS
    
'***************** ENVIO DE SMS PEDIDOS
    If chkNoEnviarSMS.Value = 0 Then
        lblEstado.Caption = "Envio SMS"
        lblEstado.Refresh
        For i = 1 To Me.VSFGNegocio.Rows - 1
            If Abs(VSFGNegocio.TextMatrix(i, 0)) = 1 Then
                frmCarteraPedidos.Show
                frmCarteraPedidos.cmbNegocio.BoundText = VSFGNegocio.TextMatrix(i, 1)
                frmCarteraPedidos.cmdActualizar_Click
                frmCarteraPedidos.cmdEnvioCorreo_Click
                frmCarteraPedidos.cmdcancelar_Click
            End If
        Next i
    End If
'***************** FIN ENVIO DE SMS PEDIDOS

    
    lblEstado.Caption = "Consultando CARTERA"
    lblEstado.Refresh

'    strSql = " SELECT cuenta_p_c.cue_p_c_codigo as c1,CONCAT(per_apellido, ' ',per_nombre) as cli,IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
'                 " FROM  cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo" & _
'                 " AND cuenta_p_c.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
'                 " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
'                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
'                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
'                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
'                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
'                 " HAVING round(d,2)>0 " & _
'                 " ORDER BY cue_p_c_egr_codigo,c1"
'    strSql = " SELECT cuenta_p_c.cue_p_c_codigo as c1,CONCAT(per_apellido, ' ',per_nombre) as cli,IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
'                 " FROM  cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo" & _
'                 " AND cuenta_p_c.per_codigo=persona.per_codigo AND persona.tip_ped_codigo='" & cmbNegocio.BoundText & "'" & _
'                 " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo " & _
'                 " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
'                 " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
'                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
'                 " AND IIF(persona.for_pag_codigo NOT IN ('EFE','CONT'),1=1,IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=7,DATEADD(d,6,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=1,DATEADD(d,5,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=2,DATEADD(d,4,cue_p_c_fechaemision),DATEADD(d,6,cue_p_c_fechaemision))))>left(CURRENT_TIMESTAMP,10))" & _
'                 " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo " & _
'                 " HAVING round(d,2)>0 " & _
'                 " ORDER BY cue_p_c_fechaemision,cue_p_c_egr_codigo,c1"
    
    strSql = " DELETE FROM pagoT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM comprobante_retencionT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM cuenta_p_cT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM personaT "
    clsCon_Def.Ejecutar strSql
    
    strSql = " INSERT INTO cuenta_p_cT SELECT cuenta_p_c.* " & _
             " FROM cuenta_p_c " & _
             " WHERE cuenta_p_c.emp_codigo = '" & strEmpresa & "'" & _
             " AND cuenta_p_c.cue_p_c_tipo = 'C'" & _
             " AND cue_p_c_pagado='0'" & _
             " AND cue_p_c_egr_codigo NOT LIKE 'R%'" & _
             " AND tip_doc_cue_codigo=1"
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO personaT SELECT per_codigo,emp_codigo,for_pag_codigo,per_ruc,per_nombre,per_apellido,tip_ped_codigo,per_codigo_resp,per_email,per_celular " & _
             " From Persona " & _
             " WHERE emp_codigo='" & strEmpresa & "' and cat_p_tipo='C' and persona.tip_ped_codigo IN (" & strTipoPedidos & ")"
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO pagoT SELECT pago.pag_codigo,pago.emp_codigo,pago.cue_p_c_codigo,pago.cue_p_c_tipo,pago.pag_monto " & _
                 " FROM  cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo IN (" & strTipoPedidos & ")" & _
                 " INNER JOIN pago ON cuenta_p_cT.emp_codigo=pago.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=pago.cue_p_c_codigo AND pag_monto!=0 " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
                 " "
    clsCon_Def.Ejecutar strSql
    strSql = " INSERT INTO comprobante_retencionT SELECT comprobante_retencion.emp_codigo,comprobante_retencion.cue_p_c_codigo,comprobante_retencion.cue_p_c_tipo,comprobante_retencion.com_ret_total " & _
                 " FROM  cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo IN (" & strTipoPedidos & ")" & _
                 " INNER JOIN comprobante_retencion ON cuenta_p_cT.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
                 " "
    clsCon_Def.Ejecutar strSql
    strSql = " SELECT cuenta_p_cT.cue_p_c_codigo as c1,CONCAT(per_apellido, ' ',per_nombre) as cli,IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                 " FROM  cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo IN (" & strTipoPedidos & ")" & _
                 " LEFT JOIN pagoT ON cuenta_p_cT.emp_codigo=pagoT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=pagoT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=pagoT.cue_p_c_codigo " & _
                 " LEFT JOIN comprobante_retencionT ON cuenta_p_cT.emp_codigo=comprobante_retencionT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=comprobante_retencionT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=comprobante_retencionT.cue_p_c_codigo " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
                 " AND IIF(personaT.for_pag_codigo NOT IN ('EFE','CONT'),1=1,IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=7,DATEADD(d,6,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=1,DATEADD(,d,5,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=2,DATEADD(d,4,cue_p_c_fechaemision),DATEADD(d,6,cue_p_c_fechaemision))))>'" & dtpFechaAl.Value & "')" & _
                 " GROUP BY cuenta_p_cT.cue_p_c_codigo,cuenta_p_cT.cue_p_c_tipo " & _
                 " HAVING round(d,2)>0 " & _
                 " ORDER BY cue_p_c_fechaemision,cue_p_c_egr_codigo,c1"
    strSql = " SELECT cuenta_p_cT.cue_p_c_codigo as c1,CONCAT(per_apellido, ' ',per_nombre) as cli,IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, RIGHT(cue_p_c_egr_codigo,7) as cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor ,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                 " FROM  cuenta_p_cT INNER JOIN personaT ON cuenta_p_cT.emp_codigo=personaT.emp_codigo" & _
                 " AND cuenta_p_cT.per_codigo=personaT.per_codigo AND personaT.tip_ped_codigo IN (" & strTipoPedidos & ")" & _
                 " LEFT JOIN pagoT ON cuenta_p_cT.emp_codigo=pagoT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=pagoT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=pagoT.cue_p_c_codigo " & _
                 " LEFT JOIN comprobante_retencionT ON cuenta_p_cT.emp_codigo=comprobante_retencionT.emp_codigo AND cuenta_p_cT.cue_p_c_tipo=comprobante_retencionT.cue_p_c_tipo AND cuenta_p_cT.cue_p_c_codigo=comprobante_retencionT.cue_p_c_codigo " & _
                 " WHERE cuenta_p_cT.emp_codigo = '" & strEmpresa & "' AND cuenta_p_cT.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                 " AND cue_p_c_egr_codigo NOT LIKE 'R%' and tip_doc_cue_codigo=1 " & _
                 " AND IIF(personaT.for_pag_codigo IN ('EFE','CONT'),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=7,DATEADD(d,6,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=1,DATEADD(d,5,cue_p_c_fechaemision),IIF(DATEPART(dw,DATEADD(d,1,cue_p_c_fechaemision))=2,DATEADD(d,4,cue_p_c_fechaemision),DATEADD(d,6,cue_p_c_fechaemision)))),'" & DateAdd("d", 1, dtpFechaAl.Value) & "')>'" & dtpFechaAl.Value & "'" & _
                 " GROUP BY cuenta_p_cT.cue_p_c_codigo,per_apellido,per_nombre,per_ruc,cue_p_c_egr_codigo,cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor,com_ret_total " & _
                 " HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)>0 "
    If chkNoCarteraPedidos.Value = 0 Then
        strSql = strSql & " UNION" & _
                     " SELECT pedido.ped_codigo,CONCAT(per_apellido, ' ',per_nombre) as cli," & _
                     " IIF(LEN(per_ruc)=13,'R',IIF(LEN(per_ruc)=10,'C','P')),per_ruc, " & _
                     " CONCAT('9',RIGHT(pedido.ped_codigo,7)), CONCAT('PEDIDO: ', pedido.ped_codigo) as cue_p_c_descripcion , " & _
                     " ped_fecha, ped_fecha," & _
                     " SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
                     " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
                     " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
                     " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) ," & _
                     " ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
                     " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
                     " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
                     " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) " & _
                     "-COALESCE(doc_pag_valor,0.00),2) as d "
        strSql = strSql & " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                     " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo IN (" & strTipoPedidos & ")" & _
                     " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo AND det_ped_incentivo=0 " & _
                     " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
                     " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                     " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                     " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                     " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                     " AND pedido.ped_codigo=producto_promo2.ped_codigo" & _
                     " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
                     " FROM doc_pago_pedido " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
                     " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
                     " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
                     " AND pedido.per_codigo=pag.per_codigo " & _
                     " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
                     " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0) AND pedido.ped_fechamod<=DATEADD(n,-10,CURRENT_TIMESTAMP)" & _
                     " AND det_pedido.det_ped_incentivo=0 " & _
                     " GROUP BY pedido.ped_codigo, per_apellido, per_nombre,per_ruc,ped_fecha, ped_fecha, " & _
                     " pedido.ped_dctoadicional,par_numero,doc_pag_valor " & _
                     " HAVING ROUND(ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00),2) " & _
                     "+ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-pedido.ped_dctoadicional/100.00) * (par_numero)/100.00,2) " & _
                     "-COALESCE(doc_pag_valor,0.00),2)>0.01"
    End If
    strSql = strSql & " ORDER BY cue_p_c_fechaemision,cue_p_c_egr_codigo,c1"
    clsCon_Def.Ejecutar strSql
    Set VSFG.DataSource = clsCon_Def.adorec_Def.DataSource
    txtTotalACobrar.Text = FormatoD2(0)
    For i = 1 To VSFG.Rows - 1
        VSFG.TextMatrix(i, 1) = QuitarCaracteresEspecialesYNumeros(VSFG.TextMatrix(i, 1))
        VSFG.TextMatrix(i, 5) = QuitarCaracteresEspeciales(Left(VSFG.TextMatrix(i, 5), 25))
        txtTotalACobrar.Text = FormatoD2(FormatoD2(txtTotalACobrar.Text) + FormatoD2(VSFG.TextMatrix(i, 9)))
    Next i
    strSql = " DELETE FROM pagoT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM comprobante_retencionT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM cuenta_p_cT "
    clsCon_Def.Ejecutar strSql
    strSql = " DELETE FROM personaT "
    clsCon_Def.Ejecutar strSql
    MsgBox Now
End Sub

Private Function QuitarCaracteresEspecialesYNumeros(cadena As String) As String
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
    QuitarCaracteresEspecialesYNumeros = CadenaFinal
End Function

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
            If (Not (Asc(Caracter) >= 65 And Asc(Caracter) <= 90)) And IsNumeric(Caracter) = False Then
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
    dtpFechaAl.Value = HoyDia
    On Error GoTo errhandler
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    'Consulta las listas de precios que estan disponibles
        
        strSql = " SELECT tip_ped_codigo " & _
                 " FROM usuario_negocio " & _
                 " WHERE usu_codigo='" & strUsuario & "'"
        
        clsCon_Def.Ejecutar (strSql)
        
        strSql = " SELECT 0 as sel,tipo_pedido.tip_ped_codigo,tip_ped_nombre " & _
                 " FROM tipo_pedido "
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            strSql = strSql & " WHERE tipo_pedido.tip_ped_codigo IN (" & _
                     " SELECT tip_ped_codigo " & _
                     " FROM usuario_negocio " & _
                     " WHERE usu_codigo='" & strUsuario & "') "
        End If
        strSql = strSql & " ORDER BY tip_ped_nombre"
        
        clsCon_Def.Ejecutar strSql
        Set Me.VSFGNegocio.DataSource = clsCon_Def.adorec_Def.DataSource
        
        strSql = " SELECT banco_cash.ban_codigo, ban_nombre " & _
                 " FROM banco_cash INNER JOIN banco ON banco_cash.ban_codigo=banco.ban_codigo " & _
                 " WHERE banco_cash.ban_cas_tipo='C' " & _
                 " ORDER BY ban_nombre "
        clsCon_Def.Ejecutar strSql
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

Private Sub VSFGNegocio_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Or Row <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFGNegocio_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        VSFGNegocio.Cell(flexcpBackColor, Row, 0, Row, VSFGNegocio.Cols - 1) = IIf(Abs(VSFGNegocio.TextMatrix(Row, 0)) = 0, vbWhite, vbYellow)
    End If
End Sub
