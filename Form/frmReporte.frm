VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReporte 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DDDDDD&
   Caption         =   "Reporte"
   ClientHeight    =   9195
   ClientLeft      =   360
   ClientTop       =   960
   ClientWidth     =   9975
   Icon            =   "frmReporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   9975
   Begin VB.PictureBox CBFactura 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   8160
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportHTML 
      Caption         =   "Guardar como HTML"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   8640
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExportPDF 
      Caption         =   "Guardar como PDF"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportRTF 
      Caption         =   "Guardar como RTF"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrint 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _cx             =   13785
      _cy             =   16113
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   16777215
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
      AbortCaption    =   "Imprimiendo..."
      AbortTextButton =   "Cancelar"
      AbortTextDevice =   "en la %s en %s"
      AbortTextPage   =   "Ahora Imprimiendo página %d de"
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
      Zoom            =   61.1601513240858
      ZoomMode        =   4
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   14540253
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
      NavBarColor     =   14540253
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Página Completa|Ancha de Página|Dos Páginas|Todas las Páginas"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport VSSubRpt3 
      Left            =   9240
      Top             =   1680
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VSReport8LibCtl.VSReport VSSubRpt2 
      Left            =   8880
      Top             =   1680
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VSReport8LibCtl.VSReport VSSubRpt 
      Left            =   8520
      Top             =   1680
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VSReport8LibCtl.VSReport VSRpt 
      Left            =   8160
      Top             =   1680
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strReporte As String
Public strAsiento As String
Public strNumero As String
Public strTipo As String
Public strSql As String
Private clsConAux1 As New clsConsulta
Private clsSql1 As New clsConsulta
Private clsSQL2 As New clsConsulta
Public Atencion As String

Private Sub cmdExportPDF_Click()
    cdArchivo.DefaultExt = "PDF"
    cdArchivo.Filter = "Archivos PDF (*.pdf)|*.pdf|Todos |*.*"
    cdArchivo.FileName = ""
    cdArchivo.FilterIndex = 1
    cdArchivo.ShowSave
    If cdArchivo.FileName <> "" Then
        VSRpt.RenderToFile cdArchivo.FileName, vsrPDF
    End If
End Sub

Private Sub cmdExportRTF_Click()
    cdArchivo.DefaultExt = "RTF"
    cdArchivo.Filter = "Archivos RTF (*.rtf)|*.rtf|Todos |*.*"
    cdArchivo.FileName = ""
    cdArchivo.FilterIndex = 1
    cdArchivo.ShowSave
    If cdArchivo.FileName <> "" Then
        VSRpt.RenderToFile cdArchivo.FileName, vsrRTF
    End If
End Sub

Private Sub cmdExportHTML_Click()
    cdArchivo.DefaultExt = "HTML"
    cdArchivo.Filter = "Archivos HTML (*.html;*.htm)|*.htm;*.html|Todos |*.*"
    cdArchivo.FileName = ""
    cdArchivo.FilterIndex = 1
    cdArchivo.ShowSave
    If cdArchivo.FileName <> "" Then
        VSRpt.RenderToFile cdArchivo.FileName, vsrHTML
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Public Sub Form_Activate()
    If strReporte <> "" Then
        If strReporte = "rptRetencionDiario" Then
            strReporteAux = strReporte
            strReporte = "rptAsiento"
            
            strAux = GetSQL
            strReporte = strReporteAux
        Else
            strAux = GetSQL
        End If
Imprimir001:
    On Error GoTo errhandler
        If errr = 0 Then
            VSRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
        Else
            VSRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
        End If
        VSRpt.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
        VSRpt.DataSource.RecordSource = strAux
        VSRpt.DataSource.GetRecordSource True
        If strReporte = "rptRetencionDiario" Then
            strReporteAux = strReporte
            strReporte = "rptRetencion"
            strAux = GetSQL
            If errr = 0 Then
                VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
            Else
                VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
            End If
            
            strReporte = strReporteAux
            If Atencion <> "" Then
                VSRpt.Sections("Footer").Fields("lblFechaPago").Text = Atencion
            End If
            VSRpt.Sections("Footer").Fields("rptRetencion").Subreport = VSSubRpt
            VSSubRpt.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
            If strPuertoLocal <> "" Then
                VSSubRpt.DataSource.ConnectionString = VSSubRpt.DataSource.ConnectionString & ", " & strPuertoLocal
            End If
            VSSubRpt.DataSource.RecordSource = strAux
            VSSubRpt.DataSource.GetRecordSource True
        
        ElseIf strReporte = "rptFacturaGuia" Then
            
            strReporteAux = strReporte
            strReporte = "rptFactura"
'            If errr = 0 Then
'                VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
'            Else
                VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
'            End If
            
            If strPtoFactura = "002" Then
            
                strReporte = "rptFacIncentivo"
                strAux = GetSQL
'                If errr = 0 Then
                    VSSubRpt2.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
'                Else
'                    VSSubRpt2.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
'                End If
                
                strReporte = "rptFactura"
                VSSubRpt.Sections("GFooter").Fields("SubRpt").Subreport = VSSubRpt2
                VSSubRpt2.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                               & "Persist Security Info=False;" _
                               & "User ID=" & strUsuario & ";" _
                               & "Password=" & strClave & ";" _
                               & "Initial Catalog=" & strBDD & ";" _
                               & "Data Source=" & strServidorBDDLocal
                If strPuertoLocal <> "" Then
                    VSSubRpt2.DataSource.ConnectionString = VSSubRpt2.DataSource.ConnectionString & ", " & strPuertoLocal
                End If
                VSSubRpt2.DataSource.RecordSource = strAux
                VSSubRpt2.DataSource.GetRecordSource True
            
            End If
            
            strAux = GetSQL
            strReporte = strReporteAux
            VSRpt.Sections("Detail").Fields("SubRpt").Subreport = VSSubRpt
            VSSubRpt.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
                If strPuertoLocal <> "" Then
                    VSSubRpt.DataSource.ConnectionString = VSSubRpt.DataSource.ConnectionString & ", " & strPuertoLocal
                End If
            VSSubRpt.DataSource.RecordSource = strAux
            VSSubRpt.DataSource.GetRecordSource True
'
'
            
            strReporteAux = strReporte
            strReporte = "rptGuiaFactura"
            strAux = GetSQL
'            If errr = 0 Then
                VSSubRpt3.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
'            Else
'                VSSubRpt3.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
'            End If

            strReporte = strReporteAux
            VSRpt.Sections("Detail").Fields("SubRpt1").Subreport = VSSubRpt3
            VSSubRpt3.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
            If strPuertoLocal <> "" Then
                VSSubRpt3.DataSource.ConnectionString = VSSubRpt3.DataSource.ConnectionString & ", " & strPuertoLocal
            End If
            VSSubRpt3.DataSource.RecordSource = strAux
            VSSubRpt3.DataSource.GetRecordSource True
        
        
        ElseIf strReporte = "rptFacturaSola" Then
            
            strReporteAux = strReporte
            strReporte = "rptFactura"
            'If errr = 0 Then
            '    VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
            'Else
                VSSubRpt.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
            'End If
            
            If strPtoFactura = "002" Then
            
                strReporte = "rptFacIncentivo"
                strAux = GetSQL
                If errr = 0 Then
                    VSSubRpt2.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
                Else
                    VSSubRpt2.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
                End If
                
                strReporte = "rptFactura"
                VSSubRpt.Sections("GFooter").Fields("SubRpt").Subreport = VSSubRpt2
                VSSubRpt2.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                               & "Persist Security Info=False;" _
                               & "User ID=" & strUsuario & ";" _
                               & "Password=" & strClave & ";" _
                               & "Initial Catalog=" & strBDD & ";" _
                               & "Data Source=" & strServidorBDDLocal
                If strPuertoLocal <> "" Then
                    VSSubRpt2.DataSource.ConnectionString = VSSubRpt2.DataSource.ConnectionString & ", " & strPuertoLocal
                End If
                VSSubRpt2.DataSource.RecordSource = strAux
                VSSubRpt2.DataSource.GetRecordSource True
                
                
            
'                strReporte = "rptFacReprogramacion"
'                strAux = GetSQL
'                If errr = 0 Then
'                    VSSubRpt3.Load App.Path & "\VSReport\reportes.xml", strReporte & strEmpresa
'                Else
'                    VSSubRpt3.Load App.Path & "\VSReport\reportes.xml", strReporte & strPtoFactura & strEmpresa
'                End If
'
'                strReporte = "rptFactura"
'                VSSubRpt.Sections("GFooter").Fields("SubRpt2").Subreport = VSSubRpt3
'                VSSubRpt3.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
'                               & "Persist Security Info=False;" _
'                               & "User ID=" & strUsuario & ";" _
'                               & "Password=" & strClave & ";" _
'                               & "Initial Catalog=" & strBDD & ";" _
'                               & "Data Source=" & strServidorBDDLocal & ", " & strPuertoLocal
'                VSSubRpt3.DataSource.RecordSource = strAux
'                VSSubRpt3.DataSource.GetRecordSource True
            
            End If
            
            strAux = GetSQL
            strReporte = strReporteAux
            VSRpt.Sections("Detail").Fields("SubRpt").Subreport = VSSubRpt
            VSSubRpt.DataSource.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
            If strPuertoLocal <> "" Then
                VSSubRpt.DataSource.ConnectionString = VSSubRpt.DataSource.ConnectionString & ", " & strPuertoLocal
            End If
            VSSubRpt.DataSource.RecordSource = strAux
            VSSubRpt.DataSource.GetRecordSource True
            
        End If
        VSRpt.Render VSPrint
    End If
    Exit Sub
errhandler:
    Select Case Err.Number
        Case 1002
            If errr = 0 Then
                errr = 1
                    GoTo Imprimir001
            End If
        Case Else
            
                MsgBox "[" & Err.Number & "] " & Err.Description
    
    End Select
    
    
    'End If
End Sub

Private Sub Form_Load()
    If strReporte = "rptIDCaja" Then
        clsConAux1.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT for_ent_nombre ,per_direccion " & _
                 " FROM persona INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo" & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "'" & _
                 " AND per_codigo='" & strNumero & "'"
        clsConAux1.Ejecutar strSql
        strAsiento = UCase(InputBox("Forma de entrega: " & clsConAux1.adorec_Def("for_ent_nombre") & vbNewLine & vbNewLine & vbNewLine & "Dirección de Entrega:", "Empaque", clsConAux1.adorec_Def("per_direccion")))
        Set clsConAux1 = Nothing
    ElseIf strReporte = "rptRetencion" Then
        'strAsiento = UCase(InputBox("Concepto de Retención", "Retención", "ADQUISICIÓN DE MERCADERÍAS"))
    End If
End Sub

Private Sub Form_Resize()
    If Me.Width - 2295 >= 0 Then
        VSPrint.Width = Me.Width - 2295
    End If
    If Me.Height - 525 >= 0 Then
        VSPrint.Height = Me.Height - 525
    End If
    If Me.Width - 1950 >= 0 Then
        cmdExportHTML.Left = Me.Width - 1950
        cmdExportPDF.Left = Me.Width - 1950
        cmdExportRTF.Left = Me.Width - 1950
        cmdSalir.Left = Me.Width - 1950
    End If
End Sub

Public Function GetSQL() As String
    Dim strSqlAux As String
    Dim clsConAUX As New clsConsulta
    Dim clsConAux2 As New clsConsulta
    Dim tNum2Text As New cNum2Text
    Dim lngValor As Long
    Dim intValor As Integer
    Dim strValor As String
    Dim strFecha As String
    Dim CDEClaveAcceso As String
    Dim Detalle As String
    Dim CBF As String
    Dim CBP As String
    Dim CBG As String
    Dim jj As Long
    Dim canti As String
    Dim dirEnvio As String
    
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    If strReporte = "rptAsiento" Then
        Me.Caption = "Asiento Contable - " & strAsiento
        GetSQL = " SELECT asiento.asi_numasiento,asi_fecha,CONCAT(LEFT(asi_descripcion,2000),IIF(LEN(asi_descripcion)>2000,'...','')) as asi_descripcion,asi_totaldebe,asi_totalhaber,asi_usumod, " & _
                 " det_asiento.cta_codigo,cta_nombre,det_asi_debe,det_asi_haber,COALESCE(cen_cos_nombre,'') as cen_cos_nombre,emp_nombre " & _
                 " FROM (asiento INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento) " & _
                 " INNER JOIN ctaconta ON det_asiento.cta_codigo=ctaconta.cta_codigo AND det_asiento.emp_codigo=ctaconta.emp_codigo " & _
                 " INNER JOIN empresa ON asiento.emp_codigo=empresa.emp_codigo " & _
                 " LEFT JOIN centro_costo ON det_asiento.cen_cos_codigo=centro_costo.cen_cos_codigo AND det_asiento.emp_codigo=centro_costo.emp_codigo " & _
                 " WHERE asiento.emp_codigo='" & strEmpresa & "' " & _
                 " AND asiento.asi_numasiento IN ('" & strAsiento & "') "
                 
    ElseIf strReporte = "rptOrdenCompraTalla" Then
        Me.Caption = "Order de Compra - " & strNumero
        GetSQL = " EXEC Sp_Rpt_Orden_Compra '" & strEmpresa & "'," & strNumero & ",'P'"
                 
    ElseIf strReporte = "rptOrdenCompraSumi" Then
        Me.Caption = "Order de Compra - " & strNumero
        GetSQL = " EXEC Sp_Rpt_Orden_Compra '" & strEmpresa & "'," & strNumero & ",'S'"
                 
    ElseIf strReporte = "rptAjuste" Then
        Me.Caption = "Ajuste - " & strNumero
        GetSQL = " SELECT emp_nombre,CONCAT(ven_apellido,' ',ven_nombre) as nombV," & _
                 " CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')) as NN1, " & _
                 " CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')) as NN2, " & _
                 " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as NN3, " & _
                 " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as NN4, " & _
                 " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5, " & _
                 " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6, " & _
                 " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7, " & _
                 " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8, " & _
                 " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9, " & _
                 " CONCAT(persona.per_apellido,' ', persona.per_nombre,' (',persona.tip_ped_codigo,') ' ) as nombC, " & _
                 " cam_observacion as obs , cam_fecha as fech, persona.per_telf, cambio.cam_codigo,cam_usumod as usumod,cam_fechamod as fechamod, " & _
                 " prd_codigo_ing,ping.prd_nombre as prd_nombre_ing,prd_codigo_ped,pped.prd_nombre as prd_nombre_ped,det_cam_cantidad " & _
                 " FROM empresa INNER JOIN persona ON empresa.emp_codigo=persona.emp_codigo INNER JOIN cambio ON persona.emp_codigo = cambio.emp_codigo " & _
                 " AND persona.per_codigo = cambio.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " INNER JOIN det_cambio ON cambio.emp_codigo = det_cambio.emp_codigo " & _
                 " AND cambio.cam_codigo = det_cambio.cam_codigo " & _
                 " INNER JOIN producto ping ON det_cambio.emp_codigo=ping.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ing=ping.prd_codigo" & _
                 " INNER JOIN producto pped ON det_cambio.emp_codigo=pped.emp_codigo" & _
                 " AND det_cambio.prd_codigo_ped=pped.prd_codigo" & _
                 " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p1.per_es_di=1 " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE cambio.cam_codigo = '" & strNumero & "' " & _
                 " AND empresa.emp_codigo='" & strEmpresa & "' ORDER BY prd_codigo_ing,prd_codigo_ped"
                 
    ElseIf strReporte = "rptTckAjuste" Then
        Me.Caption = "Comprobante de Ajuste - " & strNumero
        GetSQL = " SELECT CONCAT(ven_apellido,' ',ven_nombre) as vendedor," & _
                 " CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')) as NN1, " & _
                 " CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')) as NN2, " & _
                 " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as NN3, " & _
                 " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as NN4, " & _
                 " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5, " & _
                 " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6, " & _
                 " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7, " & _
                 " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8, " & _
                 " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9, " & _
                 " CONCAT(persona.per_apellido,' ', persona.per_nombre,' (',persona.tip_ped_codigo,') ' ) as per, " & _
                 " persona.per_telf, cambio.cam_codigo,cam_usumod as usumod,cam_fechamod as fechamod " & _
                 " FROM persona INNER JOIN cambio ON persona.emp_codigo = cambio.emp_codigo " & _
                 " AND persona.per_codigo = cambio.per_codigo " & _
                 " AND persona.cat_p_tipo='C' " & _
                 " INNER JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p1.per_es_di=1 " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE cambio.cam_codigo = '" & strNumero & "' " & _
                 " AND persona.emp_codigo='" & strEmpresa & "'"
    ElseIf strReporte = "rptComprobanteIngreso" Then
        Me.Caption = "Comprobante de Ingreso - " & strAsiento
        GetSQL = " SELECT asiento.asi_numasiento,asi_fecha,asi_descripcion,asi_totaldebe,asi_totalhaber,asi_usumod, " & _
                 " det_asiento.cta_codigo,cta_nombre,det_asi_debe,det_asi_haber,COALESCE(cen_cos_nombre,'') as cen_cos_nombre,emp_nombre " & _
                 " FROM (asiento INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento) " & _
                 " INNER JOIN ctaconta ON det_asiento.cta_codigo=ctaconta.cta_codigo AND det_asiento.emp_codigo=ctaconta.emp_codigo " & _
                 " INNER JOIN empresa ON asiento.emp_codigo=empresa.emp_codigo " & _
                 " LEFT JOIN centro_costo ON det_asiento.cen_cos_codigo=centro_costo.cen_cos_codigo AND det_asiento.emp_codigo=centro_costo.emp_codigo " & _
                 " WHERE asiento.emp_codigo='" & strEmpresa & "' " & _
                 " AND asiento.asi_numasiento='" & strAsiento & "' "
    ElseIf strReporte = "rptIDCaja" Then
        Me.Caption = "Identificación Caja - " & strNumero
        GetSQL = " SELECT LEFT(CURRENT_TIMESTAMP,10) as fecha,for_ent_nombre,CONCAT(per_apellido,' ',per_nombre) as per,'" & strAsiento & "' as per_direccion,per_direccion2,per_telf,ciu_nombre " & _
                 " FROM persona INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                 " AND persona.per_codigo='" & strNumero & "' "
    ElseIf strReporte = "rptSTKIdCaja" Then
        Me.Caption = "Identificación Caja - " & strNumero
        GetSQL = " SELECT emp_nombre,LEFT(CURRENT_TIMESTAMP,10) as fecha,for_ent_nombre," & _
                 " CONCAT(per_apellido,' ',per_nombre) as per,per_direccion," & _
                 " per_direccion2,CONCAT(per_telf,'/',per_fax,'/',per_celular) as per_telf,ciu_nombre " & _
                 " FROM empresa INNER JOIN persona ON empresa.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND persona.per_codigo='" & strNumero & "' "
                 
    ElseIf strReporte = "rptSTKUbicacion" Or strReporte = "rptSTKUbicacionMini" Then
        Me.Caption = "Ubicacion Mercaderia - " & strTipo & " - " & strNumero
        ReDim ped(UBound(Split(strNumero, ","))) As String
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBUbica" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBUbica" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " dep_codigo char(3) NOT NULL default ''," & _
                   " ubi_bod_codigo varchar(6) NOT NULL default '0', " & _
                   " ubi_bod_cdb varchar(50) default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,dep_codigo,ubi_bod_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To UBound(ped)
            strSqlAux = " INSERT INTO CDBUbica" & strUsuario & "(emp_codigo,dep_codigo,ubi_bod_codigo,ubi_bod_cdb) " & _
                        " VALUES('" & strEmpresa & "','" & strTipo & "'," & ped(jj) & "," & _
                        " '" & Replace(code128$(Replace(ped(jj), "'", "")), "'", "''") & "') "
            clsConAUX.Ejecutar strSqlAux
        Next jj
        
        
        GetSQL = " SELECT dep_nombre,ubicacion_bodega.ubi_bod_codigo," & _
                 " ubi_bod_cdb" & _
                 " FROM ubicacion_bodega INNER JOIN deposito " & _
                 " ON ubicacion_bodega.emp_codigo=deposito.emp_codigo " & _
                 " AND ubicacion_bodega.dep_codigo=deposito.dep_codigo " & _
                 " INNER JOIN CDBUbica" & strUsuario & " " & _
                 " ON ubicacion_bodega.emp_codigo=CDBUbica" & strUsuario & ".emp_codigo " & _
                 " AND ubicacion_bodega.dep_codigo=CDBUbica" & strUsuario & ".dep_codigo " & _
                 " AND ubicacion_bodega.ubi_bod_codigo=CDBUbica" & strUsuario & ".ubi_bod_codigo " & _
                 " WHERE ubicacion_bodega.emp_codigo='" & strEmpresa & "'" & _
                 " AND ubicacion_bodega.ubi_bod_codigo IN (" & strNumero & ")"
    ElseIf strReporte = "rptSTKContenedorMercaderia" Then
        Me.Caption = "Contenedor Mercaderia - " & strNumero
        GetSQL = " SELECT con_mer_codigo,con_mer_observacion,con_mer_fecha,tip_con_nombre,tip_mer_con_nombre,con_mer_usumod,CURRENT_TIMESTAMP as HOY," & _
                 " '" & Replace(code128$(strNumero), "'", "''") & "' AS GBC" & _
                 " FROM contenedor_mercaderia INNER JOIN tipo_contenedor " & _
                 " ON contenedor_mercaderia.emp_codigo=tipo_contenedor.emp_codigo " & _
                 " AND contenedor_mercaderia.tip_con_codigo=tipo_contenedor.tip_con_codigo " & _
                 " INNER JOIN tipo_mercaderia_contenedor " & _
                 " ON contenedor_mercaderia.emp_codigo=tipo_mercaderia_contenedor.emp_codigo " & _
                 " AND contenedor_mercaderia.tip_mer_con_codigo=tipo_mercaderia_contenedor.tip_mer_con_codigo " & _
                 " WHERE contenedor_mercaderia.emp_codigo='" & strEmpresa & "'" & _
                 " AND con_mer_codigo='" & strNumero & "'"
    ElseIf strReporte = "rptContenedorMercaderia" Then
        Me.Caption = "Contenedor Mercaderia - " & strNumero
        GetSQL = " SELECT contenedor_mercaderia.con_mer_codigo,con_mer_observacion,con_mer_fecha,tip_con_nombre,tip_mer_con_nombre,con_mer_usumod,con_mer_fechamod," & _
                 " '" & Replace(code128$(strNumero), "'", "''") & "' AS GBC," & _
                 " producto.prd_codigo,prd_nombre,SUM(IIF(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot" & _
                 " FROM contenedor_mercaderia INNER JOIN tipo_contenedor " & _
                 " ON contenedor_mercaderia.emp_codigo=tipo_contenedor.emp_codigo " & _
                 " AND contenedor_mercaderia.tip_con_codigo=tipo_contenedor.tip_con_codigo" & _
                 " INNER JOIN tipo_mercaderia_contenedor " & _
                 " ON contenedor_mercaderia.emp_codigo=tipo_mercaderia_contenedor.emp_codigo " & _
                 " AND contenedor_mercaderia.tip_mer_con_codigo=tipo_mercaderia_contenedor.tip_mer_con_codigo " & _
                 " INNER JOIN det_contenedor_mercaderia ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN producto ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo" & _
                 " WHERE contenedor_mercaderia.emp_codigo='" & strEmpresa & "'" & _
                 " AND contenedor_mercaderia.con_mer_codigo='" & strNumero & "'" & _
                 " GROUP BY contenedor_mercaderia.con_mer_codigo,con_mer_observacion,con_mer_fecha,tip_con_nombre,tip_mer_con_nombre,con_mer_usumod,con_mer_fechamod, producto.prd_codigo,prd_nombre " & _
                 " ORDER BY tot DESC,prd_nombre "
    ElseIf strReporte = "rptSTKGuia" Then
        Me.Caption = "Guia - " & strNumero
        
        strSqlAux = " SELECT DISTINCT '1' as n,det_contenedor.emp_codigo,det_contenedor.con_codigo,CAST(pedido.ped_codigo as varchar) as obs,ped_direccion_envio,pedido.per_codigo " & _
                 " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND det_contenedor.tip_egr_codigo=pedido.ped_tip_egr_codigo " & _
                 " AND pedido.ped_estado in (2,10) " & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & strNumero & "' " & _
                 " UNION " & _
                 " SELECT '2' as n,det_contenedor_per.emp_codigo,det_contenedor_per.con_codigo,det_contenedor_per.det_con_per_detalle,'','' " & _
                 " FROM det_contenedor_per " & _
                 " WHERE det_contenedor_per.emp_codigo='" & strEmpresa & "' AND det_contenedor_per.con_codigo='" & strNumero & "' " & _
                 " ORDER BY n,per_codigo"
        clsConAUX.Ejecutar strSqlAux
        Detalle = ""
        jj = 0
        While Not clsConAUX.adorec_Def.EOF
            Detalle = Detalle & clsConAUX.adorec_Def("obs") & ","
            jj = jj + 1
            clsConAUX.adorec_Def.MoveNext
        Wend
        strSqlAux = " SELECT con_guia " & _
                    " FROM contenedor " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND con_codigo='" & strNumero & "' "
        clsConAUX.Ejecutar strSqlAux
        
        
        GetSQL = " SELECT empresa.emp_nombre,det_con_caj_codigo,paq_env_nombre,ncajas.tc,det_con_caj_peso, contenedor.con_codigo, con_fecha, con_guia,con_peso, " & _
                 " '" & Replace(code128$(clsConAUX.adorec_Def("con_guia")), "'", "''") & "' AS GCB, con_peso, cou_nombre, '" & jj & "' AS numped, " & _
                 " CONCAT(per_apellido,' ',per_nombre) as per,RTRIM(SUBSTRING(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio),CHARINDEX(' - ',IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio))+3,500)) as per_direccion, " & _
                 " RTRIM(LEFT(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio),CHARINDEX(' - ',IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio)))) as per_direccion1,per_direccion2,CONCAT(per_telf,'/',per_fax,'/',per_celular) as per_telf,ciu_nombre,can_nombre, " & _
                 " CAST(REPLACE('" & Detalle & "',',',', ') AS varchar) AS detalle " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN persona ON empresa.emp_codigo=persona.emp_codigo " & _
                 " AND contenedor.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo " & _
                 " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo "
        GetSQL = GetSQL & " INNER JOIN det_contenedor_caja ON contenedor.emp_codigo=det_contenedor_caja.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_caja.con_codigo " & _
                 " INNER JOIN paquete_envio ON det_contenedor_caja.emp_codigo=paquete_envio.emp_codigo " & _
                 " AND det_contenedor_caja.paq_env_codigo=paquete_envio.paq_env_codigo " & _
                 " INNER JOIN (SELECT emp_codigo,con_codigo,count(*) as tc FROM det_contenedor_caja " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND con_codigo='" & strNumero & "' " & _
                 " GROUP BY emp_codigo,con_codigo) as ncajas " & _
                 " ON contenedor.emp_codigo=ncajas.emp_codigo" & _
                 " and contenedor.con_codigo=ncajas.con_codigo"
        GetSQL = GetSQL & " LEFT JOIN (SELECT TOP 1 pedido.emp_codigo,pedido.per_codigo,pedido.ped_direccion_envio " & _
                 " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND det_contenedor.tip_egr_codigo=pedido.ped_tip_egr_codigo " & _
                 " AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.for_pag_codigo IN ('CONT','EFE')" & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & strNumero & "' ) as d" & _
                 " ON persona.emp_codigo=d.emp_codigo " & _
                 " AND persona.per_codigo=d.per_codigo "
        GetSQL = GetSQL & " WHERE empresa.emp_codigo='" & strEmpresa & "' AND contenedor.con_codigo='" & strNumero & "' "
    ElseIf strReporte = "rptSTKListaEmbarque" Then
        Me.Caption = "Etiqueta Lista de Empaque - " & strNumero
        
        strSqlAux = " SELECT con_guia " & _
                    " FROM contenedor " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND con_codigo='" & strNumero & "' "
        clsConAUX.Ejecutar strSqlAux
        
        GetSQL = " SELECT empresa.emp_nombre, contenedor.con_codigo, con_fecha, " & _
                 " con_guia, '" & Replace(code128$(clsConAUX.adorec_Def("con_guia")), "'", "''") & "' as GCB," & _
                 " con_peso, cou_nombre, pedido.ped_codigo, COUNT(pedido.ped_codigo) as numped, " & _
                 " CAST(REPLACE(GROUP_CONCAT(pedido.ped_codigo),',',', ') AS CHAR(5000)) as detalle " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo " & _
                 " INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo='FAC' AND pedido.ped_estado in (2,10) " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' " & _
                 " GROUP BY empresa.emp_codigo,contenedor.con_codigo "
        GetSQL = GetSQL & " UNION SELECT empresa.emp_nombre, contenedor.con_codigo,con_fecha, " & _
                 " con_guia, '" & Replace(code128$(clsConAUX.adorec_Def("con_guia")), "'", "''") & "' as GCB," & _
                 " con_peso,cou_nombre,det_contenedor_per.det_con_per_detalle,COUNT(det_contenedor_per.det_con_per_detalle) as numped," & _
                 " CAST(REPLACE(GROUP_CONCAT(det_contenedor_per.det_con_per_detalle),',',', ') AS CHAR(5000)) as detalle " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_contenedor_per ON contenedor.emp_codigo=det_contenedor_per.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_per.con_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' " & _
                 " GROUP BY empresa.emp_codigo,contenedor.con_codigo "
    ElseIf strReporte = "rptManifiestoCarga" Or strReporte = "rptManifiestoEntregas" Then
        Me.Caption = "Manifiesto de Carga - " & strNumero
        
        
        strSqlAux = " SELECT COUNT(DISTINCT contenedor.con_guia) as nguia " & _
                    " FROM manifiesto_carga INNER JOIN det_manifiesto_carga ON manifiesto_carga.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                    " AND manifiesto_carga.man_car_codigo=det_manifiesto_carga.man_car_codigo " & _
                    " INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                    " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                    " WHERE manifiesto_carga.emp_codigo='" & strEmpresa & "' " & _
                    " AND manifiesto_carga.man_car_codigo='" & strNumero & "' " & _
                    " GROUP BY manifiesto_carga.emp_codigo "
        clsConAUX.Ejecutar strSqlAux
        
        GetSQL = " SELECT empresa.emp_nombre, manifiesto_carga.man_car_codigo,man_car_fecha,man_car_observacion," & _
                 " courier.cou_nombre, man_car_placa,man_car_responsable, " & _
                 " det_manifiesto_carga.con_codigo,contenedor.con_guia,CONCAT(per_apellido,' ',per_nombre) as cli, " & _
                 " det_manifiesto_carga.paq_env_codigo,paq_env_nombre,'" & clsConAUX.adorec_Def("nguia") & "' as nguia,COUNT(det_man_car_codigo) as cantidad,man_car_usumod " & _
                 " FROM empresa INNER JOIN manifiesto_carga ON empresa.emp_codigo=manifiesto_carga.emp_codigo " & _
                 " INNER JOIN courier ON manifiesto_carga.emp_codigo=courier.emp_codigo " & _
                 " AND manifiesto_carga.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_manifiesto_carga ON manifiesto_carga.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                 " AND manifiesto_carga.man_car_codigo=det_manifiesto_carga.man_car_codigo " & _
                 " INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                 " INNER JOIN paquete_envio ON det_manifiesto_carga.emp_codigo=paquete_envio.emp_codigo " & _
                 " AND det_manifiesto_carga.paq_env_codigo=paquete_envio.paq_env_codigo " & _
                 " INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo " & _
                 " AND contenedor.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C'" & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND manifiesto_carga.man_car_codigo='" & strNumero & "' " & _
                 " GROUP BY empresa.emp_nombre, manifiesto_carga.man_car_codigo,man_car_fecha,man_car_observacion, courier.cou_nombre, man_car_placa,man_car_responsable, det_manifiesto_carga.con_codigo,contenedor.con_guia, per_apellido,per_nombre, det_manifiesto_carga.paq_env_codigo,paq_env_nombre,man_car_usumod" & _
                 " ORDER BY det_manifiesto_carga.paq_env_codigo "
    ElseIf strReporte = "rptManifiestoEntregasCliente" Then
        Me.Caption = "Manifiesto de Carga - " & strNumero
        
        
        strSqlAux = " SELECT COUNT(DISTINCT contenedor.con_guia) as nguia " & _
                    " FROM manifiesto_carga INNER JOIN det_manifiesto_carga ON manifiesto_carga.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                    " AND manifiesto_carga.man_car_codigo=det_manifiesto_carga.man_car_codigo " & _
                    " INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                    " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                    " WHERE manifiesto_carga.emp_codigo='" & strEmpresa & "' " & _
                    " AND manifiesto_carga.man_car_codigo='" & strNumero & "' " & _
                    " GROUP BY manifiesto_carga.emp_codigo "
        clsConAUX.Ejecutar strSqlAux
        
        GetSQL = " SELECT empresa.emp_nombre, manifiesto_carga.man_car_codigo,man_car_fecha,man_car_observacion," & _
                 " courier.cou_nombre, man_car_placa,man_car_responsable, " & _
                 " det_manifiesto_carga.con_codigo,contenedor.con_guia," & _
                 " CONCAT(persona.per_apellido,' ',persona.per_nombre) as cli, " & _
                 " CONCAT(pl.per_apellido,' ',pl.per_nombre) as clipl, " & _
                 " det_manifiesto_carga.paq_env_codigo,paq_env_nombre," & _
                 " man_car_usumod,det_contenedor.egr_codigo,'" & clsConAUX.adorec_Def("nguia") & "' as nguia," & _
                 " COUNT(det_man_car_codigo) as cantidad " & _
                 " FROM empresa INNER JOIN manifiesto_carga ON empresa.emp_codigo=manifiesto_carga.emp_codigo " & _
                 " INNER JOIN courier ON manifiesto_carga.emp_codigo=courier.emp_codigo " & _
                 " AND manifiesto_carga.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_manifiesto_carga ON manifiesto_carga.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                 " AND manifiesto_carga.man_car_codigo=det_manifiesto_carga.man_car_codigo " & _
                 " INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                 " INNER JOIN persona pl ON contenedor.emp_codigo=pl.emp_codigo " & _
                 " AND contenedor.per_codigo=pl.per_codigo AND pl.cat_p_tipo='C'"
        GetSQL = GetSQL & " INNER JOIN paquete_envio ON det_manifiesto_carga.emp_codigo=paquete_envio.emp_codigo " & _
                 " AND det_manifiesto_carga.paq_env_codigo=paquete_envio.paq_env_codigo " & _
                 " INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo " & _
                 " INNER JOIN egreso ON det_contenedor.emp_codigo=egreso.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_anulado=0 " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo " & _
                 " AND egreso.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C'" & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND manifiesto_carga.man_car_codigo='" & strNumero & "' " & _
                 " GROUP BY det_manifiesto_carga.con_codigo,det_manifiesto_carga.paq_env_codigo,det_contenedor.egr_codigo"
        GetSQL = GetSQL & " UNION " & _
                 " SELECT empresa.emp_nombre, manifiesto_carga.man_car_codigo,man_car_fecha,man_car_observacion," & _
                 " courier.cou_nombre, man_car_placa,man_car_responsable, " & _
                 " det_manifiesto_carga.con_codigo,contenedor.con_guia," & _
                 " CONCAT(persona.per_apellido,' ',persona.per_nombre) as cli, " & _
                 " CONCAT(pl.per_apellido,' ',pl.per_nombre) as clipl, " & _
                 " det_manifiesto_carga.paq_env_codigo,paq_env_nombre," & _
                 " man_car_usumod,det_con_per_detalle,'" & clsConAUX.adorec_Def("nguia") & "' as nguia," & _
                 " COUNT(det_man_car_codigo) as cantidad " & _
                 " FROM empresa INNER JOIN manifiesto_carga ON empresa.emp_codigo=manifiesto_carga.emp_codigo " & _
                 " INNER JOIN courier ON manifiesto_carga.emp_codigo=courier.emp_codigo " & _
                 " AND manifiesto_carga.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_manifiesto_carga ON manifiesto_carga.emp_codigo=det_manifiesto_carga.emp_codigo " & _
                 " AND manifiesto_carga.man_car_codigo=det_manifiesto_carga.man_car_codigo " & _
                 " INNER JOIN contenedor ON det_manifiesto_carga.emp_codigo=contenedor.emp_codigo " & _
                 " AND det_manifiesto_carga.con_codigo=contenedor.con_codigo " & _
                 " INNER JOIN persona pl ON contenedor.emp_codigo=pl.emp_codigo " & _
                 " AND contenedor.per_codigo=pl.per_codigo AND pl.cat_p_tipo='C'"
        GetSQL = GetSQL & " INNER JOIN paquete_envio ON det_manifiesto_carga.emp_codigo=paquete_envio.emp_codigo " & _
                 " AND det_manifiesto_carga.paq_env_codigo=paquete_envio.paq_env_codigo " & _
                 " INNER JOIN det_contenedor_per ON contenedor.emp_codigo=det_contenedor_per.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_per.con_codigo " & _
                 " INNER JOIN persona ON det_contenedor_per.emp_codigo=persona.emp_codigo " & _
                 " AND det_contenedor_per.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C'" & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND manifiesto_carga.man_car_codigo='" & strNumero & "' " & _
                 " GROUP BY det_manifiesto_carga.con_codigo,det_manifiesto_carga.paq_env_codigo,det_con_per_detalle" & _
                 " ORDER BY clipl,cli,con_guia,paq_env_codigo  "
    ElseIf strReporte = "rptListaEmbarque" Then
        Me.Caption = "Lista de Empaque - " & strNumero
        GetSQL = " SELECT empresa.emp_nombre, contenedor.con_codigo,con_fecha,con_observacion,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " con_guia,con_peso,cou_nombre,CAST(CHARINDEX(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre,' - ',persona.per_direccion2),d.ped_direccion_envio),CHARINDEX(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre,' - ',persona.per_direccion2),d.ped_direccion_envio),' - ')+3,500) as varchar) as per_direccion2,persona.per_telf,IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciudad.ciu_nombre,'/',canton.can_nombre,'/',pais.pai_nombre),LEFT(d.ped_direccion_envio,CHARINDEX(d.ped_direccion_envio,' - '))) as ciu_nombre,zona.zon_nombre, " & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet," & _
                 " CAST(egreso.egr_codigo as varchar) as egr_codigo,pedido.ped_codigo,con_usumod,egr_total,sum(det_egr_cantidad) as egr_unidades " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo AND contenedor.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo " & _
                 " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                 " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo INNER JOIN det_contenedor ON contenedor.emp_codigo=det_contenedor.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor.con_codigo "
        GetSQL = GetSQL & " INNER JOIN egreso ON det_contenedor.emp_codigo=egreso.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_contenedor.tip_egr_codigo AND egreso.egr_anulado=0 " & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND det_egreso.prd_codigo NOT LIKE 'PR-%'" & _
                 " INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona pd ON egreso.emp_codigo=pd.emp_codigo " & _
                 " AND egreso.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo "
        GetSQL = GetSQL & " LEFT JOIN (SELECT TOP 1 pedido.emp_codigo,pedido.per_codigo,pedido.ped_direccion_envio " & _
                 " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
                 " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=det_contenedor.tip_egr_codigo " & _
                 " AND pedido.ped_estado in (2,10) " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
                 " AND pedido.per_codigo=persona.per_codigo " & _
                 " AND persona.for_pag_codigo IN ('CONT','EFE')" & _
                 " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & strNumero & "') as d" & _
                 " ON persona.emp_codigo=d.emp_codigo " & _
                 " AND persona.per_codigo=d.per_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' GROUP BY empresa.emp_nombre,contenedor.con_codigo,con_fecha,con_observacion,persona.per_nombre,persona.per_apellido,persona.per_telf, con_guia,con_peso,cou_nombre,d.ped_direccion_envio,ciudad.ciu_nombre,canton.can_nombre,pais.pai_nombre,persona.per_direccion2, zona.zon_nombre,pd.per_apellido,pd.per_nombre,cd.ciu_nombre,zd.zon_nombre, egreso.egr_codigo,pedido.ped_codigo,con_usumod,egr_total"
        GetSQL = GetSQL & " UNION SELECT empresa.emp_nombre, contenedor.con_codigo,con_fecha,con_observacion,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " con_guia,con_peso,cou_nombre,persona.per_direccion2,persona.per_telf,ciudad.ciu_nombre,zona.zon_nombre, " & _
                 " CONCAT(pd.per_apellido,' ',pd.per_nombre) as perdet,cd.ciu_nombre as ciudet,zd.zon_nombre as zondet," & _
                 " det_contenedor_per.det_con_per_detalle,'0' as ped_codigo,con_usumod,0,1 " & _
                 " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
                 " INNER JOIN persona ON contenedor.emp_codigo=persona.emp_codigo AND contenedor.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN zona ON persona.zon_codigo=zona.zon_codigo " & _
                 " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo " & _
                 " INNER JOIN det_contenedor_per ON contenedor.emp_codigo=det_contenedor_per.emp_codigo " & _
                 " AND contenedor.con_codigo=det_contenedor_per.con_codigo " & _
                 " INNER JOIN persona pd ON det_contenedor_per.emp_codigo=pd.emp_codigo " & _
                 " AND det_contenedor_per.per_codigo=pd.per_codigo " & _
                 " INNER JOIN ciudad cd ON pd.ciu_codigo=cd.ciu_codigo " & _
                 " INNER JOIN zona zd ON pd.zon_codigo=zd.zon_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor.con_codigo='" & strNumero & "' "
    ElseIf strReporte = "rptRolPagos" Then
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        clsConAux2.Inicializar AdoConn, AdoConnMaster
        Me.Caption = "Rol de Pagos " '& strAsiento
        Dim Fecha() As String
        Fecha = Split(strAsiento, ",")
        
            strSql = " CREATE TABLE EstadoCuentaVB " & _
                    " (emp varchar(255), persona varchar(255), nombre varchar(100),nombres varchar(100), " & _
                    " valor numeric(17,2), valores numeric(17,2),tipo int,orden int,tipos int,fecha varchar(20))"
            clsConAux1.Ejecutar (strSql)
            strSql = " INSERT INTO EstadoCuentaVB(emp,persona,nombre,valor,tipo,orden,fecha) " & _
                   " SELECT emp_nombre,concat(epl_apellidos,' ',epl_nombres) AS persona, tip_des_nombre AS producto, " & _
                    " des_valor AS valor,tip_des_ingreso,tip_des_orden AS orden,'" & Fecha(2) & "'" & _
                    " FROM descuento " & _
                    " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                    " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                    " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                    " INNER JOIN empresa ON empresa.emp_codigo=empleado.emp_codigo " & _
                    " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & strNumero & "'" & _
                    " AND des_fecha BETWEEN '" & Fecha(0) & "' AND '" & Fecha(1) & "' AND tipo_descuento.cta_codigo='" & strTipo & "' AND tipo_descuento.tip_des_ingreso=1 AND des_pagado=0 " & _
                    " ORDER BY tip_des_ingreso,tip_des_orden "
              clsConAux1.Ejecutar (strSql)
          strSql = " INSERT INTO EstadoCuentaVB(emp,persona,nombre,valor,tipo,orden,fecha) " & _
                   " SELECT emp_nombre,concat(epl_apellidos,' ',epl_nombres) AS persona, tip_des_nombre AS producto, " & _
                   " des_valor AS valor, tip_des_ingreso, tip_des_orden AS orden,'" & Fecha(2) & "'" & _
                      " FROM descuento " & _
                      " INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
                      " INNER JOIN empleado ON descuento.epl_codigo=empleado.epl_codigo AND descuento.emp_codigo=empleado.emp_codigo" & _
                      " INNER JOIN det_tip_descuento ON det_tip_descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND det_tip_descuento.emp_codigo=tipo_descuento.emp_codigo AND det_tip_descuento.are_lab_codigo=empleado.are_lab_codigo" & _
                      " INNER JOIN empresa ON empresa.emp_codigo=empleado.emp_codigo " & _
                      " WHERE descuento.emp_codigo LIKE '" & strEmpresa & "' AND descuento.epl_codigo LIKE '" & strNumero & "'" & _
                      " AND des_fecha BETWEEN '" & Fecha(0) & "' AND '" & Fecha(1) & "' AND det_tip_descuento.cta_codigo='" & strTipo & "' AND tipo_descuento.tip_des_ingreso=0 AND des_pagado=0 " & _
                      " ORDER BY tip_des_ingreso,tip_des_orden "
          clsConAux1.Ejecutar (strSql)
        
            strSql = " SELECT nombre,valor,tipo,orden FROM EstadoCuentaVB ORDER BY tipo ASC,orden "
            clsConAux2.Ejecutar (strSql)
            strSql = " SELECT nombre,valor,tipo,orden FROM EstadoCuentaVB ORDER BY tipo DESC,orden "
            clsConAux1.Ejecutar (strSql)
            While clsConAux1.adorec_Def.EOF = False
                strSql = " UPDATE EstadoCuentaVB SET nombres='" & clsConAux2.adorec_Def(0) & "', " & _
                         "valores='" & clsConAux2.adorec_Def(1) & "'," & _
                         "tipos='" & clsConAux2.adorec_Def(2) & "'" & _
                         " WHERE nombre ='" & clsConAux1.adorec_Def(0) & "' "
                clsConAUX.Ejecutar (strSql)
                clsConAux1.adorec_Def.MoveNext
                clsConAux2.adorec_Def.MoveNext
            Wend
          strSql = "select TOP 1 tipo,count(*) as t " & _
                   " From EstadoCuentaVB " & _
                   " group by tipo order by t desc "
          clsConAux1.Ejecutar (strSql)
          GetSQL = " SELECT * from EstadoCuentaVB ORDER BY tipo DESC, orden limit 0," & clsConAux1.adorec_Def(1)


    ElseIf strReporte = "rptPlanCuenta" Then
        Me.Caption = "Plan de Cuentas"
        GetSQL = " SELECT cta_codigo,cta_nombre,emp_nombre " & _
                 " FROM empresa INNER JOIN ctaconta ON empresa.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " ORDER BY cta_codigo "
    ElseIf strReporte = "rptLiqComision" Then
        Me.Caption = "Liquidación de Comisiones"
        GetSQL = " SELECT 'CAM-" & strNumero & vbNewLine & strAsiento & "' as descripcion_campania,'Comisión de la campaña: " & strNumero & "' as descripcion_campania2,li.per_codigo as li_per_codigo,concat(li.per_apellido,' ',li.per_nombre) as lider," & _
                 " em.per_codigo as em_per_codigo,concat(em.per_apellido,' ',em.per_nombre) as emprendedor," & _
                 " COALESCE(mul_nombre,'EJECUTIVO') as nivel," & _
                 " rc.red_cam_venta_directa,rc.red_cam_venta_directa_no_comi," & _
                 " IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta) as red_cam_venta_indirecta,IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_venta_indirecta_no_comi) as red_cam_venta_indirecta_no_comi," & _
                 " rc.red_cam_activo_directo,IIF(li.per_codigo=em.per_codigo,0,rc.red_cam_activo_indirecto) as red_cam_activo_indirecto," & _
                 " det_red_campania.mul_comision,'" & strTipo & "' as fechaLim"
        GetSQL = GetSQL & " FROM red_campania inner join persona li " & _
                 " ON red_campania.emp_codigo=li.emp_codigo " & _
                 " AND red_campania.per_codigo=li.per_codigo " & _
                 " INNER JOIN det_red_campania " & _
                 " ON red_campania.emp_codigo=det_red_campania.emp_codigo " & _
                 " AND red_campania.cam_anio=det_red_campania.cam_anio " & _
                 " AND red_campania.cam_mes=det_red_campania.cam_mes " & _
                 " AND red_campania.per_codigo=det_red_campania.per_papa_codigo " & _
                 " INNER JOIN persona em " & _
                 " ON det_red_campania.emp_codigo=em.emp_codigo " & _
                 " AND det_red_campania.per_codigo=em.per_codigo " & _
                 " INNER JOIN red_campania rc " & _
                 " ON det_red_campania.emp_codigo=rc.emp_codigo " & _
                 " AND det_red_campania.per_codigo=rc.per_codigo " & _
                 " AND red_campania.cam_anio=rc.cam_anio " & _
                 " AND red_campania.cam_mes=rc.cam_mes " & _
                 " LEFT JOIN multinivel " & _
                 " ON det_red_campania.emp_codigo=multinivel.emp_codigo " & _
                 " AND det_red_campania.mul_codigo=multinivel.mul_codigo "
        GetSQL = GetSQL & " WHERE red_campania.emp_codigo='" & strEmpresa & "' " & _
                 " AND red_campania.cam_anio='" & Left(strNumero, 4) & "' " & _
                 " AND red_campania.cam_mes='" & Right(strNumero, 2) & "' " & _
                 " AND li.per_codigo='" & Atencion & "' " & _
                 " AND (rc.red_cam_venta_directa!=0 or rc.red_cam_venta_directa_no_comi!=0 " & _
                 " or rc.red_cam_venta_indirecta!=0 or rc.red_cam_venta_indirecta_no_comi!=0 " & _
                 " or rc.red_cam_activo_directo!=0 or rc.red_cam_activo_indirecto!=0)" & _
                 "order by lider,mul_comision desc,mul_orden desc "
    ElseIf strReporte = "rptNotaVenta" Then
        strSqlAux = "SELECT egr_total " & _
                   " FROM egreso " & _
                   " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                   " AND egreso.egr_codigo='" & strNumero & "' " & _
                   " AND egreso.tip_egr_codigo='NOT' "
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("egr_total"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("egr_total") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        Set tNum2Text = Nothing
        GetSQL = " SELECT egreso.egr_codigo,CONCAT(per_apellido,' ',per_nombre) as per, egreso.per_codigo,per_ruc,per_direccion,per_telf,ciu_nombre,egr_fecha,DATEADD(d,for_pag_tiempo,egr_fecha) as vence,CONCAT(ven_apellido,' ',ven_nombre) as ven,det_egreso.prd_codigo,prd_nombre,det_egr_cantidad,uni_nombre,det_egr_precio,det_egr_precio*det_egr_cantidad as utot,egr_subtotal,egr_dcto,egr_subtotal_o,egr_impuesto,egr_total,for_pag_nombre,egreso.egr_observacion,'" & strValor & "' as valLetra,'" & PorIVA & "' as Piva " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON egreso.emp_codigo=forma_pago.emp_codigo AND egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN unidad ON producto.emp_codigo=unidad.emp_codigo AND producto.uni_codigo=unidad.uni_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='NOT' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' " & _
                 " UNION " & _
                 " SELECT egreso.egr_codigo,CONCAT(per_apellido,' ',per_nombre) as per, egreso.per_codigo,per_ruc,per_direccion,per_telf,ciu_nombre,egr_fecha,DATEADD(d,for_pag_tiempo,egr_fecha) as vence,CONCAT(ven_apellido,' ',ven_nombre) as ven,det_egreso_c.oca_codigo,oca_nombre,det_egr_c_cantidad,'' as uni_nombre,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,egr_subtotal,egr_dcto , egr_subtotal_o, egr_impuesto, egr_total, for_pag_nombre,egreso.egr_observacion,'VALOR EN LETRAS' as valLetra,'" & PorIVA & "' as Piva " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON egreso.emp_codigo=forma_pago.emp_codigo AND egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " INNER JOIN det_egreso_c ON egreso.emp_codigo=det_egreso_c.emp_codigo AND egreso.tip_egr_codigo=det_egreso_c.tip_egr_codigo AND egreso.egr_codigo=det_egreso_c.egr_codigo " & _
                 " INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='NOT' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' "
    ElseIf strReporte = "rptSTKDespacho" Then
        ReDim ped(UBound(Split(strNumero, ","))) As String
        
        Me.Caption = "Pedidos - " & strNumero
        ped = Split(strNumero, ",")
        If FormatoD0(strTipo) <> 2 Then
            strTipo = UBound(ped)
        End If
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBPedido" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBPedido" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " ped_codigo decimal(14,0) NOT NULL default '0', " & _
                   " ped_cdb varchar(50) default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,ped_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To UBound(ped)
            strSqlAux = " INSERT INTO CDBPedido" & strUsuario & "(emp_codigo,ped_codigo,ped_cdb) " & _
                        " VALUES('" & strEmpresa & "','" & ped(jj) & "'," & _
                        " '" & Replace(code128$(ped(jj)), "'", "''") & "') "
            clsConAUX.Ejecutar strSqlAux
        Next jj
        
        
        If FormatoD0(strTipo) = 2 Then
            strSqlAux = " SELECT per_fac_flete " & _
                        " FROM pedido INNER JOIN persona " & _
                        " ON pedido.emp_codigo=persona.emp_codigo " & _
                        " AND pedido.per_codigo=persona.per_codigo " & _
                        " AND persona.cat_p_tipo='C' " & _
                        " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
                        " AND pedido.ped_codigo in (" & strNumero & ")"
            clsConAUX.Ejecutar strSqlAux
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                If FormatoD0(clsConAUX.adorec_Def("per_fac_flete")) = 1 Then
                    strReporte = "rptSTKDespachoDirecto"
                End If
            End If
        End If
        
        GetSQL = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)IN ('EFE','CONT'),1,0) as ordenfp,pedido.ped_codigo,ped_cdb,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " IIF(persona.for_pag_codigo NOT IN ('EFE','CONT') OR pedido.ped_direccion_envio IS NULL or pedido.ped_direccion_envio='' OR LEFT(pedido.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,'-',persona.per_direccion2,' (',for_ent_nombre,')'),CONCAT(pedido.ped_direccion_envio,' (',for_ent_nombre,')')) as per_direccion2,persona.per_direccion,CONCAT(persona.per_telf,'/',persona.per_fax,'/',persona.per_celular) as per_tfc,persona.per_codigo_postal," & _
                 " ped_fecha,for_pag_nombre,COALESCE(dis_pol_nombre,'') as dis_pol_nombre," & _
                 " CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1, " & _
                 " CURRENT_TIMESTAMP as hoy, " & _
                 " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')),''))))))))) as papa"
        GetSQL = GetSQL & " FROM pedido INNER JOIN CDBPedido" & strUsuario & " ON pedido.emp_codigo=CDBPedido" & strUsuario & ".emp_codigo" & _
                 " AND pedido.ped_codigo=CDBPedido" & strUsuario & ".ped_codigo " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo " & _
                 " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
                 " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo " & _
                 " AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1 " & _
                 " LEFT JOIN persona as N3 ON persona.emp_codigo = N3.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1 " & _
                 " LEFT JOIN persona as N4 ON persona.emp_codigo = N4.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " LEFT JOIN distribucion_politica ON persona.dis_pol_codigo=distribucion_politica.dis_pol_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9,pedido.ped_codigo "
    
        'rptResumenDespacho
    ElseIf strReporte = "rptResumenDespacho" Then
        ReDim ped(FormatoD0(strTipo)) As String
        
        Me.Caption = "Pedidos - " & strNumero
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBPedido" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBPedido" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " ped_codigo decimal(14,0) NOT NULL default '0', " & _
                   " PRIMARY KEY  (emp_codigo,ped_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To FormatoD0(strTipo) - 2
            
            strSqlAux = " INSERT INTO CDBPedido" & strUsuario & "(emp_codigo,ped_codigo) " & _
                        " VALUES('" & strEmpresa & "','" & ped(jj) & "') "
            clsConAUX.Ejecutar strSqlAux
            
        Next jj
        
        GetSQL = " SELECT IIF(persona.for_pag_codigo='CONT',0,1) as orden,pedido.ped_codigo,pedido.ped_usumod,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per," & _
                 " for_pag_nombre,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')) as nn1," & _
                 " LEFT(CURRENT_TIMESTAMP,10) as hoy," & _
                 " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,'')))>2,CONCAT(COALESCE(N4.per_apellido,''),' ',COALESCE(N4.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,'')))>2,CONCAT(COALESCE(N3.per_apellido,''),' ',COALESCE(N3.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,'')))>2,CONCAT(COALESCE(N2.per_apellido,''),' ',COALESCE(N2.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')))>2,CONCAT(COALESCE(N1.per_apellido,''),' ',COALESCE(N1.per_nombre,'')),''))))))))) as papa"
        GetSQL = GetSQL & " FROM pedido INNER JOIN CDBPedido" & strUsuario & " ON pedido.emp_codigo=CDBPedido" & strUsuario & ".emp_codigo" & _
                 " AND pedido.ped_codigo=CDBPedido" & strUsuario & ".ped_codigo " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona N1 ON N1.emp_codigo=persona.emp_codigo " & _
                 " AND N1.per_codigo=persona.per_codigo_ref AND N1.per_es_gz=1 " & _
                 " LEFT JOIN persona N2 ON N2.emp_codigo=persona.emp_codigo " & _
                 " AND N2.per_codigo=persona.per_codigo_ref2 AND N2.per_es_di=1 " & _
                 " LEFT JOIN persona as N3 ON persona.emp_codigo = N3.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = N3.per_codigo AND N3.per_es_em=1 " & _
                 " LEFT JOIN persona as N4 ON persona.emp_codigo = N4.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = N4.per_codigo AND N4.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo in (" & strNumero & ") " & _
                 " ORDER BY orden,nn1,papa,pedido.ped_codigo "
    
    ElseIf strReporte = "rptEtiquetaDespacho" Then
        Dim PreGuia As String
        Me.Caption = "Factura - " & strNumero
        CBF = Replace(code128$(strNumero), "'", "''")
        strSqlAux = "SELECT par_texto " & _
                   " FROM parametro " & _
                   " WHERE emp_codigo='" & strEmpresa & "' AND par_codigo='GTR' "
        clsConAUX.Ejecutar strSqlAux
        PreGuia = clsConAUX.adorec_Def("par_texto")
        CBG = Replace(code128$(PreGuia & strNumero), "'", "''")
        
        GetSQL = " SELECT '" & CBF & "' as CBF,'" & CBG & "' as CBG,egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, persona.per_ruc,persona.per_direccion,persona.per_telf,ciu_nombre as ciudad," & _
                 " egr_fecha,CONCAT(LEFT(DATEADD(d,for_pag_tiempo,egr_fecha),10),' (',for_pag_nombre,')') as vence,for_pag_nombre," & _
                 " egr_total,CONCAT('Nº: ',egreso.egr_codigo,' - ',time_format(current_timestamp,'%H:%i'),persona.cat_p_codigo) as todo,for_ent_nombre,CONCAT(ciu_nombre,' ',persona.per_direccion2) as per_direccion2," & _
                 " egr_fechamod as fechamod, egr_usumod as usumod, CURRENT_TIMESTAMP as hoy"
        GetSQL = GetSQL & " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='" & strTipo & "' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' "
    
    ElseIf strReporte = "rptGuiaFactura" Then
        Me.Caption = "GuiaEgreso - " & strNumero
        
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBGuiFac" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBGuiFac" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " tip_egr_codigo char(3) NOT NULL default ''," & _
                   " egr_codigo decimal(14,0) NOT NULL default '0', " & _
                   " aut_cdb varchar(50) default NULL, " & _
                   " aut varchar(50) default NULL, " & _
                   " cla varchar(50) default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,tip_egr_codigo,egr_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To UBound(ped)
            
            CBF = Replace(code128$(ped(jj)), "'", "''")
            
            clsConAUX.Inicializar AdoConn, AdoConnMaster
            
            strSql = " SELECT COALESCE(doc_ele_claveacceso,'') as doc_ele_claveacceso,COALESCE(doc_ele_autorizacion,'') as doc_ele_autorizacion " & _
                     " FROM doc_electronico INNER JOIN egreso_guia " & _
                     " ON doc_electronico.emp_codigo=egreso_guia.emp_codigo " & _
                     " AND doc_electronico.doc_ele_codigo=egreso_guia.egr_gui_codigo " & _
                     " WHERE doc_electronico.emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_ele_coddoc='06' " & _
                     " AND egreso_guia.tip_egr_codigo='FAC'" & _
                     " AND egreso_guia.egr_codigo='" & ped(jj) & "'"
            
            clsConAUX.Ejecutar strSql
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                CDEClaveAcceso = Replace(code128$(clsConAUX.adorec_Def("doc_ele_claveacceso")), "'", "''")
                ClaveAcceso = clsConAUX.adorec_Def("doc_ele_claveacceso")
                Autori = clsConAUX.adorec_Def("doc_ele_autorizacion")
            Else
                CDEClaveAcceso = ""
                ClaveAcceso = ""
                Autori = ""
            End If
            
        
            strSqlAux = " INSERT INTO CDBGuiFac" & strUsuario & "(emp_codigo,tip_egr_codigo,egr_codigo," & _
                        " aut_cdb,aut,cla) " & _
                        " VALUES('" & strEmpresa & "','FAC','" & ped(jj) & "'," & _
                        " '" & CDEClaveAcceso & "','" & Autori & "','" & ClaveAcceso & "') "
            clsConAUX.Ejecutar strSqlAux
        Next jj
        
        
        GetSQL = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp ,orden," & _
                 " aut_cdb AS CDEClaveAcceso,emp_nombre,emp_direccion,emp_telf,emp_ruc,cla as doc_ele_claveacceso,aut as doc_ele_autorizacion," & _
                 " egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf," & _
                 " emp_nombre,egr_fecha,DATEADD(d,1,egr_fecha) as egr_fecha2," & _
                 " de.prd_codigo as prd_codigo,de.nombre as nombre,ROUND(cantidad,2) as cantidad," & _
                 " de.mar_codigo,de.gru_codigo," & _
                 " FORMAT(egr_fecha,'yyyy-MM-dd') as fech,tip_egr_nombre,IIF(egreso.tip_egr_codigo='FAC','VENTA','CONSIGNACION') AS motivo_traslado," & _
                 " COALESCE(cou_razon_social,'') as cou_razon_social,COALESCE(cou_ruc,'') as cou_ruc," & _
                 " COALESCE(egr_gui_placa,'') as egr_gui_placa,COALESCE(IIF(cou_direccion='',persona.per_direccion,cou_direccion),persona.per_direccion) as cou_direccion,COALESCE(egr_gui_serie,'') as egr_gui_serie,COALESCE(egr_gui_numero,0) as egr_gui_numero,COALESCE(cou_ruta,'') as cou_ruta"
        GetSQL = GetSQL & " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo" & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN CDBGuiFac" & strUsuario & " ON egreso.emp_codigo=CDBGuiFac" & strUsuario & ".emp_codigo " & _
                 " AND egreso.tip_egr_codigo=CDBGuiFac" & strUsuario & ".tip_egr_codigo " & _
                 " AND egreso.egr_codigo=CDBGuiFac" & strUsuario & ".egr_codigo " & _
                 " INNER JOIN ("
        GetSQL = GetSQL & " SELECT '0' AS orden,det_egreso.emp_codigo,det_egreso.tip_egr_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo, " & _
                 " ROUND(det_egr_cantidad,2) AS cantidad," & _
                 " prd_nombre AS nombre,mar_codigo,producto.gru_codigo " & _
                 " FROM det_egreso_ubicacion det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
        GetSQL = GetSQL & " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso.tip_egr_codigo='FAC' and det_egreso.prd_codigo NOT LIKE 'PR-%' " & _
                 " AND det_egreso.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & " UNION SELECT '2' as orden,det_egreso_c.emp_codigo,det_egreso_c.tip_egr_codigo,det_egreso_c.egr_codigo,det_egreso_c.oca_codigo, " & _
                 " ROUND(det_egr_c_cantidad,2) as cantidad, " & _
                 " oca_nombre as nombre,'' as mar_codigo,'' as gru_codigo " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                 " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                 " AND det_egreso_c.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & ") de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo " & _
                 " LEFT JOIN (SELECT egreso_guia.emp_codigo,egreso_guia.tip_egr_codigo,egreso_guia.egr_codigo," & _
                 " cou_razon_social,cou_ruc,egr_gui_placa,cou_direccion,egr_gui_serie,egr_gui_numero,cou_ruta " & _
                 " FROM egreso_guia INNER JOIN courier ON egreso_guia.emp_codigo=courier.emp_codigo" & _
                 " AND egreso_guia.cou_codigo=courier.cou_codigo" & _
                 " WHERE egreso_guia.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso_guia.tip_egr_codigo='FAC' " & _
                 " AND egreso_guia.egr_codigo in (" & strNumero & ") " & _
                 " )egr_fac ON egreso.emp_codigo=egr_fac.emp_codigo" & _
                 " AND egreso.tip_egr_codigo=egr_fac.tip_egr_codigo" & _
                 " AND egreso.egr_codigo=egr_fac.egr_codigo" & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo,orden,mar_codigo,LEFT(gru_codigo,2),nombre  "
        
    
    ElseIf strReporte = "rptFacturaGuia" Or strReporte = "rptFacturaSola" Then
        Me.Caption = "Factura - " & strNumero
        
        GetSQL = " SELECT DISTINCT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp," & _
                 " per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo " & _
                 " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo "
        GetSQL = GetSQL & " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " egr_codigo "
        
    ElseIf strReporte = "rptFactura" Then
        Me.Caption = "Factura - " & strNumero
        
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBFac" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBFac" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " egr_codigo decimal(14,0) NOT NULL default '0'," & _
                   " ped_codigo decimal(14,0) NOT NULL default '0', " & _
                   " egr_cdb varchar(50) default NULL, " & _
                   " ped_cdb varchar(50) default NULL, " & _
                   " aut_cdb varchar(50) default NULL, " & _
                   " aut varchar(50) default NULL, " & _
                   " cla varchar(50) default NULL, " & _
                   " egr_cantidad decimal(14,2) NOT NULL default '0', " & _
                   " egr_totalletra varchar(255) default NULL, " & _
                   " dir text default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,egr_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To UBound(ped)
            
            CBF = Replace(code128$(ped(jj)), "'", "''")
            
            clsConAUX.Inicializar AdoConn, AdoConnMaster
            
            strSql = " SELECT COALESCE(doc_ele_claveacceso,'') as doc_ele_claveacceso,COALESCE(doc_ele_autorizacion,'') as doc_ele_autorizacion " & _
                     " FROM doc_electronico " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_ele_coddoc='01' " & _
                     " ANd doc_ele_codigo='" & ped(jj) & "'"
            
            clsConAUX.Ejecutar strSql
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                CDEClaveAcceso = Replace(code128$(clsConAUX.adorec_Def("doc_ele_claveacceso")), "'", "''")
                ClaveAcceso = clsConAUX.adorec_Def("doc_ele_claveacceso")
                Autori = clsConAUX.adorec_Def("doc_ele_autorizacion")
            Else
                CDEClaveAcceso = ""
            End If
            
            strSqlAux = " SELECT egreso.emp_codigo,egr_total,COALESCE(ped_codigo,'0') as ped_codigo," & _
                       " COALESCE(sum(det_egr_cantidad),0) as n,COALESCE(ped_direccion_envio,'') as dir " & _
                       " FROM egreso LEFT JOIN pedido " & _
                       " ON egreso.emp_codigo=pedido.emp_codigo" & _
                       " AND egreso.tip_egr_codigo=pedido.ped_tip_egr_codigo" & _
                       " AND egreso.egr_codigo=pedido.ped_egr_codigo AND pedido.ped_estado IN (2,8) " & _
                       " AND pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_tip_egr_codigo='FAC' " & _
                       " AND pedido.ped_egr_codigo='" & ped(jj) & "'" & _
                       " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                       " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                       " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                       " AND det_egreso.prd_codigo!='PR-CARGOO100330TU' " & _
                       " AND det_egreso.emp_codigo='" & strEmpresa & "' AND det_egreso.tip_egr_codigo='FAC' " & _
                       " AND det_egreso.egr_codigo='" & ped(jj) & "'" & _
                       " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo='FAC' " & _
                       " AND egreso.egr_codigo='" & ped(jj) & "' " & _
                       " GROUP BY egreso.emp_codigo,egr_total,pedido.ped_codigo,ped_direccion_envio"
            clsConAUX.Ejecutar strSqlAux
            dirEnvio = UCase(clsConAUX.adorec_Def("dir"))
            lngValor = Int(clsConAUX.adorec_Def("egr_total"))
            intValor = Right(Str(Int(clsConAUX.adorec_Def("egr_total") * 100)), 2)
            strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
            Set tNum2Text = Nothing
            
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                canti = FormatoD2(clsConAUX.adorec_Def("n"))
                CBP = Replace(code128$(clsConAUX.adorec_Def("ped_codigo")), "'", "''")
            Else
                canti = 0
                CBP = ""
            End If
            
            
        
            strSqlAux = " INSERT INTO CDBFac" & strUsuario & "(emp_codigo,egr_codigo,ped_codigo,egr_cdb," & _
                        " ped_cdb,aut_cdb,aut,cla,egr_cantidad,egr_totalletra,dir) " & _
                        " VALUES('" & strEmpresa & "','" & ped(jj) & "','" & clsConAUX.adorec_Def("ped_codigo") & "','" & CBF & "'," & _
                        " '" & CBP & "','" & CDEClaveAcceso & "','" & Autori & "','" & ClaveAcceso & "','" & canti & "','" & strValor & "','" & Replace(dirEnvio, "'", " ") & "') "
            clsConAUX.Ejecutar strSqlAux
        Next jj
        
        
        GetSQL = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp ,orden, ped_codigo,egr_cdb as CBF,ped_cdb as CBP," & _
                 " aut_cdb AS CDEClaveAcceso,emp_nombre,emp_direccion,emp_telf,emp_ruc,cla as doc_ele_claveacceso,aut as doc_ele_autorizacion," & _
                 " egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf," & _
                 " ciu_nombre as ciudad,vendedor.ven_codigo as ven," & _
                 " emp_nombre,egr_fecha,CONCAT(LEFT(DATEADD(d, for_pag_tiempo,egr_fecha),10),' (',for_pag_nombre,')') as vence," & _
                 " CONCAT(ven_apellido,' ',ven_nombre) as vendedor,egr_dcto,for_pag_nombre," & _
                 " de.prd_codigo as prd_codigo,de.nombre as nombre,de.prd_ubica_linea,ROUND(cantidad,2) as cantidad," & _
                 " ROUND(det_egr_precio,3) as det_egr_precio,utot,egr_subtotal," & _
                 " egr_dcto,egr_subtotal_o,egr_impuesto,det_egr_dcto,s,egr_total,for_pag_nombre," & _
                 " CONCAT('Observaciones: ',egreso.egr_observacion) as egr_observacion,egr_totalletra as valLetra," & _
                 " cod_iva_porcentaje as Piva,de.mar_codigo,de.gru_codigo,de.gru_nombre," & _
                 " CONCAT('Nº de Factura: ',FORMAT(RIGHT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 7),3)*1,'000'),'-',FORMAT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 10)*1,'000'),'-',FORMAT(Right(egreso.egr_codigo, 7)*1,'000000000'),' - ',FORMAT(current_timestamp,'HH:MM'),persona.cat_p_codigo) as todo," & _
                 " FORMAT(egr_fecha,'yyyy-MM-dd') as fech,egr_cantidad as cants,for_ent_nombre," & _
                 " IIF(CAST(dir as varchar)='',persona.per_direccion2,dir) as per_direccion2,"
        GetSQL = GetSQL & " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')))>2,CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')))>2,CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')))>2,CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')))>2,CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')),''))))))))) as papa,"
        GetSQL = GetSQL & " CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')) as gerente," & _
                 " CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')) as director," & _
                 " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as EMPR," & _
                 " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as EJES," & _
                 " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5," & _
                 " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6," & _
                 " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7," & _
                 " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8," & _
                 " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9,"
        GetSQL = GetSQL & " IIF(p1.per_codigo IS NULL,'','Coord. 9:') as ger," & _
                 " IIF(p2.per_codigo IS NULL,'','Coord. 8:') as dir," & _
                 " IIF(EMP.per_codigo IS NULL,'','Coord. 7:') as nemp," & _
                 " IIF(EJE.per_codigo IS NULL,'','Coord. 6:') as neje," & _
                 " IIF(N5.per_codigo IS NULL,'','Coord. 5:') as nt5," & _
                 " IIF(N6.per_codigo IS NULL,'','Coord. 4:') as nt6," & _
                 " IIF(N7.per_codigo IS NULL,'','Coord. 3:') as nt7," & _
                 " IIF(N8.per_codigo IS NULL,'','Coord. 2:') as nt8," & _
                 " IIF(N9.per_codigo IS NULL,'','Coord. 1:') as nt9," & _
                 " egr_fechamod as fechamod, egr_usumod as usumod," & _
                 " persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9, " & _
                 " CONCAT('Estimados. De acuerdo con la información registrada en nuestro sistema, en tu  mail: - ',persona.per_email,' - recibirás tus documentos electrónicos autorizada por el SRI, según las nueva ley en vigencia. Si no tienes actualizados tus datos comunicate al 1800CATALOGOS para pedir esta actualización') as mensaje, " & _
                 " IIF(aut='' OR aut is null,'DOCUMENTO SIN VALIDEZ TRIBUTARIA','') as mensaje2 "
        GetSQL = GetSQL & " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN codigo_iva ON egreso.cod_iva_codigo=codigo_iva.cod_iva_codigo" & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN CDBFac" & strUsuario & " ON egreso.emp_codigo=CDBFac" & strUsuario & ".emp_codigo " & _
                 " AND egreso.egr_codigo=CDBFac" & strUsuario & ".egr_codigo " & _
                 " INNER JOIN ("
        GetSQL = GetSQL & " SELECT '0' AS orden,det_egreso.emp_codigo,det_egreso.tip_egr_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo, " & _
                 " ROUND(det_egr_cantidad,2) AS cantidad," & _
                 " ROUND(det_egr_precio,3) AS det_egr_precio,det_egr_precio*(det_egr_cantidad) AS utot, " & _
                 " COALESCE(det_egr_pdcto,(det_egr_dcto/det_egr_cantidad* det_egr_cantidad)) AS det_egr_dcto, " & _
                 " IIF(det_egr_pdcto IS NOT NULL,'%','') AS s,prd_nombre AS nombre,mar_codigo,producto.gru_codigo,gru_nombre,det_egreso.prd_ubica_linea " & _
                 " FROM det_egreso_ubicacion det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
        GetSQL = GetSQL & " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso.tip_egr_codigo='FAC' " & _
                 " AND det_egreso.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & " UNION SELECT '2' as orden,det_egreso_c.emp_codigo,det_egreso_c.tip_egr_codigo,det_egreso_c.egr_codigo,det_egreso_c.oca_codigo, " & _
                 " ROUND(det_egr_c_cantidad,2) as cantidad,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,'0.0000' as det_egr_dcto, " & _
                 " '' as s,oca_nombre as nombre,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,'' as prd_ubica_linea " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                 " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='FAC' " & _
                 " AND det_egreso_c.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & ") de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p2.per_es_di=1 " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " ped_codigo,egr_codigo,orden,prd_ubica_linea ,mar_codigo,LEFT(gru_codigo,2),gru_nombre,nombre  "
                 
    
        
    ElseIf strReporte = "rptNotaEntregaSuministro" Then
        
        Me.Caption = "Nota Entrega Suministro - " & strNumero
        
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBFac" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBFac" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " egr_codigo decimal(14,0) NOT NULL default '0',ped_codigo decimal(14,0) NOT NULL default '0', " & _
                   " egr_cdb varchar(50) default NULL, " & _
                   " ped_cdb varchar(50) default NULL, " & _
                   " aut_cdb varchar(50) default NULL, " & _
                   " aut varchar(50) default NULL, " & _
                   " cla varchar(50) default NULL, " & _
                   " egr_cantidad decimal(14,2) NOT NULL default '0', " & _
                   " egr_totalletra varchar(255) default NULL, " & _
                   " dir text default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,egr_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To UBound(ped)
            
            CBF = Replace(code128$(ped(jj)), "'", "''")
            
            clsConAUX.Inicializar AdoConn, AdoConnMaster
            
            
            
            strSqlAux = " SELECT egreso.emp_codigo,egr_total,COALESCE(ped_codigo,'0') as ped_codigo,COALESCE(sum(det_egr_cantidad),0) as n,COALESCE(ped_direccion_envio,'') as dir " & _
                       " FROM egreso LEFT JOIN pedido " & _
                       " ON egreso.emp_codigo=pedido.emp_codigo" & _
                       " AND egreso.tip_egr_codigo=pedido.ped_tip_egr_codigo" & _
                       " AND egreso.egr_codigo=pedido.ped_egr_codigo AND pedido.ped_estado IN (2,8,10) " & _
                       " AND pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_tip_egr_codigo='NET' " & _
                       " AND pedido.ped_egr_codigo='" & ped(jj) & "'" & _
                       " LEFT JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
                       " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                       " AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                       " AND det_egreso.prd_codigo!='PR-CARGOO100330TU' " & _
                       " AND det_egreso.emp_codigo='" & strEmpresa & "' AND det_egreso.tip_egr_codigo='NET' " & _
                       " AND det_egreso.egr_codigo='" & ped(jj) & "'" & _
                       " WHERE egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo='NET' " & _
                       " AND egreso.egr_codigo='" & ped(jj) & "' " & _
                       " GROUP BY egreso.emp_codigo,egr_total,pedido.ped_codigo,ped_direccion_envio"
            clsConAUX.Ejecutar strSqlAux
            dirEnvio = UCase(clsConAUX.adorec_Def("dir"))
            lngValor = Int(clsConAUX.adorec_Def("egr_total"))
            intValor = Right(Str(Int(clsConAUX.adorec_Def("egr_total") * 100)), 2)
            strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
            Set tNum2Text = Nothing
            
            If clsConAUX.adorec_Def.RecordCount > 0 Then
                canti = FormatoD0(clsConAUX.adorec_Def("n"))
                CBP = Replace(code128$(clsConAUX.adorec_Def("ped_codigo")), "'", "''")
            Else
                canti = 0
                CBP = ""
            End If
            
            
        
            strSqlAux = " INSERT INTO CDBFac" & strUsuario & "(emp_codigo,egr_codigo,ped_codigo,egr_cdb," & _
                        " ped_cdb,aut_cdb,aut,cla,egr_cantidad,egr_totalletra,dir) " & _
                        " VALUES('" & strEmpresa & "','" & ped(jj) & "','" & clsConAUX.adorec_Def("ped_codigo") & "','" & CBF & "'," & _
                        " '" & CBP & "','" & CDEClaveAcceso & "','" & Autori & "','" & ClaveAcceso & "','" & canti & "','" & strValor & "','" & dirEnvio & "') "
            clsConAUX.Ejecutar strSqlAux
        Next jj
        
        
        GetSQL = " SELECT IIF(IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp) IN ('CONT','EFE'),1,0) as ordenfp ,orden, ped_codigo,egr_cdb as CBF,ped_cdb as CBP," & _
                 " aut_cdb AS CDEClaveAcceso,emp_nombre,emp_direccion,emp_telf,emp_ruc,cla as doc_ele_claveacceso,aut as doc_ele_autorizacion," & _
                 " egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf," & _
                 " ciu_nombre as ciudad,vendedor.ven_codigo as ven," & _
                 " emp_nombre,egr_fecha,CONCAT(LEFT(DATEADD(d, for_pag_tiempo,egr_fecha),10),' (',for_pag_nombre,')') as vence," & _
                 " CONCAT(ven_apellido,' ',ven_nombre) as vendedor,egr_dcto,for_pag_nombre," & _
                 " de.prd_codigo as prd_codigo,de.nombre as nombre,de.prd_ubica_linea,ROUND(cantidad,2) as cantidad," & _
                 " ROUND(det_egr_precio,3) as det_egr_precio,utot,egr_subtotal," & _
                 " egr_dcto,egr_subtotal_o,egr_impuesto,det_egr_dcto,s,egr_total,for_pag_nombre," & _
                 " CONCAT('Observaciones: ',egreso.egr_observacion) as egr_observacion,egr_totalletra as valLetra," & _
                 " de.mar_codigo,de.gru_codigo,de.gru_nombre," & _
                 " CONCAT('Nº de Nota de Entrega: ',FORMAT(RIGHT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 7),3)*1,'000'),'-',FORMAT(LEFT(egreso.egr_codigo, LEN(egreso.egr_codigo) - 10)*1,'000'),'-',FORMAT(Right(egreso.egr_codigo, 7)*1,'000000000'),' - ',FORMAT(current_timestamp,'HH:MM'),persona.cat_p_codigo) as todo," & _
                 " FORMAT(egr_fecha,'yyyy-MM-dd') as fech,egr_cantidad as cants,for_ent_nombre," & _
                 " IIF(CAST(dir as varchar)='',persona.per_direccion2,dir) as per_direccion2,"
        GetSQL = GetSQL & " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')))>2,CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')))>2,CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')))>2,CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')))>2,CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')),''))))))))) as papa,"
        GetSQL = GetSQL & " CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')) as gerente," & _
                 " CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')) as director," & _
                 " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as EMPR," & _
                 " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as EJES," & _
                 " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5," & _
                 " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6," & _
                 " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7," & _
                 " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8," & _
                 " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9,"
        GetSQL = GetSQL & " IIF(p1.per_codigo IS NULL,'','Coord. 9:') as ger," & _
                 " IIF(p2.per_codigo IS NULL,'','Coord. 8:') as dir," & _
                 " IIF(EMP.per_codigo IS NULL,'','Coord. 7:') as nemp," & _
                 " IIF(EJE.per_codigo IS NULL,'','Coord. 6:') as neje," & _
                 " IIF(N5.per_codigo IS NULL,'','Coord. 5:') as nt5," & _
                 " IIF(N6.per_codigo IS NULL,'','Coord. 4:') as nt6," & _
                 " IIF(N7.per_codigo IS NULL,'','Coord. 3:') as nt7," & _
                 " IIF(N8.per_codigo IS NULL,'','Coord. 2:') as nt8," & _
                 " IIF(N9.per_codigo IS NULL,'','Coord. 1:') as nt9," & _
                 " egr_fechamod as fechamod, egr_usumod as usumod," & _
                 " persona.per_codigo_ref,persona.per_codigo_ref2,persona.per_codigo_ref3,persona.per_codigo_ref4,persona.per_codigo_ref5,persona.per_codigo_ref6,persona.per_codigo_ref7,persona.per_codigo_ref8,persona.per_codigo_ref9, " & _
                 " '' as mensaje, " & _
                 " '' as mensaje2 "
        GetSQL = GetSQL & " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND IIF(persona.for_pag_codigo_imp IS NULL OR persona.for_pag_codigo_imp='',persona.for_pag_codigo,persona.for_pag_codigo_imp)=forma_pago.for_pag_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN CDBFac" & strUsuario & " ON egreso.emp_codigo=CDBFac" & strUsuario & ".emp_codigo " & _
                 " AND egreso.egr_codigo=CDBFac" & strUsuario & ".egr_codigo " & _
                 " INNER JOIN ("
        GetSQL = GetSQL & " SELECT '0' AS orden,det_egreso.emp_codigo,det_egreso.tip_egr_codigo,det_egreso.egr_codigo,det_egreso.prd_codigo, " & _
                 " ROUND(det_egr_cantidad,2) AS cantidad," & _
                 " ROUND(det_egr_precio,3) AS det_egr_precio,det_egr_precio*(det_egr_cantidad) AS utot, " & _
                 " COALESCE(det_egr_pdcto,(det_egr_dcto/det_egr_cantidad* det_egr_cantidad)) AS det_egr_dcto, " & _
                 " IIF(det_egr_pdcto IS NOT NULL,'%','') AS s,prd_nombre AS nombre,mar_codigo,producto.gru_codigo,gru_nombre,det_egreso.prd_ubica_linea " & _
                 " FROM det_egreso_ubicacion det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
        GetSQL = GetSQL & " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso.tip_egr_codigo='NET' " & _
                 " AND det_egreso.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & " UNION SELECT '2' as orden,det_egreso_c.emp_codigo,det_egreso_c.tip_egr_codigo,det_egreso_c.egr_codigo,det_egreso_c.oca_codigo, " & _
                 " ROUND(det_egr_c_cantidad,2) as cantidad,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,'0.0000' as det_egr_dcto, " & _
                 " '' as s,oca_nombre as nombre,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,'' as prd_ubica_linea " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                 " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='NET' " & _
                 " AND det_egreso_c.egr_codigo in (" & strNumero & ")"
        GetSQL = GetSQL & ") de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p2.per_es_di=1 " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='NET' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY ordenfp,per_codigo_ref,per_codigo_ref2,per_codigo_ref3,per_codigo_ref4,per_codigo_ref5,per_codigo_ref6,per_codigo_ref7,per_codigo_ref8,per_codigo_ref9," & _
                 " ped_codigo,egr_codigo,orden,prd_ubica_linea ,mar_codigo,LEFT(gru_codigo,2),gru_nombre,nombre  "
                 
    
    ElseIf strReporte = "rptFacIncentivo" Then
        'Dim canti As String
        'Dim dirEnvio As String
        
        GetSQL = " SELECT orden,egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " de.prd_codigo as prd_codigo,de.nombre as nombre,ROUND(cantidad,2) as cantidad "
        GetSQL = GetSQL & " FROM egreso INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ("
        GetSQL = GetSQL & " SELECT '0' as orden,pedido.emp_codigo,pedido.ped_tip_egr_codigo as tip_egr_codigo," & _
                 " pedido.ped_egr_codigo as egr_codigo,'' as prd_codigo, " & _
                 " 0 as cantidad," & _
                 " 'VACIO' as nombre" & _
                 " FROM pedido "
        GetSQL = GetSQL & " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_tip_egr_codigo='FAC' " & _
                 " AND pedido.ped_egr_codigo in (" & strNumero & ") " & _
                 " UNION"
        GetSQL = GetSQL & " SELECT '1' as orden,pedido.emp_codigo,pedido.ped_tip_egr_codigo as tip_egr_codigo," & _
                 " pedido.ped_egr_codigo as egr_codigo,det_pedido.prd_codigo, " & _
                 " ROUND(det_ped_cant_entregada,2) as cantidad," & _
                 " prd_nombre as nombre" & _
                 " FROM pedido INNER JOIN det_pedido " & _
                 " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo "
        GetSQL = GetSQL & " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_tip_egr_codigo='FAC' " & _
                 " AND pedido.ped_egr_codigo in (" & strNumero & ") " & _
                 " AND det_pedido.det_ped_incentivo=1 " & _
                 " AND ROUND(det_ped_cant_entregada,2)!=0 "
        GetSQL = GetSQL & " ) de ON egreso.emp_codigo=de.emp_codigo AND egreso.tip_egr_codigo=de.tip_egr_codigo AND egreso.egr_codigo=de.egr_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo in (" & strNumero & ") " & _
                 " ORDER BY " & _
                 " egr_codigo,orden DESC,nombre "
    ElseIf strReporte = "rptFacReprogramacion" Then
        
        GetSQL = " SELECT * " & _
                 " FROM Fn_DetFacturaReprogamada('" & strEmpresa & "'," & strNumero & ")"
        
    ElseIf strReporte = "rptRC" Then
        Dim cantid As String
        Me.Caption = "Factura - " & strNumero
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = "SELECT egr_total " & _
                   " FROM egreso " & _
                   " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                   " AND egreso.egr_codigo='" & strNumero & "' "
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("egr_total"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("egr_total") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        Set tNum2Text = Nothing
        
        strSqlAux = "SELECT COALESCE(sum(det_egr_cantidad),0) as n " & _
                   " FROM det_egreso " & _
                   " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
                   " AND det_egreso.egr_codigo='" & strNumero & "' " & _
                   " AND tip_egr_codigo='FAC' " & _
                   " GROUP BY emp_codigo "
        clsConAUX.Ejecutar strSqlAux
        If clsConAUX.adorec_Def.RecordCount > 0 Then
            cantid = FormatoD2(clsConAUX.adorec_Def("n"))
        Else
            cantid = 0
        End If
        
        GetSQL = " SELECT 1,egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf,ciu_nombre as ciudad,vendedor.ven_codigo as ven," & _
                 " emp_nombre,egr_fecha,DATEADD(d,for_pag_tiempo,egr_fecha) as vence,CONCAT(ven_apellido,' ',ven_nombre) as vendedor,egr_dcto,for_pag_nombre," & _
                 " det_egreso.prd_codigo as prd_codigo,prd_nombre as nombre,ROUND(det_egr_cantidad,2) as cantidad,ROUND(det_egr_precio,3) as det_egr_precio,det_egr_precio*det_egr_cantidad-det_egr_dcto as utot,egr_subtotal," & _
                 " egr_dcto,egr_subtotal_o,egr_impuesto,COALESCE(det_egr_pdcto,det_egr_dcto) as det_egr_dcto,IIF(det_egr_pdcto is not null,'%','') as s,egr_total,for_pag_nombre,egreso.egr_observacion as egr_observacion,'" & strValor & "' as valLetra,'" & PorIVA & "%' as Piva,mar_codigo,producto.gru_codigo,gru_nombre,CONCAT('Nº: ',egreso.egr_codigo,' - ',time_format(current_timestamp,'%H:%i'),persona.cat_p_codigo) as todo,FORMAT(egr_fecha,'dd-mm-yyyy') as fech,'" & cantid & "' as cants,for_ent_nombre,CONCAT(ciu_nombre,' ',persona.per_direccion2) as per_direccion2,COALESCE(CONCAT(p1.per_apellido,' ',p1.per_nombre),'') as gerente,COALESCE(CONCAT(p2.per_apellido,' ',p2.per_nombre),'') as director,IIF(p1.per_codigo IS NULL,'','Distribuidor:') as ger,IIF(p2.per_codigo IS NULL,'','Director:') as dir,egr_fechamod as fechamod, egr_usumod as usumod " & _
                 " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                 " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' "
        GetSQL = GetSQL & " UNION " & _
                 " SELECT 2,egreso.egr_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, egreso.per_codigo,persona.per_ruc,persona.per_direccion,persona.per_telf,ciu_nombre as ciudad,vendedor.ven_codigo as ven," & _
                 " emp_nombre,egr_fecha,DATEADD(d,for_pag_tiempo,egr_fecha) as vence,CONCAT(ven_apellido,' ',ven_nombre) as vendedor,egr_dcto,for_pag_nombre," & _
                 " det_egreso_c.oca_codigo AS prd_codigo,oca_nombre as nombre,ROUND(det_egr_c_cantidad,2) as cantidad,det_egr_c_precio,det_egr_c_precio*det_egr_c_cantidad as utot,egr_subtotal," & _
                 " egr_dcto , egr_subtotal_o, egr_impuesto,'0.0000' as det_egr_dcto,'' as s, egr_total, for_pag_nombre,egreso.egr_observacion,'" & strValor & "' as valLetra,'" & PorIVA & "%' as Piva,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,CONCAT('Nº: ',egreso.egr_codigo,' - ',time_format(current_timestamp,'%H:%i'),persona.cat_p_codigo) as todo,FORMAT(egr_fecha,'dd-mm-yyyy') as fech,'" & cantid & "' as cants,for_ent_nombre,CONCAT(ciu_nombre,' ',persona.per_direccion2) as per_direccion2,COALESCE(CONCAT(p1.per_apellido,' ',p1.per_nombre),'') as gerente,COALESCE(CONCAT(p2.per_apellido,' ',p2.per_nombre),'') as director,IIF(p1.per_codigo IS NULL,'','Distribuidor:') as ger,IIF(p2.per_codigo IS NULL,'','Director:') as dir,egr_fechamod as fechamod, egr_usumod as usumod " & _
                 " FROM empresa INNER JOIN egreso ON empresa.emp_codigo=egreso.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND persona.for_pag_codigo_imp=forma_pago.for_pag_codigo " & _
                 " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                 " INNER JOIN det_egreso_c ON egreso.emp_codigo=det_egreso_c.emp_codigo AND egreso.tip_egr_codigo=det_egreso_c.tip_egr_codigo AND egreso.egr_codigo=det_egreso_c.egr_codigo " & _
                 " INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo" & _
                 " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
                 " WHERE empresa.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='FAC' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' " & _
                 " ORDER BY 1,mar_codigo,LEFT(gru_codigo,2),gru_nombre,nombre  "
    ElseIf strReporte = "rptPedido" Then
    
        ReDim ped(FormatoD0(strTipo)) As String
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        
        Me.Caption = "Pedidos - " & strNumero
        ped = Split(strNumero, ",")
        strSqlAux = " EXEC Sp_Drop_Table_if_Exist 'CDBPedido" & strUsuario & "' "
        clsConAUX.Ejecutar strSqlAux
        strSqlAux = " CREATE TABLE CDBPedido" & strUsuario & "( " & _
                   " emp_codigo char(3) NOT NULL default ''," & _
                   " ped_codigo decimal(14,0) NOT NULL default '0', " & _
                   " ped_cdb varchar(50) default NULL, " & _
                   " PRIMARY KEY  (emp_codigo,ped_codigo)) "
        clsConAUX.Ejecutar strSqlAux
        
        For jj = 0 To FormatoD0(strTipo) - 2
            
            strSqlAux = " INSERT INTO CDBPedido" & strUsuario & "(emp_codigo,ped_codigo,ped_cdb) " & _
                        " VALUES('" & strEmpresa & "','" & ped(jj) & "','" & Replace(code128$(ped(jj)), "'", "''") & "') "
            clsConAUX.Ejecutar strSqlAux
            
        Next jj
        
        GetSQL = " SELECT IIF(persona.for_pag_codigo='CONT',0,1) as orden,emp_nombre,ped_cdb AS CBF,tip_ped_nombre,pedido.ped_codigo,ped_usumod,ped_fecha as fech,CONCAT(persona.per_apellido,' ',persona.per_nombre) as nombC,persona.per_ruc, CONCAT(persona.per_telf,'/',persona.per_fax) telf,persona.per_celular,CONCAT(persona.per_direccion,' (',persona.per_direccion2,')')as dirC,ciu_nombre," & _
                 " CONCAT(ven_apellido,' ',ven_nombre) as nombV," & _
                 " CONCAT(GZ.per_apellido,' ',GZ.per_nombre) as nombG," & _
                 " CONCAT(DI.per_apellido,' ',DI.per_nombre) as nombD," & _
                 " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as EMPR," & _
                 " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as EJES," & _
                 " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5," & _
                 " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6," & _
                 " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7," & _
                 " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8," & _
                 " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9," & _
                 " ped_observacion as obs,for_pag_nombre, " & _
                 " det_pedido.dep_codigo,prd_ubica_linea,det_pedido.prd_codigo,prd_nombre,tal_nombre,col_nombre,det_ped_cant_pedida,det_ped_cant_entregada, SUM(exi_cantidad) as existen, IIF(SUM(exi_cantidad)>0,0,1) as exorden, "
        GetSQL = GetSQL & " IIF(LEN(CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')))>2,CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')))>2,CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')))>2,CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')))>2,CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')))>2,CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')))>2,CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')))>2,CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(DI.per_apellido,''),' ',COALESCE(DI.per_nombre,'')))>2,CONCAT(COALESCE(DI.per_apellido,''),' ',COALESCE(DI.per_nombre,''))," & _
                 " IIF(LEN(CONCAT(COALESCE(GZ.per_apellido,''),' ',COALESCE(GZ.per_nombre,'')))>2,CONCAT(COALESCE(GZ.per_apellido,''),' ',COALESCE(GZ.per_nombre,'')),''))))))))) as papa,COALESCE(pedRes.cantRes,0) as cantRes"
        GetSQL = GetSQL & " FROM pedido INNER JOIN empresa ON pedido.emp_codigo=empresa.emp_codigo" & _
                 " INNER JOIN CDBPedido" & strUsuario & " ON pedido.emp_codigo=CDBPedido" & strUsuario & ".emp_codigo " & _
                 " AND pedido.ped_codigo=CDBPedido" & strUsuario & ".ped_codigo " & _
                 " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo" & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                 " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo " & _
                 " AND COALESCE(IIF(persona.for_pag_codigo_imp='',NULL,persona.for_pag_codigo_imp),persona.for_pag_codigo)=forma_pago.for_pag_codigo " & _
                 " INNER JOIN vendedor ON pedido.emp_codigo=vendedor.emp_codigo AND IIF(pedido.ven_codigo='' or pedido.ven_codigo is null,persona.ven_codigo,pedido.ven_codigo)=vendedor.ven_codigo " & _
                 " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN talla ON producto.tal_codigo=talla.tal_codigo AND producto.emp_codigo=talla.emp_codigo INNER JOIN color ON producto.col_codigo=color.col_codigo AND producto.emp_codigo=color.emp_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                 " INNER JOIN existencia ON det_pedido.emp_codigo=existencia.emp_codigo AND det_pedido.dep_codigo=existencia.dep_codigo AND det_pedido.prd_codigo=existencia.prd_codigo "
        GetSQL = GetSQL & " LEFT JOIN ( SELECT det_pedido.emp_codigo,det_pedido.dep_codigo,det_pedido.prd_codigo,sum(det_pedido.det_ped_cant_pedida) as cantRes " & _
                 " FROM pedido INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_estado in (0,1) " & _
                 " AND det_pedido.prd_codigo not like 'PR-%' " & _
                 " GROUP BY det_pedido.emp_codigo,det_pedido.dep_codigo,det_pedido.prd_codigo " & _
                 " ) pedRes ON  det_pedido.emp_codigo=pedRes.emp_codigo " & _
                 " AND  det_pedido.dep_codigo=pedRes.dep_codigo " & _
                 " AND  det_pedido.prd_codigo=pedRes.prd_codigo"
        GetSQL = GetSQL & " LEFT JOIN persona GZ ON GZ.emp_codigo=persona.emp_codigo AND GZ.per_codigo=persona.per_codigo_ref AND GZ.per_es_gz=1 " & _
                 " LEFT JOIN persona DI ON DI.emp_codigo=persona.emp_codigo AND DI.per_codigo=persona.per_codigo_ref2 AND DI.per_es_di=1 " & _
                 " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                 " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                 " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                 " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                 " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                 " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                 " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                 " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                 " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                 " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                 " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                 " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                 " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                 " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo in (" & strNumero & ") " & _
                 " GROUP BY persona.for_pag_codigo,emp_nombre,ped_cdb,tip_ped_nombre,pedido.ped_codigo,ped_usumod,ped_fecha,persona.per_apellido,persona.per_nombre,persona.per_ruc,persona.per_telf,persona.per_fax,persona.per_celular, persona.per_direccion, persona.per_direccion2,ciu_nombre,ven_apellido,ven_nombre,GZ.per_apellido, GZ.per_nombre, DI.per_apellido, DI.per_nombre, EMP.per_apellido,EMP.per_nombre, EJE.per_apellido, EJE.per_nombre, N5.per_apellido, N5.per_nombre, N6.per_apellido, N6.per_nombre, N7.per_apellido, N7.per_nombre,N8.per_apellido, N8.per_nombre, N9.per_apellido, N9.per_nombre, ped_observacion,for_pag_nombre,  det_pedido.dep_codigo,prd_ubica_linea,det_pedido.prd_codigo,prd_nombre,tal_nombre,col_nombre,det_ped_cant_pedida,det_ped_cant_entregada, pedRes.cantRes,mar_codigo,producto.gru_codigo,gru_nombre " & _
                 " ORDER BY orden,nombG,papa,pedido.ped_codigo,prd_ubica_linea,exorden,mar_codigo,LEFT(producto.gru_codigo,2),gru_nombre,prd_nombre "
                 '" ORDER BY pedido.ped_codigo,exorden,mar_codigo,LEFT(producto.gru_codigo,2),gru_nombre,prd_nombre  "
    ElseIf strReporte = "rptPreFactura" Then
        Me.Caption = "PreFactura - Pedido No. " & strNumero
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = " SELECT SUM(det_ped_cant_pedida) as cantP,SUM(det_ped_cant_confirmada) as cantE,SUM(det_ped_cant_confirmada*det_ped_precio)+'0' as suman,SUM(ROUND(det_ped_cant_confirmada*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2)) as Descu,par_numero as IVA," & _
                    " ROUND((SUM(det_ped_cant_confirmada*det_ped_precio) - SUM(ROUND(det_ped_cant_confirmada*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2))) * (100.00+par_numero)/100.00,2) as total," & _
                    " pedido.ped_codigo " & _
                    " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo AND pedido.per_codigo=persona.per_codigo AND persona.cat_p_tipo='C' " & _
                    " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                    " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo " & _
                    " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
                    " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
                    " LEFT JOIN vendedor ON pedido.emp_codigo=vendedor.emp_codigo AND pedido.ven_codigo=vendedor.ven_codigo " & _
                    " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
                    " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
                    " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
                    " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
                    " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
                    " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_codigo='" & strNumero & "' " & _
                    " GROUP BY pedido.ped_codigo,parametro.par_numero "
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("total"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("total") * 100#)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100.00 "
        Set tNum2Text = Nothing
           
        GetSQL = " SELECT 1,pedido.ped_codigo, ped_fecha,DATEADD(d,for_pag_tiempo,ped_fecha) as vence, CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
             " ped_observacion, est_descripcion, tipo_fac_descripcion, persona.per_codigo, cot_codigo,persona.per_ruc,for_ent_nombre, " & _
             " pedido.ven_codigo,persona.per_observacion,pedido.tar_cre_codigo,tar_cre_nombre,persona.for_pag_codigo,for_pag_nombre,persona.per_direccion,CONCAT(persona.per_telf,' ',persona.per_fax) as per_telf," & _
             " persona.per_sec_publico,persona.per_siniva,dep_codigo, det_pedido.prd_codigo as codprod, prd_nombre as nomprod, det_ped_cant_pedida, det_ped_cant_confirmada, det_ped_precio as precio, det_ped_dcto,det_ped_cant_confirmada as cantidadE,det_ped_cant_pedida as cantidadP,CONCAT(ven_apellido,' ',ven_nombre) as ven, " & _
             " ROUND((det_ped_cant_confirmada * det_ped_precio),2) as total,ROUND(ROUND(det_ped_cant_confirmada * det_ped_precio,2)-ROUND(ROUND(det_ped_cant_confirmada*det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(persona.per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(persona.per_dcto,0))/100.00,2),2) as utot," & _
             " ROUND(ROUND(det_ped_cant_confirmada*det_ped_precio,2)*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(persona.per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(persona.per_dcto,0))/100.00,2) as Descu,prd_iva,'" & strValor & "' as valLetra,'" & strAsiento & "' as Piva,producto.mar_codigo,producto.gru_codigo,gru_nombre,det_pedido.prd_codigo," & FormatoD2(strTipo) & " as reca,COALESCE(CONCAT(p1.per_apellido,' ',p1.per_nombre),'') as gerente,COALESCE(CONCAT(p2.per_apellido,' ',p2.per_nombre),'') as director,IIF(p1.per_codigo IS NULL,'','Distribuidor:') as ger,IIF(p2.per_codigo IS NULL,'','Director:') as dir,ped_usumod as usumod,ped_fechamod as fechamod " & _
             " FROM pedido INNER JOIN est_pedido ON est_pedido.est_codigo = pedido.ped_estado " & _
             " INNER JOIN det_pedido ON pedido.ped_codigo = det_pedido.ped_codigo " & _
             " AND pedido.emp_codigo = det_pedido.emp_codigo INNER JOIN producto " & _
             " ON det_pedido.emp_codigo = producto.emp_codigo AND det_pedido.prd_codigo = producto.prd_codigo " & _
             " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
             " INNER JOIN persona ON pedido.per_codigo = persona.per_codigo AND pedido.emp_codigo = persona.emp_codigo " & _
             " INNER JOIN tipo_factura ON pedido.tipo_fac_codigo = tipo_factura.tipo_fac_codigo " & _
             " INNER JOIN forma_pago ON persona.emp_codigo=forma_pago.emp_codigo AND persona.for_pag_codigo=forma_pago.for_pag_codigo " & _
             " INNER JOIN forma_entrega ON persona.emp_codigo=forma_entrega.emp_codigo AND persona.for_ent_codigo=forma_entrega.for_ent_codigo " & _
             " LEFT JOIN vendedor ON pedido.emp_codigo=vendedor.emp_codigo AND pedido.ven_codigo=vendedor.ven_codigo " & _
             " LEFT JOIN tarjeta_credito ON pedido.emp_codigo = tarjeta_credito.emp_codigo AND pedido.tar_cre_codigo = tarjeta_credito.tar_cre_codigo " & _
             " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
             " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo "
        GetSQL = GetSQL & " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
             " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
             " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
             " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND pedido.ped_codigo='" & strNumero & "' " & _
             "  " & _
             " "
        GetSQL = GetSQL & " ORDER BY 1,mar_codigo,LEFT(grupo.gru_codigo,2),gru_nombre,prd_codigo "
'        GetSQL = GetSQL & " UNION  SELECT 2,'' as ped_codigo,'' as ped_fecha,'' as vence, '' as per, " & _
'             " '' as ped_observacion,'' as est_descripcion,'' as tipo_fac_descripcion,'' as per_codigo,'' as cot_codigo,'' as per_ruc,'' as for_en_nombre, " & _
'             " '' as ven_codigo,'' as per_observacion,'' as tar_cre_codigo,'' as tar_cre_nombre,'' as for_pag_codigo,'' as for_pag_nombre,'' as per_direccion,'' as per_telf," & _
'             " '' as per_sec_publico,'' as per_siniva,'' as dep_codigo,cod as codprod, prod as nomprod, '' as det_ped_cant_pedida,'' as det_ped_cant_confirmada, prec as precio,0.00 as det_ped_dcto, 1 as cantidadE,1 as cantidadP,'' as ven, " & _
'             " 0.00 as total,prec as utot,0.00 as Descu,'' as prd_iva,'" & strValor & "' as valLetra,'" & strAsiento & "' as Piva,'' as mar_codigo,'' as gru_codigo,'' as gru_nombre,'' as prd_codigo," & FormatoD2(strTipo) & " as reca,'' as gerente,'' as director,'' as ger,'' as dir,'' as usumod,'' as fechamod " & _
'             " FROM recs" & strNumero & " " & _
'             " ORDER BY 1,mar_codigo,LEFT(gru_codigo,2),gru_nombre,prd_codigo "
          
        
    ElseIf strReporte = "rptGuiaRemision" Then
        Me.Caption = "Guia de Remisión - " & strNumero
        GetSQL = " SELECT egreso.egr_codigo,emp_direccion,CONCAT(per_apellido,' ',per_nombre) as per, egreso.per_codigo,egr_observacion," & _
                 " per_ruc,per_direccion,per_telf,per_fax,ciu_nombre,egr_fecha,COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as vendedor,vendedor.ven_codigo as ven," & _
                 " gru_nombre as nombre, SUM(det_egr_cantidad) as cantidad," & _
                 " emp_nombre,tip_egr_nombre,egr_factura " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo " & _
                 " INNER JOIN det_egreso_ubicacion det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,5)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo " & _
                 " LEFT JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='" & strTipo & "' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' " & _
                 " GROUP BY grupo.gru_codigo " & _
                 " ORDER BY grupo.gru_codigo "
    ElseIf strReporte = "rptGuiaRemision2" Then
        Me.Caption = "Guia de Remisión - " & strNumero
        GetSQL = " SELECT egreso.egr_codigo,emp_direccion,CONCAT(per_apellido,' ',per_nombre) as per, egreso.per_codigo,unidad.uni_codigo,egr_observacion," & _
                 " per_ruc,per_direccion,per_telf,per_fax,ciu_nombre,egr_fecha,COALESCE(CONCAT(ven_apellido,' ',ven_nombre),'') as vendedor,vendedor.ven_codigo as ven," & _
                 " det_egreso.prd_codigo as prd_codigo,prd_nombre as nombre,det_egreso.prd_ubica_linea, ROUND(det_egr_cantidad,2) as cantidad,uni_nombre,det_egr_precio," & _
                 " det_egr_precio*det_egr_cantidad-ROUND((det_egr_dcto/det_egr_cantidad*ROUND(det_egr_cantidad,2)),2) as utot,ROUND((det_egr_dcto/det_egr_cantidad*ROUND(det_egr_cantidad,2)),2) as det_egr_dcto,egr_subtotal,egr_dcto,egr_subtotal_o,egr_impuesto,egr_total, emp_nombre,tip_egr_nombre,egr_factura " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo " & _
                 " INNER JOIN det_egreso_ubicacion det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo AND det_egreso.prd_codigo=producto.prd_codigo " & _
                 " INNER JOIN unidad ON producto.emp_codigo=unidad.emp_codigo AND producto.uni_codigo=unidad.uni_codigo " & _
                 " INNER JOIN grupo ON LEFT(producto.gru_codigo,8)=grupo.gru_codigo AND producto.emp_codigo=grupo.emp_codigo "
        GetSQL = GetSQL & " LEFT JOIN forma_pago ON egreso.emp_codigo=forma_pago.emp_codigo AND egreso.for_pag_codigo=forma_pago.for_pag_codigo " & _
                 " LEFT JOIN vendedor ON egreso.emp_codigo=vendedor.emp_codigo AND egreso.ven_codigo=vendedor.ven_codigo " & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.tip_egr_codigo='" & strTipo & "' " & _
                 " AND egreso.egr_codigo='" & strNumero & "' " & _
                 " ORDER BY prd_ubica_linea,producto.mar_codigo,LEFT(producto.gru_codigo,2),grupo.gru_nombre,producto.prd_nombre "
    ElseIf strReporte = "rptIngresoMercaderia" Then
        Me.Caption = "Ingreso de Mercaderia - " & strNumero
        GetSQL = " SELECT ingreso.ing_codigo, COALESCE(CONCAT(per_apellido,' ',per_nombre),'') as nombC, ing_observacion as obs, ing_fecha as fech,tip_ing_nombre as tdoc, emp_nombre, " & _
                 " CONCAT(COALESCE(ing_factura,''), ' / ',COALESCE(ing_serie,''),' - ',COALESCE(ing_numero,'')) as factura,det_ingreso.prd_codigo,det_ingreso.dep_codigo," & _
                 " CONCAT(LEFT(prd_nombre,40),' -',clc_nombre) as prd_nombre,COALESCE(ubi.ubica,'--')as ubica,COALESCE(det_con_mer_cantidad,det_ing_cantidad) as cant,ing_fechamod as fechamod, ing_usumod as usumod " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo AND producto.clc_codigo=coleccion.clc_codigo" & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo "
        GetSQL = GetSQL & " LEFT JOIN (SELECT det_contenedor_mercaderia.emp_codigo,tip_mov_codigo,mov_codigo,dep_codigo,det_contenedor_mercaderia.prd_codigo,CONCAT(contenedor_mercaderia.ubi_bod_codigo,'-',contenedor_mercaderia.con_mer_codigo) as ubica,det_con_mer_cantidad " & _
                 " FROM det_contenedor_mercaderia INNER JOIN contenedor_mercaderia ON det_contenedor_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE det_contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_mov_codigo='" & strTipo & "' " & _
                 " AND det_con_mer_cantidad!=0 " & _
                 " AND mov_codigo in (" & strNumero & ")) ubi ON det_ingreso.emp_codigo=ubi.emp_codigo " & _
                 " AND det_ingreso.tip_ing_codigo=ubi.tip_mov_codigo " & _
                 " AND det_ingreso.ing_codigo=ubi.mov_codigo " & _
                 " AND det_ingreso.dep_codigo=ubi.dep_codigo " & _
                 " AND det_ingreso.prd_codigo=ubi.prd_codigo " & _
                 " LEFT JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo " & _
                 " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='" & strTipo & "' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY ubica,prd_nombre"
    ElseIf strReporte = "rptDetalleAdjunto" Then
        Me.Caption = "Ingreso de Mercaderia - " & strNumero
        GetSQL = " SELECT ingreso.ing_codigo, COALESCE(CONCAT(per_apellido,' ',per_nombre),'') as nombC, ing_observacion as obs, ing_fecha as fech,tip_ing_nombre as tdoc, emp_nombre, " & _
                 " CONCAT(COALESCE(ing_factura,''), ' / ',COALESCE(ing_serie,''),' - ',COALESCE(ing_numero,'')) as factura,ing_subtotal,ing_subtotal_o , ing_impuesto, ing_total,det_ingreso.prd_codigo,det_ingreso.dep_codigo,CONCAT(LEFT(prd_nombre,40),' -',clc_nombre) as prd_nombre,det_ing_cantidad as cant,det_ing_precio,det_ing_dcto,det_ing_precio*det_ing_cantidad-det_ing_dcto as utot, ing_fechamod as fechamod, ing_usumod as usumod " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN coleccion ON producto.emp_codigo=coleccion.emp_codigo AND producto.clc_codigo=coleccion.clc_codigo" & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo " & _
                 " LEFT JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo " & _
                 " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='" & strTipo & "' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'"
    ElseIf strReporte = "rptRecepcionMercaderia" Then
        Me.Caption = "Recepción de Mercadería - " & strNumero
        GetSQL = " SELECT ingreso.ing_codigo, COALESCE(CONCAT(per_apellido,' ',per_nombre),'') as nombC, ing_observacion as obs, ing_fecha as fech,tip_ing_nombre as tdoc, emp_nombre,mar_codigo as grupo,RIGHT(producto.prd_codigo,LEN(producto.prd_codigo)-4) as codigo,uni_codigo," & _
                 " CONCAT(COALESCE(ing_factura,''), ' / ',COALESCE(ing_serie,''),' - ',COALESCE(ing_numero,'')) as factura,det_ingreso.prd_codigo,dep_codigo,prd_nombre,det_ing_cantidad as cant,ing_factura " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo " & _
                 " LEFT JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo " & _
                 " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='" & strTipo & "' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'"
    ElseIf strReporte = "rptEgresoMercaderia" Then
        Me.Caption = "Egreso de Mercaderia - " & strNumero
        GetSQL = " SELECT egreso.egr_codigo, COALESCE(CONCAT(per_apellido,' ',per_nombre),'') as nombC, egr_observacion as obs, egr_fecha as fech,tip_egr_nombre as tdoc, emp_nombre, " & _
                 " CONCAT(COALESCE(egr_factura,''), ' / ',COALESCE(egr_serie,''),' - ',COALESCE(egr_numero,'')) as factura,det_egreso_ubicacion.prd_codigo,det_egreso_ubicacion.dep_codigo," & _
                 " prd_nombre,det_egreso_ubicacion.prd_ubica_linea as ubica,det_egr_cantidad as cant,egr_fechamod as fechamod, egr_usumod as usumod " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_egreso_ubicacion ON egreso.emp_codigo=det_egreso_ubicacion.emp_codigo AND egreso.tip_egr_codigo=det_egreso_ubicacion.tip_egr_codigo AND egreso.egr_codigo=det_egreso_ubicacion.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso_ubicacion.prd_codigo=producto.prd_codigo AND det_egreso_ubicacion.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona ON egreso.emp_codigo=persona.emp_codigo AND egreso.per_codigo=persona.per_codigo " & _
                 " WHERE egreso.egr_codigo = '" & strNumero & "' " & _
                 " AND egreso.tip_egr_codigo='" & strTipo & "' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY ubica,prd_nombre"
    ElseIf strReporte = "rptTransformacionMercaderia" Then
        Me.Caption = "Transformación de Mercaderia - " & strNumero
        GetSQL = " SELECT '1' as orde,egreso.egr_codigo, egr_observacion as obs, egr_fecha as fech,tip_egr_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(egr_factura,'') as factura,det_egreso.prd_codigo,det_egreso.dep_codigo,prd_nombre," & _
                 " COALESCE(det_egreso.prd_ubica_linea,'--') as ubica," & _
                 " det_egr_cantidad as cant,det_egr_precio as prec " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_egreso_ubicacion det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.prd_codigo=producto.prd_codigo AND det_egreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo "
        GetSQL = GetSQL & " WHERE egreso.egr_codigo = '" & strNumero & "' " & _
                 " AND egreso.tip_egr_codigo='ETN' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "'" & _
                 " UNION "
        GetSQL = GetSQL & " SELECT '2' as orde,ingreso.ing_codigo, ing_observacion as obs, ing_fecha as fech,tip_ing_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(ing_factura,'') as factura,det_ingreso.prd_codigo,det_ingreso.dep_codigo,prd_nombre," & _
                 " COALESCE(det_ingreso.prd_ubica_linea,'--') as ubica," & _
                 " det_ing_cantidad as cant,det_ing_precio as prec " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso_ubicacion det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo "
        GetSQL = GetSQL & " LEFT JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo "
        GetSQL = GetSQL & " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='ITN' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY orde,ubica,prd_codigo "
    ElseIf strReporte = "rptTransferencia" Then
        Me.Caption = "Transferencia de Mercaderia - " & strNumero
        GetSQL = " SELECT '0' as orde,egreso.egr_codigo, egr_observacion as obs, egr_fecha as fech,tip_egr_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(egr_factura,'') as factura,det_egreso.prd_codigo,det_egreso.dep_codigo,prd_nombre," & _
                 " det_egreso.prd_ubica_linea as ubica," & _
                 " ROUND(det_egr_cantidad,2) as cant,det_egr_precio as prec " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_egreso_ubicacion det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.prd_codigo=producto.prd_codigo AND det_egreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo " & _
                 " WHERE egreso.egr_codigo = '" & strNumero & "' " & _
                 " AND egreso.tip_egr_codigo='ETR' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "'"
        GetSQL = GetSQL & " UNION " & _
                 " SELECT '2' as orde,ingreso.ing_codigo, ing_observacion as obs, ing_fecha as fech,tip_ing_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(ing_factura,'') as factura,det_ingreso.prd_codigo,det_ingreso.dep_codigo,prd_nombre," & _
                 " det_ingreso.prd_ubica_linea as ubica," & _
                 " det_ing_cantidad as cant,det_ing_precio as prec " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso_ubicacion det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo " & _
                 " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='ITR' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'"
        If MsgBox("Ordenado por pasillo?" & vbNewLine & "Si responde que NO se ordenará por Referencia", vbQuestion + vbYesNo, "Transferencia") = vbYes Then
            GetSQL = GetSQL & " ORDER BY orde,dep_codigo,ubica,prd_nombre "
        Else
            GetSQL = GetSQL & " ORDER BY orde,prd_nombre,ubica "
        End If
    ElseIf strReporte = "rptCambioProducto" Then
        Me.Caption = "Cambio de Productos - " & strNumero
        GetSQL = " SELECT '1' as orde,ingreso.ing_codigo as codigo, ing_observacion as obs, ing_fecha as fecha,tip_ing_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(ing_factura,'') as factura,det_ingreso.prd_codigo,det_ingreso.dep_codigo,prd_nombre,ubica,det_ing_cantidad as cant,det_ing_precio as prec " & _
                 " FROM ingreso INNER JOIN empresa ON ingreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                 " INNER JOIN producto ON det_ingreso.prd_codigo=producto.prd_codigo AND det_ingreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_ingreso ON ingreso.emp_codigo=tipo_ingreso.emp_codigo AND ingreso.tip_ing_codigo=tipo_ingreso.tip_ing_codigo "
        GetSQL = GetSQL & " LEFT JOIN (SELECT det_contenedor_mercaderia.emp_codigo,tip_mov_codigo,mov_codigo,dep_codigo,det_contenedor_mercaderia.prd_codigo,CONCAT(contenedor_mercaderia.ubi_bod_codigo,'-',contenedor_mercaderia.con_mer_codigo) as ubica,det_con_mer_cantidad" & _
                 " FROM det_contenedor_mercaderia INNER JOIN contenedor_mercaderia ON det_contenedor_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE det_contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_mov_codigo='ica' " & _
                 " AND det_con_mer_cantidad!=0 " & _
                 " AND mov_codigo in (" & strNumero & ")) ubi ON det_ingreso.emp_codigo=ubi.emp_codigo " & _
                 " AND det_ingreso.tip_ing_codigo=ubi.tip_mov_codigo " & _
                 " AND det_ingreso.ing_codigo=ubi.mov_codigo " & _
                 " AND det_ingreso.dep_codigo=ubi.dep_codigo " & _
                 " AND det_ingreso.prd_codigo=ubi.prd_codigo "
        GetSQL = GetSQL & " WHERE ingreso.ing_codigo = '" & strNumero & "' " & _
                 " AND ingreso.tip_ing_codigo='ICA' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "'" & _
                 " UNION " & _
                 " SELECT '2' as orde,egreso.egr_codigo as codigo, egr_observacion as obs, egr_fecha as fecha,tip_egr_nombre as tdoc, emp_nombre, " & _
                 " COALESCE(egr_factura,'') as factura,det_egreso.prd_codigo,det_egreso.dep_codigo,prd_nombre,ubica,det_egr_cantidad as cant,det_egr_precio as prec " & _
                 " FROM egreso INNER JOIN empresa ON egreso.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo AND egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " INNER JOIN producto ON det_egreso.prd_codigo=producto.prd_codigo AND det_egreso.emp_codigo=producto.emp_codigo " & _
                 " INNER JOIN tipo_egreso ON egreso.emp_codigo=tipo_egreso.emp_codigo AND egreso.tip_egr_codigo=tipo_egreso.tip_egr_codigo "
        GetSQL = GetSQL & " LEFT JOIN (SELECT det_contenedor_mercaderia.emp_codigo,tip_mov_codigo,mov_codigo,dep_codigo,det_contenedor_mercaderia.prd_codigo,CONCAT(contenedor_mercaderia.ubi_bod_codigo,'-',contenedor_mercaderia.con_mer_codigo) as ubica,det_con_mer_cantidad" & _
                 " FROM det_contenedor_mercaderia INNER JOIN contenedor_mercaderia ON det_contenedor_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
                 " WHERE det_contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_mov_codigo='ECA' " & _
                 " AND det_con_mer_cantidad!=0 " & _
                 " AND mov_codigo in (" & strNumero & ")) ubi ON det_egreso.emp_codigo=ubi.emp_codigo " & _
                 " AND det_egreso.tip_egr_codigo=ubi.tip_mov_codigo " & _
                 " AND det_egreso.egr_codigo=ubi.mov_codigo " & _
                 " AND det_egreso.dep_codigo=ubi.dep_codigo " & _
                 " AND det_egreso.prd_codigo=ubi.prd_codigo "
        GetSQL = GetSQL & " WHERE egreso.egr_codigo = '" & strNumero & "' " & _
                 " AND egreso.tip_egr_codigo='ECA' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "'" & _
                 " ORDER BY orde "
    ElseIf strReporte = "rptRetencion" Then
        Dim i As Integer
        Me.Caption = "Comprobante de Retención"
        
        clsSql1.Inicializar AdoConn, AdoConnMaster
        clsSQL2.Inicializar AdoConn, AdoConnMaster
       
       clsConAUX.Inicializar AdoConn, AdoConnMaster
        
        strSql = " SELECT COALESCE(doc_ele_claveacceso,'') as doc_ele_claveacceso " & _
                 " FROM doc_electronico " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND doc_ele_coddoc='01' " & _
                 " ANd doc_ele_codigo='" & strNumero & "'"
        
        clsConAUX.Ejecutar strSql
        
        If clsConAUX.adorec_Def.RecordCount > 0 Then
            CDEClaveAcceso = Replace(code128$(clsConAUX.adorec_Def("doc_ele_claveacceso")), "'", "''")
        Else
            CDEClaveAcceso = ""
        End If
       
'''        strSql = " CREATE TABLE Detalle" & strUsuario & " (" & _
'''                 " item bigint(2) NOT NULL default '0', " & _
'''                 " cue_p_c_codigo decimal(11,0) NOT NULL default '0', " & _
'''                 " cue_p_c_tipo char(1) NOT NULL default '', " & _
'''                 " ret_codigo_f decimal(4,1) NOT NULL default '-1'," & _
'''                 " det_com_ret_valor_f decimal(14,2) NOT NULL default '0.00', " & _
'''                 " det_com_ret_porcentaje_f decimal(5,2) NOT NULL default '0.00', " & _
'''                 " total_f double(21,4) default NULL, " & _
'''                 " ret_codigo char(3) NOT NULL default '') "
'''        clsSql1.Ejecutar strSql
'''        strSql = " DELETE FROM Detalle" & strUsuario
'''        clsSql1.Ejecutar strSql
'''
'''        strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,det_comp_ret.ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_comp_ret.ret_codigo " & _
'''                 " FROM ((cuenta_p_c INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo) " & _
'''                 " INNER JOIN det_comp_ret ON comprobante_retencion.emp_codigo=det_comp_ret.emp_codigo AND comprobante_retencion.cue_p_c_codigo=det_comp_ret.cue_p_c_codigo AND comprobante_retencion.cue_p_c_tipo=det_comp_ret.cue_p_c_tipo)" & _
'''                 " INNER JOIN retencion ON det_comp_ret.ret_codigo=retencion.ret_codigo AND det_comp_ret.emp_codigo=retencion.emp_codigo " & _
'''                 " AND comprobante_retencion.com_ret_fecha BETWEEN retencion.ret_fechaini AND retencion.ret_fechafin " & _
'''                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
'''                 " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "' " & _
'''                 " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
'''                 " AND det_comp_ret.ret_codigo NOT IN ('1','2','3') "
'''        clsSql1.Ejecutar strSql
'''        Dim a As Integer
'''        a = FormatoD0(clsSql1.adorec_Def.RecordCount)
'''
'''        If clsSql1.adorec_Def.EOF = False Then
'''           While Not clsSql1.adorec_Def.EOF
'''                strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo) " & _
'''                         " VALUES('" & i & "','" & strNumero & "','" & strTipo & "','" & clsSql1.adorec_Def("ret_codigo") & "','" & clsSql1.adorec_Def("det_com_ret_valor") & "','" & clsSql1.adorec_Def("det_com_ret_porcentaje") & "','" & Format((clsSql1.adorec_Def("det_com_ret_valor") * clsSql1.adorec_Def("det_com_ret_porcentaje")) / 100.00, "###0.00") & "','" & clsSql1.adorec_Def("ret_codigo") & "')"
'''                clsSQL2.Ejecutar strSql
'''                clsSql1.adorec_Def.MoveNext
'''           Wend
'''
'''            If a = 1 Then
'''                i = i + 1
'''                strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
'''                        " VALUES('" & i & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
'''               clsSQL2.Ejecutar strSql
'''            End If
'''        Else
'''''           strSql = " INSERT INTO Detalle(item,cue_p_c_codigo,cue_p_c_tipo)" & _
'''''                    " VALUES('" & i & "','" & strNumero & "','" & strTipo & "')"
'''            strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
'''                     " VALUES('" & i & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
'''           clsSQL2.Ejecutar strSql
'''           clsSQL2.Ejecutar strSql
'''        End If
'''
'''
'''
'''        strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,det_comp_ret.ret_codigo,det_com_ret_valor,det_com_ret_porcentaje,det_comp_ret.ret_codigo " & _
'''                 " FROM ((cuenta_p_c INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo) " & _
'''                 " INNER JOIN det_comp_ret ON comprobante_retencion.emp_codigo=det_comp_ret.emp_codigo AND comprobante_retencion.cue_p_c_codigo=det_comp_ret.cue_p_c_codigo AND comprobante_retencion.cue_p_c_tipo=det_comp_ret.cue_p_c_tipo)" & _
'''                 " INNER JOIN retencion ON det_comp_ret.ret_codigo=retencion.ret_codigo AND det_comp_ret.emp_codigo=retencion.emp_codigo " & _
'''                 " AND comprobante_retencion.com_ret_fecha BETWEEN retencion.ret_fechaini AND retencion.ret_fechafin " & _
'''                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
'''                 " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "' " & _
'''                 " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
'''                 " AND det_comp_ret.ret_codigo IN ('1','2','3') "
'''        clsSql1.Ejecutar strSql
'''        If clsSql1.adorec_Def.EOF = False Then
'''          BuscarPorc 30, 2
'''          BuscarPorc 70, 4
'''          BuscarPorc 100, 6
'''        Else
'''            For i = 0 To 5
'''              strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
'''                       " VALUES('" & i + 2 & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
'''              clsSQL2.Ejecutar strSql
'''            Next i
'''        End If
'''
'''        GetSQL = " SELECT detalle" & strUsuario & ".ret_codigo as item,CONCAT(com_ret_serie,' - ',com_ret_numero) as com_ret,com_ret_fecha,year(com_ret_fecha) as ejercicio,CONCAT(per_apellido,' ',per_nombre) as per, " & _
'''                " per_ruc,per_direccion,tip_doc_cue_descripcion,cue_p_c_serie,cue_p_c_numero,concat(cue_p_c_serie,'-',LPAD(cue_p_c_numero,9,'0')) as numero," & _
'''                " detalle" & strUsuario & ".ret_codigo,Detalle" & strUsuario & ".ret_codigo_f,IIF(Detalle" & strUsuario & ".det_com_ret_valor_f=0,'',ROUND(Detalle" & strUsuario & ".det_com_ret_valor_f,2)) as valor,IIF(Detalle" & strUsuario & ".det_com_ret_porcentaje_f=0,'',ROUND(Detalle" & strUsuario & ".det_com_ret_porcentaje_f,2)) as porcen,IIF(Detalle" & strUsuario & ".total_f=0,'',ROUND(Detalle" & strUsuario & ".total_f,2)) as total, " & _
'''                " '" & strAsiento & "' as concepto,cue_p_c_egr_codigo,com_ret_fecha as dia,com_ret_fecha as mes,com_ret_fecha as anio,cue_p_c_fechaemision,cue_p_c_autorizacion,tip_doc_cue_descripcion as tip_doc_cue_nombre " & _
'''                " FROM Detalle" & strUsuario & " LEFT JOIN cuenta_p_c ON Detalle" & strUsuario & ".cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND Detalle" & strUsuario & ".cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo LEFT JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
'''                " " & _
'''                " LEFT JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo " & _
'''                " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
'''                " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
'''                " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
'''                " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "' ORDER by detalle" & strUsuario & ".item "
        GetSQL = " SELECT '" & CDEClaveAcceso & "' AS CDEClaveAcceso,emp_nombre,emp_direccion,emp_telf,emp_ruc,doc_ele_claveacceso,doc_ele_autorizacion,com_ret_serie,com_ret_numero,com_ret_fecha,cue_p_c_fechapropuesta,CONCAT(per_apellido,' ',per_nombre) as per, " & _
                 " per_ruc,per_direccion,tip_doc_cue_descripcion,cue_p_c_serie,cue_p_c_numero,ret_nombre," & _
                 " det_comp_ret.ret_codigo,det_com_ret_valor,det_com_ret_porcentaje," & _
                 " ROUND(det_com_ret_valor * det_com_ret_porcentaje / 100.00, 2) As rtot," & _
                 " CONCAT('Estimados. De acuerdo con la información registrada en nuestro sistema, en tu  mail: - ',persona.per_email,' - recibirás tus documentos electrónicos autorizada por el SRI, según las nueva ley en vigencia. Si no tienes actualizados tus datos comunicate al 1800CATALOGOS para pedir esta actualización') as mensaje, " & _
                 " IIF(doc_ele_autorizacion='' OR doc_ele_autorizacion is null,'DOCUMENTO SIN VALIDEZ TRIBUTARIA','') as mensaje2 " & _
                 " FROM empresa INNER JOIN cuenta_p_c ON empresa.emp_codigo=cuenta_p_c.emp_codigo " & _
                 " INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
                 " INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                 " INNER JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                 " INNER JOIN det_comp_ret ON comprobante_retencion.emp_codigo=det_comp_ret.emp_codigo AND comprobante_retencion.cue_p_c_codigo=det_comp_ret.cue_p_c_codigo AND comprobante_retencion.cue_p_c_tipo=det_comp_ret.cue_p_c_tipo " & _
                 " INNER JOIN retencion ON det_comp_ret.emp_codigo=retencion.emp_codigo AND det_comp_ret.ret_codigo=retencion.ret_codigo " & _
                 " LEFT JOIN doc_electronico ON cuenta_p_c.emp_codigo=doc_electronico.emp_codigo AND cuenta_p_c.cue_p_c_codigo=doc_electronico.doc_ele_codigo " & _
                 " AND doc_electronico.doc_ele_coddoc='07' " & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
                 " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "'" & _
                 "" '" AND ROUND(det_com_ret_valor * det_com_ret_porcentaje / 100.00, 2)!=ROUND(0.00,2)"
        
    ElseIf strReporte = "rptReciboCaja" Then
        Me.Caption = "Recibo de Caja"
        GetSQL = " SELECT doc_pago.doc_pag_codigo,doc_pag_fecha_recepcion,doc_pag_fecha_doc,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, " & _
                 " CONCAT(cue_p_c_serie,'-',FORMAT(cue_p_c_numero,'0000000')) as numero,IIF(doc_pago.tip_doc_pag_codigo IS NULL OR doc_pago.tip_doc_pag_codigo='','EFECTIVO',tip_doc_pag_nombre) as pag_nombre," & _
                 " doc_pag_numero,doc_pag_valor as valor,COALESCE(doc_pag_observacion,'') as observacion,IIF(doc_pago.ban_codigo IS NULL OR doc_pago.ban_codigo='','EFECTIVO',ban_nombre) as banco," & _
                 " pag_monto as monto,CONCAT(usu_nombre,' ',usu_apellido) as usu,cue_p_c_valor,COALESCE(CONCAT(p1.per_apellido,' ',p1.per_nombre),'') as gerente,COALESCE(CONCAT(p2.per_apellido,' ',p2.per_nombre),'') as director " & _
                 " FROM doc_pago INNER JOIN pago ON pago.emp_codigo=doc_pago.emp_codigo " & _
                 " AND doc_pago.doc_pag_codigo=pago.doc_pag_codigo " & _
                 " INNER JOIN cuenta_p_c ON cuenta_p_c.emp_codigo=pago.emp_codigo " & _
                 " AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                 " AND cuenta_p_c.cue_p_c_tipo='C' " & _
                 " INNER JOIN persona ON doc_pago.emp_codigo=persona.emp_codigo AND doc_pago.per_codigo=persona.per_codigo " & _
                 " LEFT JOIN banco ON banco.ban_codigo=doc_pago.ban_codigo " & _
                 " LEFT JOIN tipo_doc_pago ON tipo_doc_pago.tip_doc_pag_codigo=doc_pago.tip_doc_pag_codigo " & _
                 " INNER JOIN usuario ON doc_pago.doc_pag_usumod=usuario.usu_codigo " & _
                 " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
                 " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
                 " WHERE doc_pago.emp_codigo='" & strEmpresa & "' " & _
                 " AND doc_pago.doc_pag_codigo='" & strNumero & "' "
    ElseIf strReporte = "rptLiquidacionCompras" Then
        Me.Caption = "Liquidación de Compras y Servicios"
        ' transformar numeros a letras
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = " SELECT cue_p_c_valor " & _
                 " FROM cuenta_p_c " & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
                 " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "'"
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("cue_p_c_valor"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("cue_p_c_valor") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        Set tNum2Text = Nothing
        GetSQL = " SELECT CONCAT(cue_p_c_serie,' - ',cue_p_c_numero) as num, cuenta_p_c.per_codigo, CONCAT(per_apellido,' ',per_nombre) as per, per_direccion, " & _
                 " per_telf, per_ruc, ciu_nombre, cue_p_c_fechaemision as fecha, cue_p_c_descripcion, (cue_p_c_st_prod+cue_p_c_st_serv) as SubTotConIVA,emp_nombre, " & _
                 " cue_p_c_st_cero as SubTotSinIVA,cue_p_c_iva as IVA, cue_p_c_valor as Total, '" & strValor & "' as TotalLetras " & _
                 " FROM cuenta_p_c INNER JOIN persona ON cuenta_p_c.emp_codigo=persona.emp_codigo AND cuenta_p_c.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo INNER JOIN empresa ON cuenta_p_c.emp_codigo=empresa.emp_codigo " & _
                 " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND cuenta_p_c.cue_p_c_tipo='" & strTipo & "' " & _
                 " AND cuenta_p_c.cue_p_c_codigo='" & strNumero & "'"
    ElseIf strReporte = "rptNotaCredito" Or strReporte = "rptNotaCreditoUbicacion" Or strReporte = "rptNotaCreditoValor" Then
        Me.Caption = "Nota de Crédito - " & strNumero
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        
        Dim numDocModificado As String
        
        strSql = " SELECT COALESCE(doc_ele_claveacceso,'') as doc_ele_claveacceso " & _
                 " FROM doc_electronico " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND doc_ele_coddoc='04' " & _
                 " ANd doc_ele_codigo='" & strNumero & "'"
        
        clsConAUX.Ejecutar strSql
        
        If clsConAUX.adorec_Def.RecordCount > 0 Then
            CDEClaveAcceso = Replace(code128$(clsConAUX.adorec_Def("doc_ele_claveacceso")), "'", "''")
        Else
            CDEClaveAcceso = ""
        End If
        
        
        strSqlAux = "SELECT ing_total,ing_factura,ing_fecha,per_codigo " & _
                   " FROM ingreso " & _
                   " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                   " AND ingreso.ing_codigo='" & strNumero & "' " & _
                   " AND ingreso.tip_ing_codigo='DCL' "
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("ing_total"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("ing_total") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        
        
        If Len(clsConAUX.adorec_Def("ing_factura")) = 11 Then
            
            numDocModificado = Right(Left(clsConAUX.adorec_Def("ing_factura"), Len(clsConAUX.adorec_Def("ing_factura")) - 7), 3) & "-" & _
                               Format(Left(clsConAUX.adorec_Def("ing_factura"), Len(clsConAUX.adorec_Def("ing_factura")) - 10), "000") & "-" & _
                               Format(Right(clsConAUX.adorec_Def("ing_factura"), 7), "000000000")
        Else
            strSql = " SELECT TOP 1 egr_codigo,egr_fecha " & _
                     " FROM egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='FAC' " & _
                     " AND egr_anulado=0 " & _
                     " AND egr_fecha<='" & clsConAUX.adorec_Def("ing_fecha") & "'" & _
                     " AND per_codigo='" & clsConAUX.adorec_Def("per_codigo") & "'" & _
                     " ORDER BY egr_codigo "
            clsConAUX.Ejecutar strSql
            If clsConAUX.adorec_Def.RecordCount > 0 Then
            numDocModificado = Right(Left(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 7), 3) & "-" & _
                               Format(Left(clsConAUX.adorec_Def("egr_codigo"), Len(clsConAUX.adorec_Def("egr_codigo")) - 10), "000") & "-" & _
                               Format(Right(clsConAUX.adorec_Def("egr_codigo"), 7), "000000000")
            End If
        End If
        
        Set tNum2Text = Nothing
        strSqlAux = " SELECT COALESCE(COUNT(*),0) as num " & _
                 " FROM det_ingreso " & _
                 " WHERE det_ingreso.tip_ing_codigo='DCL' " & _
                 " AND det_ingreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_ingreso.ing_codigo='" & strNumero & "'"
        clsConAUX.Ejecutar strSqlAux
        If clsConAUX.adorec_Def("num") > 0 Then
            Dim j As Long
            strSql = " EXEC Sp_Drop_Table_if_Exist 'Temp' "
            clsConAUX.Ejecutar strSql
            
'            strSQL = " CREATE TABLE Temp AS SELECT det_ing_cantidad,det_ingreso.prd_codigo as prd_codi,prd_nombre," & _
'                    " CONCAT(ROUND(det_ing_precio,2),'            ') as det_ing_precio,(det_ing_precio*det_ing_cantidad) as tot " & _
'                    " FROM (det_ingreso INNER JOIN producto ON det_ingreso.emp_codigo=producto.emp_codigo AND det_ingreso.prd_codigo=producto.prd_codigo) " & _
'                    " WHERE det_ingreso.emp_codigo='" & strEmpresa & "' " & _
'                    " AND det_ingreso.ing_codigo='" & strNumero & "' " & _
'                    " AND det_ingreso.tip_ing_codigo='DCL'"
'            clsConAUX.Ejecutar strSQL
'
'            strSQL = " SELECT * FROM Temp "
'            clsConAUX.Ejecutar strSQL
'            j = FormatoD0(clsConAUX.adorec_Def.RecordCount)

                GetSQL = " SELECT CONCAT('Nº de Nota Credito: ',FORMAT(1*Left(ingreso.ing_codigo, LEN(ingreso.ing_codigo) - 10)*1,'000'),'-',FORMAT(1*SUBSTRING(CAST(ingreso.ing_codigo as varchar), LEN(ingreso.ing_codigo) - 9, 3)*1,'000'),'-',FORMAT(1*Right(ingreso.ing_codigo, 7)*1,'000000000'),' - ',FORMAT(current_timestamp,'HH:MM')) as todo,ingreso.ing_codigo,ingreso.ven_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, ingreso.per_codigo,persona.per_ruc,CONCAT(FORMAT(1*IIF(len(ing_factura)>7,LEFT(ing_factura,LEN(ing_factura)-7),0),'000000'),'-', FORMAT(1*RIGHT(ing_factura,7),'000000000')) AS ing_factura,ing_observacion," & _
                        " '" & CDEClaveAcceso & "' AS CDEClaveAcceso,'" & numDocModificado & "' as numDocModificado,emp_nombre,emp_direccion,emp_telf,emp_ruc,doc_ele_claveacceso,doc_ele_autorizacion," & _
                        " persona.per_direccion,persona.per_direccion2,persona.per_telf,ciu_nombre,ing_fecha," & _
                        " det_ingreso.prd_codigo as prd_codigo,prd_nombre as nombre,COALESCE(det_ingreso.prd_ubica_linea,'--') as ubica," & _
                        " det_ingreso.prd_codigo,det_ing_cantidad," & _
                        " det_ing_cantidad as cantidad," & _
                        " uni_nombre,det_ing_precio,det_ing_precio*det_ing_cantidad-(det_ing_dcto/det_ing_cantidad*det_ing_cantidad) as utot," & _
                        " ing_subtotal,ing_dcto,COALESCE(egr_fecha,'') as egr_fecha,det_ing_dcto," & _
                        " ing_subtotal_o , ing_impuesto, ing_total,'" & strValor & "' as valLetra,cod_iva_porcentaje as Piva,ing_usumod,CONCAT(ven_apellido,' ',ven_nombre) as vendedor,persona.per_fax,ing_fechamod,"
                        
                GetSQL = GetSQL & " CONCAT(COALESCE(p1.per_apellido,''),' ',COALESCE(p1.per_nombre,'')) as gerente," & _
                        " CONCAT(COALESCE(p2.per_apellido,''),' ',COALESCE(p2.per_nombre,'')) as director," & _
                        " CONCAT(COALESCE(EMP.per_apellido,''),' ',COALESCE(EMP.per_nombre,'')) as EMPR," & _
                        " CONCAT(COALESCE(EJE.per_apellido,''),' ',COALESCE(EJE.per_nombre,'')) as EJES," & _
                        " CONCAT(COALESCE(N5.per_apellido,''),' ',COALESCE(N5.per_nombre,'')) as NN5," & _
                        " CONCAT(COALESCE(N6.per_apellido,''),' ',COALESCE(N6.per_nombre,'')) as NN6," & _
                        " CONCAT(COALESCE(N7.per_apellido,''),' ',COALESCE(N7.per_nombre,'')) as NN7," & _
                        " CONCAT(COALESCE(N8.per_apellido,''),' ',COALESCE(N8.per_nombre,'')) as NN8," & _
                        " CONCAT(COALESCE(N9.per_apellido,''),' ',COALESCE(N9.per_nombre,'')) as NN9," & _
                        " IIF(p1.per_codigo IS NULL,'','G.Z:') as ger," & _
                        " IIF(p2.per_codigo IS NULL,'','Dir:') as dir," & _
                        " IIF(EMP.per_codigo IS NULL,'','Emprendedor:') as nemp, " & _
                        " IIF(EJE.per_codigo IS NULL,'','Ejecut:') as neje," & _
                        " IIF(N5.per_codigo IS NULL,'','Coord. 5:') as nt5," & _
                        " IIF(N6.per_codigo IS NULL,'','Coord. 4:') as nt6," & _
                        " IIF(N7.per_codigo IS NULL,'','Coord. 3:') as nt7," & _
                        " IIF(N8.per_codigo IS NULL,'','Coord. 2:') as nt8," & _
                        " IIF(N9.per_codigo IS NULL,'','Coord. 1:') as nt9, " & _
                        " CONCAT('Estimados. De acuerdo con la información registrada en nuestro sistema, en tu  mail: - ',persona.per_email,' - recibirás tus documentos electrónicos autorizada por el SRI, según las nueva ley en vigencia. Si no tienes actualizados tus datos comunicate al 1800CATALOGOS para pedir esta actualización') as mensaje, " & _
                        " IIF(doc_ele_autorizacion='' OR doc_ele_autorizacion is null,'DOCUMENTO SIN VALIDEZ TRIBUTARIA','') as mensaje2 "
                GetSQL = GetSQL & " FROM empresa inner join ingreso ON empresa.emp_codigo=ingreso.emp_codigo " & _
                        " INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo " & _
                        " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                        " INNER JOIN codigo_iva ON ingreso.cod_iva_codigo=codigo_iva.cod_iva_codigo" & _
                        " INNER JOIN det_ingreso_ubicacion det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo AND ingreso.ing_codigo=det_ingreso.ing_codigo " & _
                        " INNER JOIN producto ON det_ingreso.emp_codigo=producto.emp_codigo AND det_ingreso.prd_codigo=producto.prd_codigo " & _
                        " INNER JOIN unidad ON producto.emp_codigo=unidad.emp_codigo AND producto.uni_codigo=unidad.uni_codigo "
'                GetSQL = GetSQL & " LEFT JOIN (SELECT det_contenedor_mercaderia.emp_codigo,tip_mov_codigo,mov_codigo,dep_codigo,det_contenedor_mercaderia.prd_codigo,CONCAT(contenedor_mercaderia.ubi_bod_codigo,'-',contenedor_mercaderia.con_mer_codigo) as ubica,det_contenedor_mercaderia.det_con_mer_cantidad" & _
'                        " FROM det_contenedor_mercaderia INNER JOIN contenedor_mercaderia ON det_contenedor_mercaderia.emp_codigo=contenedor_mercaderia.emp_codigo " & _
'                        " AND det_contenedor_mercaderia.con_mer_codigo=contenedor_mercaderia.con_mer_codigo " & _
'                        " WHERE det_contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
'                        " AND tip_mov_codigo='DCL' " & _
'                        " AND det_con_mer_cantidad!=0 " & _
'                        " AND mov_codigo in (" & strNumero & ")) ubi ON det_ingreso.emp_codigo=ubi.emp_codigo " & _
'                        " AND det_ingreso.tip_ing_codigo=ubi.tip_mov_codigo " & _
'                        " AND det_ingreso.ing_codigo=ubi.mov_codigo " & _
'                        " AND det_ingreso.dep_codigo=ubi.dep_codigo " & _
'                        " AND det_ingreso.prd_codigo=ubi.prd_codigo "
                GetSQL = GetSQL & " LEFT JOIN egreso ON ingreso.emp_codigo=egreso.emp_codigo AND ingreso.ing_factura=CAST(egreso.egr_codigo AS varchar) AND egreso.tip_egr_codigo='FAC' LEFT JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
                        " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref AND p1.per_es_gz=1 " & _
                        " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 AND p1.per_es_di=1 " & _
                        " LEFT JOIN persona as EMP ON persona.emp_codigo = EMP.emp_codigo " & _
                        " AND persona.per_codigo_ref3 = EMP.per_codigo AND EMP.per_es_em=1 " & _
                        " LEFT JOIN persona as EJE ON persona.emp_codigo = EJE.emp_codigo " & _
                        " AND persona.per_codigo_ref4 = EJE.per_codigo AND EJE.per_es_ee=1 " & _
                        " LEFT JOIN persona as N5 ON persona.emp_codigo = N5.emp_codigo " & _
                        " AND persona.per_codigo_ref5 = N5.per_codigo AND N5.per_es_n5=1 " & _
                        " LEFT JOIN persona as N6 ON persona.emp_codigo = N6.emp_codigo " & _
                        " AND persona.per_codigo_ref6 = N6.per_codigo AND N6.per_es_n6=1 " & _
                        " LEFT JOIN persona as N7 ON persona.emp_codigo = N7.emp_codigo " & _
                        " AND persona.per_codigo_ref7 = N7.per_codigo AND N7.per_es_n7=1 " & _
                        " LEFT JOIN persona as N8 ON persona.emp_codigo = N8.emp_codigo " & _
                        " AND persona.per_codigo_ref8 = N8.per_codigo AND N8.per_es_n8=1 " & _
                        " LEFT JOIN persona as N9 ON persona.emp_codigo = N9.emp_codigo " & _
                        " AND persona.per_codigo_ref9 = N9.per_codigo AND N9.per_es_n9=1 " & _
                        " LEFT JOIN doc_electronico ON ingreso.emp_codigo=doc_electronico.emp_codigo AND ingreso.ing_codigo=doc_electronico.doc_ele_codigo " & _
                        " AND doc_electronico.doc_ele_coddoc='04' "
                GetSQL = GetSQL & " WHERE ingreso.tip_ing_codigo='DCL' " & _
                        " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                        " AND ingreso.ing_codigo='" & strNumero & "'" & _
                        " ORDER BY ubica,prd_nombre"
        Else
            strReporte = "rptNotaCreditoValor"
            GetSQL = " SELECT CONCAT('Nº de Nota Credito: ',FORMAT(1*Left(ingreso.ing_codigo, LEN(ingreso.ing_codigo) - 10),'000'),'-',FORMAT(1*SUBSTRING(CAST(ingreso.ing_codigo as varchar), LEN(ingreso.ing_codigo) - 9, 3),'000'),'-',FORMAT(1*Right(ingreso.ing_codigo, 7),'000000000'),' - ',format(current_timestamp,'HH:MM')) as todo,ingreso.ing_codigo,ingreso.ven_codigo,CONCAT(persona.per_apellido,' ',persona.per_nombre) as per, ingreso.per_codigo,persona.per_ruc,CONCAT(FORMAT(1*IIF(len(ing_factura)>7,LEFT(ing_factura,LEN(ing_factura)-7),0),'000000'),'-', FORMAT(1*RIGHT(ing_factura,7),'000000000')) AS ing_factura,ing_observacion," & _
                     " '" & CDEClaveAcceso & "' AS CDEClaveAcceso,'" & numDocModificado & "' as numDocModificado,emp_nombre,emp_direccion,emp_telf,emp_ruc,doc_ele_claveacceso,doc_ele_autorizacion," & _
                     " persona.per_direccion,persona.per_direccion2,persona.per_telf,ciu_nombre,ing_fecha,ing_observacion,ing_dcto*(-1) as utot, ing_subtotal, ing_dcto,'1' as cantidad,COALESCE(egr_fecha,'') as egr_fecha," & _
                     " ing_subtotal_o , ing_impuesto, ing_total,'" & strValor & "' as valLetra,cod_iva_porcentaje as Piva,ing_usumod,CONCAT(ven_apellido,' ',ven_nombre) as vendedor,persona.per_fax,ing_fechamod,CONCAT(p1.per_apellido,' ',p1.per_nombre) as gerente,CONCAT(p2.per_apellido,' ',p2.per_nombre) as director,IIF(p1.per_codigo IS NULL,'','G.Z:') as ger,IIF(p2.per_codigo IS NULL,'','Dir:') as dir, " & _
                     " CONCAT('Estimados. De acuerdo con la información registrada en nuestro sistema, en tu  mail: - ',persona.per_email,' - recibirás tus documentos electrónicos autorizada por el SRI, según las nueva ley en vigencia. Si no tienes actualizados tus datos comunicate al 1800CATALOGOS para pedir esta actualización') as mensaje, " & _
                     " IIF(doc_ele_autorizacion='' OR doc_ele_autorizacion is null,'DOCUMENTO SIN VALIDEZ TRIBUTARIA','') as mensaje2 " & _
                     " FROM empresa inner join ingreso ON empresa.emp_codigo=ingreso.emp_codigo " & _
                     " INNER JOIN persona ON ingreso.emp_codigo=persona.emp_codigo AND ingreso.per_codigo=persona.per_codigo " & _
                     " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
                     " INNER JOIN codigo_iva ON ingreso.cod_iva_codigo=codigo_iva.cod_iva_codigo" & _
                     " LEFT JOIN egreso ON ingreso.emp_codigo=egreso.emp_codigo AND ingreso.ing_factura=CAST(egreso.egr_codigo as varchar) AND egreso.tip_egr_codigo='FAC' LEFT JOIN vendedor ON persona.emp_codigo=vendedor.emp_codigo AND persona.ven_codigo=vendedor.ven_codigo " & _
                     " LEFT JOIN persona p1 ON p1.emp_codigo=persona.emp_codigo AND p1.per_codigo=persona.per_codigo_ref " & _
                     " LEFT JOIN persona p2 ON p2.emp_codigo=persona.emp_codigo AND p2.per_codigo=persona.per_codigo_ref2 " & _
                     " LEFT JOIN doc_electronico ON ingreso.emp_codigo=doc_electronico.emp_codigo AND ingreso.ing_codigo=doc_electronico.doc_ele_codigo " & _
                     " AND doc_electronico.doc_ele_coddoc='04' " & _
                     " WHERE ingreso.tip_ing_codigo='DCL' " & _
                     " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                     " AND ingreso.ing_codigo='" & strNumero & "'"
        End If
    ElseIf strReporte = "rptCotizacion" Then
        Me.Caption = "Cotización - " & strNumero
        
        GetSQL = " SELECT TRIM(CONCAT(per_nombre,' ',per_apellido)) as nombC, TRIM(CONCAT(IIF(ven_signatura<>'_',ven_signatura,''),' ',ven_nombre,' ',ven_apellido)) as nombV,COALESCE('" & Atencion & "','') as atencion,cotizacion.cot_codigo as codigo, " & _
             " cot_fecha, cot_observacion, det_cotizacion.prd_codigo, CONCAT(prd_nombre,' (',mar_nombre,')') as prd_nombre, det_cot_cantidad, round(det_cot_precio,2) as det_cot_precio, round(det_cot_cantidad * det_cot_precio,2) as totPrd, ven_cargo,CONCAT(per_telf,'/',per_fax) as per_TF,det_cot_item,uni_codigo,ven_telf2,ven_email " & _
             " FROM ((((persona INNER JOIN proyecto_venta ON (persona.per_codigo = proyecto_venta.per_codigo) AND (persona.emp_codigo = proyecto_venta.emp_codigo)) " & _
             " INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo)) " & _
             " INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
             " INNER JOIN producto ON (det_cotizacion.emp_codigo = producto.emp_codigo) AND (det_cotizacion.prd_codigo = producto.prd_codigo)) " & _
             " INNER JOIN vendedor ON (proyecto_venta.ven_codigo = vendedor.ven_codigo) AND (proyecto_venta.emp_codigo = vendedor.emp_codigo) " & _
             " INNER JOIN marca ON (marca.mar_codigo = producto.mar_codigo) AND (producto.emp_codigo = marca.emp_codigo) " & _
             " WHERE vendedor.emp_codigo='" & strEmpresa & "' AND cotizacion.cot_codigo='" & strNumero & "' " & _
             " ORDER BY det_cot_item "
    ElseIf strReporte = "rptCotizacionF" Then
        Me.Caption = "Cotización - " & strNumero
        GetSQL = " SELECT TRIM(CONCAT(per_nombre,' ',per_apellido)) as nombC, TRIM(CONCAT(IIF(ven_signatura<>'_',ven_signatura,''),' ',ven_nombre,' ',ven_apellido)) as nombV,COALESCE('" & Atencion & "','') as atencion,cotizacion.cot_codigo as codigo, " & _
             " cot_fecha, cot_observacion, det_cotizacion.prd_codigo, CONCAT(prd_nombre,' (',mar_nombre,')') as prd_nombre, det_cot_cantidad, round(det_cot_precio,2) as det_cot_precio, round(det_cot_cantidad * det_cot_precio,2) as totPrd, ven_cargo,CONCAT(per_telf,'/',per_fax) as per_TF,det_cot_item,uni_codigo,ven_telf2,ven_email " & _
             " FROM ((((persona INNER JOIN proyecto_venta ON (persona.per_codigo = proyecto_venta.per_codigo) AND (persona.emp_codigo = proyecto_venta.emp_codigo)) " & _
             " INNER JOIN cotizacion ON (proyecto_venta.emp_codigo = cotizacion.emp_codigo) AND (proyecto_venta.pro_ven_codigo = cotizacion.pro_ven_codigo)) " & _
             " INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
             " INNER JOIN producto ON (det_cotizacion.emp_codigo = producto.emp_codigo) AND (det_cotizacion.prd_codigo = producto.prd_codigo)) " & _
             " INNER JOIN vendedor ON (proyecto_venta.ven_codigo = vendedor.ven_codigo) AND (proyecto_venta.emp_codigo = vendedor.emp_codigo) " & _
             " INNER JOIN marca ON (marca.mar_codigo = producto.mar_codigo) AND (producto.emp_codigo = marca.emp_codigo) " & _
             " WHERE vendedor.emp_codigo='" & strEmpresa & "' AND cotizacion.cot_codigo='" & strNumero & "' " & _
             " ORDER BY det_cot_item "
    ElseIf strReporte = "rptComprobanteEgreso" Then
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = " SELECT com_egr_ch_valor,com_egr_ch_fecha " & _
                 " FROM comp_egreso " & _
                 " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND comp_egreso.com_egr_codigo='" & strNumero & "'"
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("com_egr_ch_valor"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("com_egr_ch_valor") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        j = Len(strValor)
        For i = j To 10000
            strValor = strValor & " --"
        Next i
        strFecha = Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "dd") & " de " & Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "MMMM") & " del " & Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "yyyy")
        Set clsConAUX = Nothing
        Set tNum2Text = Nothing
        Me.Caption = "Comprobante de Egreso - " & strNumero
        GetSQL = " SELECT COALESCE(CONCAT(per_apellido,' ',per_nombre),com_egr_nombre2) as per_apenom,com_egr_ch_valor,'" & strValor & "' as com_egr_ch_valor_letras," & _
                 " '" & strFecha & "' AS com_egr_ch_fecha, com_egr_ch_num, cta_ban_numero, ban_nombre,asiento.asi_numasiento,asi_fecha,asi_descripcion," & _
                 " asi_totaldebe,asi_totalhaber,asi_usumod,det_asiento.cta_codigo,cta_nombre,det_asi_debe,det_asi_haber, COALESCE(cen_cos_nombre,'') as cen_cos_nombre,per_telf,emp_nombre " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo=persona.per_codigo AND comp_egreso.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN banco ON comp_egreso.ban_codigo=banco.ban_codigo " & _
                 " INNER JOIN asiento ON comp_egreso.emp_codigo=asiento.emp_codigo AND comp_egreso.asi_numasiento=asiento.asi_numasiento " & _
                 " INNER JOIN empresa ON asiento.emp_codigo=empresa.emp_codigo " & _
                 " INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
                 " INNER JOIN ctaconta ON det_asiento.cta_codigo=ctaconta.cta_codigo AND det_asiento.emp_codigo=ctaconta.emp_codigo " & _
                 " LEFT JOIN centro_costo ON det_asiento.cen_cos_codigo=centro_costo.cen_cos_codigo AND det_asiento.emp_codigo=centro_costo.emp_codigo " & _
                 " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND comp_egreso.com_egr_codigo='" & strNumero & "'"
    ElseIf strReporte = "rptComprobanteIngreso" Then
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = " SELECT not_d_c_monto ,FORMAT(not_d_c_fecha,'yyyy-mm-dd') as not_d_c_fecha " & _
                 " FROM nota_d_c " & _
                 " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND nota_d_c.asi_numasiento='" & strAsiento & "'"
        clsConAUX.Ejecutar strSqlAux
        lngValor = Int(clsConAUX.adorec_Def("not_d_c_monto"))
        intValor = Right(Str(Int(clsConAUX.adorec_Def("not_d_c_monto") * 100)), 2)
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        j = Len(strValor)
        For i = j To 10000
            strValor = strValor & " --"
        Next i
        strFecha = Format(clsConAUX.adorec_Def("not_d_c_fecha"), "dd") & " de " & Format(clsConAUX.adorec_Def("not_d_c_fecha"), "MMMM") & " del " & Format(clsConAUX.adorec_Def("not_d_c_fecha"), "yyyy")
        Set clsConAUX = Nothing
        Set tNum2Text = Nothing
        Me.Caption = "Comprobante de Ingreso - " & strNumero
        GetSQL = " SELECT '' as per_apenom,not_d_c_monto,'" & strValor & "' as com_egr_ch_valor_letras," & _
                 " '" & strFecha & "' AS not_d_c_fecha, '' as chnum, cta_ban_numero, ban_nombre,asiento.asi_numasiento,asi_fecha,asi_descripcion," & _
                 " asi_totaldebe,asi_totalhaber,asi_usumod,det_asiento.cta_codigo,cta_nombre,det_asi_debe,det_asi_haber, COALESCE(cen_cos_nombre,'') as cen_cos_nombre " & _
                 " FROM nota_d_c " & _
                 " INNER JOIN banco ON nota_d_c.ban_codigo=banco.ban_codigo " & _
                 " INNER JOIN asiento ON nota_d_c.emp_codigo=asiento.emp_codigo AND nota_d_c.asi_numasiento=asiento.asi_numasiento " & _
                 " INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
                 " INNER JOIN ctaconta ON det_asiento.cta_codigo=ctaconta.cta_codigo AND det_asiento.emp_codigo=ctaconta.emp_codigo " & _
                 " LEFT JOIN centro_costo ON det_asiento.cen_cos_codigo=centro_costo.cen_cos_codigo AND det_asiento.emp_codigo=centro_costo.emp_codigo " & _
                 " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND nota_d_c.asi_numasiento='" & strAsiento & "'"
    ElseIf strReporte = "rptCheque" Then
        clsConAUX.Inicializar AdoConn, AdoConnMaster
        strSqlAux = " SELECT CONCAT(per_apellido,' ',per_nombre) as per_apenom,com_egr_ch_valor,'VALOR EN LETRAS' as com_egr_ch_valor_letras,com_egr_ch_fecha,com_egr_ch_num,cta_ban_numero, ban_nombre,asiento.asi_numasiento,asi_fecha,asi_descripcion,asi_totaldebe,asi_totalhaber,asi_usumod,det_asiento.cta_codigo,cta_nombre,det_asi_debe,det_asi_haber  " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo=persona.per_codigo AND comp_egreso.emp_codigo=persona.emp_codigo " & _
                 " INNER JOIN banco ON comp_egreso.ban_codigo=banco.ban_codigo " & _
                 " INNER JOIN asiento ON comp_egreso.emp_codigo=asiento.emp_codigo AND comp_egreso.asi_numasiento=asiento.asi_numasiento " & _
                 " INNER JOIN det_asiento ON asiento.emp_codigo=det_asiento.emp_codigo AND asiento.asi_numasiento=det_asiento.asi_numasiento " & _
                 " INNER JOIN ctaconta ON det_asiento.cta_codigo=ctaconta.cta_codigo AND det_asiento.emp_codigo=ctaconta.emp_codigo " & _
                 " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND comp_egreso.com_egr_codigo='" & strNumero & "'"
        clsConAUX.Ejecutar strSqlAux
        If clsConAUX.adorec_Def.RecordCount > 0 Then
            lngValor = Int(clsConAUX.adorec_Def("com_egr_ch_valor"))
            intValor = Right(Str(Int(clsConAUX.adorec_Def("com_egr_ch_valor") * 100)), 2)
        End If
        strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
        
        j = Len(strValor)
        For i = j To 10000
            strValor = strValor & " --"
        Next i
        If clsConAUX.adorec_Def.RecordCount > 0 Then
            strFecha = Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "yyyy") & " / " & Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "mm") & " / " & Format(clsConAUX.adorec_Def("com_egr_ch_fecha"), "dd")
        End If
        Set clsConAUX = Nothing
        Set tNum2Text = Nothing
        Me.Caption = "Comprobante de Egreso - " & strNumero
        GetSQL = " SELECT COALESCE(com_egr_nombre2,CONCAT(per_apellido,' ',per_nombre)) as per_apenom,com_egr_ch_valor,'" & strValor & "' as com_egr_ch_valor_letras," & _
                 " '" & strFecha & "' AS com_egr_ch_fecha " & _
                 " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo=persona.per_codigo AND comp_egreso.emp_codigo=persona.emp_codigo " & _
                 " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND comp_egreso.com_egr_codigo='" & strNumero & "'"
    ElseIf strReporte = "rptConteo" Then
        GetSQL = " SELECT inventario.inv_codigo,inventario.inv_fecha,inventario.inv_observacion," & _
                " dep_codigo,det_inv_cantidad,det_inventario.prd_codigo,prd_nombre,emp_nombre " & _
                " FROM inventario INNER JOIN det_inventario ON inventario.emp_codigo=det_inventario.emp_codigo AND inventario.inv_codigo=det_inventario.inv_codigo " & _
                " INNER JOIN producto ON det_inventario.emp_codigo=producto.emp_codigo AND det_inventario.prd_codigo=producto.prd_codigo " & _
                " INNER JOIN empresa ON inventario.emp_codigo=empresa.emp_codigo " & _
                " WHERE inventario.emp_codigo='" & strEmpresa & "' " & _
                " AND inventario.inv_codigo='" & strNumero & "'"
    End If
End Function


Private Sub BuscarPorc(Valor As Long, item As Integer)
  clsSql1.adorec_Def.MoveFirst
  Dim a As Boolean
  a = True
  Do While clsSql1.adorec_Def.EOF = False
    If clsSql1.adorec_Def("det_com_ret_porcentaje") = Valor Then
      strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,total_f,ret_codigo,det_com_ret_porcentaje_f) " & _
               " VALUES('" & item & "','" & strNumero & "','" & strTipo & "','" & clsSql1.adorec_Def("ret_codigo") & "','" & clsSql1.adorec_Def("det_com_ret_valor") & "','" & Format((clsSql1.adorec_Def("det_com_ret_valor") * clsSql1.adorec_Def("det_com_ret_porcentaje")) / 100, "###0.00") & "','" & Valor & "','" & Valor & "')"
      clsSQL2.Ejecutar strSql
      strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
               " VALUES('" & item + 1 & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
      clsSQL2.Ejecutar strSql
      a = False
      Exit Sub
    End If
    clsSql1.adorec_Def.MoveNext
  Loop
'  strSql = " INSERT INTO Detalle(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo)" & _
'           " VALUES('" & item & "','" & strNumero & "','" & strTipo & "','" & Valor & "')"
    strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
             " VALUES('" & item & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
    'clsSQL2.Ejecutar strSql
    strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
             " VALUES('" & item + 1 & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
    clsSQL2.Ejecutar strSql
    If a = True Then
        strSql = " INSERT INTO Detalle" & strUsuario & "(item,cue_p_c_codigo,cue_p_c_tipo,ret_codigo_f,det_com_ret_valor_f,det_com_ret_porcentaje_f,total_f,ret_codigo)" & _
                 " VALUES('" & item + 1 & "','" & strNumero & "','" & strTipo & "','-1','','','','')"
        clsSQL2.Ejecutar strSql
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim clsConAux1 As New clsConsulta
    clsConAux1.Inicializar AdoConn, AdoConnMaster
    If strReporte = "rptRolPagos" Then
        strSql = " EXEC Sp_Drop_Table_if_Exist 'EstadoCuentaVB'"
        clsConAux1.Ejecutar strSql
    ElseIf strReporte = "rptPreFactura" Then
        strSql = " EXEC Sp_Drop_Table_if_Exist 'recs" & strNumero & "'"
        clsConAux1.Ejecutar strSql
    End If
    strSql = " EXEC Sp_Drop_Table_if_Exist 'recsDetalle" & strUsuario & "'"
    clsConAux1.Ejecutar strSql
End Sub
