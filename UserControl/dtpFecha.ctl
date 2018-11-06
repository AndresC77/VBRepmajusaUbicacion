VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl dtpFecha 
   BackStyle       =   0  'Transparent
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   330
   ScaleWidth      =   1440
   Begin MSComCtl2.DTPicker dtpFecha 
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   16711683
      CurrentDate     =   37463
   End
End
Attribute VB_Name = "dtpFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private clsSql As New clsConsulta
Private strSql As String
Public pedirClave As Boolean
'Private strClaveMAESTRA As String
Private valorAntes As Date

Public Event Change()

Private Sub dtpFecha_Change()
       RaiseEvent Change
End Sub

Private Sub dtpFecha_GotFocus()
    valorAntes = dtpFecha.Value
End Sub

Private Sub dtpFecha_LostFocus()
    strSql = " SELECT COALESCE(COUNT(*),0) " & _
             " FROM cierre_mes " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cie_mes_ano=" & Year(dtpFecha.Value) & " AND cie_mes_mes=" & Month(dtpFecha.Value)
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        If FormatoD0(clsSql.adorec_Def(0)) > 0 Then
            MsgBox "El mes está cerrado por contabilidad", vbInformation, "Fecha"
            dtpFecha.Value = HoyDia
            Exit Sub
        End If
    End If
    
    If pedirClave = True Then
        strSql = " SELECT par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND par_codigo = 'CMA' "
        clsSql.Ejecutar strSql
        strClaveMAESTRA = clsSql.adorec_Def("par_texto")
        If dtpFecha.Value <> HoyDia Then
            'frmClave.strClaveMAESTRA = strClaveMAESTRA
            frmClave.dblPrecio = "Fecha"
            frmClave.Show vbModal
            If frmClave.Ret = False Then
                dtpFecha.Value = HoyDia
            End If
        Else
            dtpFecha.Value = HoyDia
        End If
    End If
    
End Sub

Public Property Let Value(Valor As Date)
    dtpFecha.Value = Valor
End Property
Public Property Get Value() As Date
    Value = dtpFecha.Value
End Property

Public Property Let Enabled(Valor As Boolean)
    dtpFecha.Enabled = Valor
End Property
Public Property Get Enabled() As Boolean
    Enabled = dtpFecha.Enabled
End Property

Public Property Let Format(Valor As String)
    dtpFecha.CustomFormat = Valor
End Property
Public Property Get Format() As String
    Format = dtpFecha.CustomFormat
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    dtpFecha.Value = PropBag.ReadProperty("Value", Now)
    dtpFecha.Enabled = PropBag.ReadProperty("Enabled", True)
    dtpFecha.CustomFormat = PropBag.ReadProperty("CustomFormat", "yyyy-MM-dd")

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", dtpFecha.Value, Now)
    Call PropBag.WriteProperty("Enabled", dtpFecha.Enabled, True)
    Call PropBag.WriteProperty("Format", dtpFecha.CustomFormat, "yyyy-MM-dd")

End Sub


Private Sub UserControl_Initialize()
    clsSql.Inicializar AdoConn, AdoConnMaster
End Sub

Private Sub UserControl_Resize()
    dtpFecha.Width = Width
    dtpFecha.Height = Height
    dtpFecha.Enabled = Enabled
End Sub
