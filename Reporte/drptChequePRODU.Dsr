VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} drptChequePRODU 
   Caption         =   "Cheque"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   MDIChild        =   -1  'True
   _ExtentX        =   15849
   _ExtentY        =   4048
   _Version        =   393216
   _DesignerVersion=   100688210
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   1
   GridY           =   1
   LeftMargin      =   490
   RightMargin     =   1440
   TopMargin       =   640
   BottomMargin    =   1440
   _Settings       =   7
   NumSections     =   1
   SectionCode0    =   4
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "secDetalle"
      Object.Height          =   1590
      NumControls     =   5
      ItemType0       =   4
      BeginProperty Item0 {1C13A8E2-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "txtBeneficiario"
         Object.Left            =   315
         Object.Top             =   135
         Object.Width           =   4530
         Object.Height          =   285
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataField       =   "per_apenom"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      ItemType1       =   4
      BeginProperty Item1 {1C13A8E2-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "txtValorN"
         Object.Left            =   6060
         Object.Top             =   135
         Object.Width           =   1185
         Object.Height          =   330
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataField       =   "com_egr_ch_valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      ItemType2       =   3
      BeginProperty Item2 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "txtValorL"
         Object.Left            =   315
         Object.Top             =   465
         Object.Width           =   6945
         Object.Height          =   390
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Valor En Letras"
      EndProperty
      ItemType3       =   3
      BeginProperty Item3 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "txtFecha"
         Object.Left            =   480
         Object.Top             =   1185
         Object.Width           =   2085
         Object.Height          =   300
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Dia de Mes del Año"
      EndProperty
      ItemType4       =   3
      BeginProperty Item4 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "txtCiudad"
         Object.Top             =   1200
         Object.Width           =   495
         Object.Height          =   300
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Quito,"
      EndProperty
   EndProperty
End
Attribute VB_Name = "drptChequePRODU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clsCon_Def As clsConsulta
Private tNum2Text As cNum2Text

Private Sub DataReport_Terminate()
    Dim i As Long
    Set Me.DataSource = Nothing
    Set clsCon_Def = Nothing
    Set tNum2Text = Nothing

End Sub
Private Sub DataReport_Activate()
    Dim strSql As String
    Dim strValor As String
    Dim lngValor As Long
    Dim intValor As Integer
    Dim i As Long
    Dim j As Long
    
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar AdoConn
    
    strSql = " SELECT CONCAT(per_apellido,' ',per_nombre) as per_apenom,com_egr_ch_valor,com_egr_ch_fecha " & _
             " FROM comp_egreso INNER JOIN persona ON comp_egreso.per_codigo=persona.per_codigo AND comp_egreso.emp_codigo=persona.emp_codigo " & _
             " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND comp_egreso.com_egr_codigo='" & Me.Tag & "' "
    clsCon_Def.Ejecutar strSql
    Set Me.DataSource = clsCon_Def.adorec_Def
    Set tNum2Text = New cNum2Text
    lngValor = Int(clsCon_Def.adorec_Def("com_egr_ch_valor"))
    intValor = Right(Str(Int(clsCon_Def.adorec_Def("com_egr_ch_valor") * 100)), 2)
    strValor = UCase(tNum2Text.Numero2Letra(lngValor, , 0, "", "centavo", 1, 1)) & " " & Format(intValor, "00") & "/100 "
    j = Len(strValor)
    For i = j To 10000
        strValor = strValor & " --"
    Next i
    Me.Sections("secDetalle").Controls("txtValorL").Caption = strValor
    Me.Sections("secDetalle").Controls("txtFecha").Caption = Format(clsCon_Def.adorec_Def("com_egr_ch_fecha"), "dd") & " de " & Format(clsCon_Def.adorec_Def("com_egr_ch_fecha"), "MMMM") & " del " & Format(clsCon_Def.adorec_Def("com_egr_ch_fecha"), "yyyy")
End Sub

Private Sub DataReport_Initialize()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = ((mdiPrincipal.Height - Me.Height) / 2) - (mdiPrincipal.Height / 40)
End Sub

