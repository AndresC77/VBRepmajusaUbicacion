Attribute VB_Name = "SubMain"
Sub Main()
    Dim Argumentos() As String
    Argumentos = Split(CStr(Command()), " ")
    NombreComercial = "RYB Importadores"
    CorreoServicioAlCliente = "servicioalcliente@rbimportadores.com"
    CorreoCartera = "auditoriainterna@rbimportadores.com"
    CorreoCompras = "gerenciadeoperaciones@rbimportadores.com"
    CorreoAsistenteCos = "asistentecos@rbimportadores.com"
    CorreoSupervisorDeTransportes = "supervisordetransportes@rbimportadores.com"
    CorreoNoticias = "noticias@rbimportadores.com"
    CorreoAsistenteCartera = "asistentecartera3@rbimportadores.com"
    
'    NombreComercial = "StudioModa"
'    CorreoServicioAlCliente = "atencioncliente@puntomoda.com.ec"
'    CorreoAsistenteCos = "digitador@puntomoda.com.ec"
'    CorreoSupervisorDeTransportes = "bodega@puntomoda.com.ec"
'    CorreoNoticias = "atencioncliente@puntomoda.com.ec"
    
    frmConexion.Show
    If UBound(Argumentos) > 0 Then
        frmConexion.txtUsuario.Text = Argumentos(0)
        frmConexion.txtClave.Text = Argumentos(1)
        frmConexion.cmdAceptar_Click
    End If
    
End Sub
