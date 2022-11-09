VERSION 5.00
Begin VB.Form FrmVentasMail 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "FrmVentasMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strCadena = "SELECT mail,nombre_completo FROM persona WHERE dni='" & Trim(FrmVentas.TxtCodCliente.Text) & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
     Call enviar(Trim(rstT("mail")), rstT("nombre_completo"))
    
Else

   MsgBox "Usuario no tiene mail para el envio", vbInformation, KEY_EMPRESA
   Exit Sub
End If
End Sub

Private Sub enviar(ByVal correo As String, ByVal Nombre As String)
        Dim introduccion As String
        Dim reporte As String
        Dim FechaEnvio As String
        
    
'-------------------------------------
    'cmdSend.Enabled = False
    'lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With sendmail1

        'Valida (opcional)
        .SMTPHostValidacion = VALIDATE_HOST_NONE
        'Valida la sintaxis de l cuenta (opcional)
        .ValidarEmail = VALIDATE_SYNTAX
        'Opcional
        .Delimitador = ";"
        'Texto  para visualizar en el campo De (opcional)
        .FromDisplayName = "WWW.VITEKEY.COM"
        'Requerido (Nombre del servidor SMTP)
        .SMTPHost = "smtp.live.com"
        'Requerido
        .Remitente = "percy19_is@hotmail.com"
        'Requerido
        .Destinatario = Trim(correo)
        'Asunto del mensaje
        .Asunto = "PAGO ONLINE:" + Space(2) + "-" + Trim(KEY_FECHA) + Space(1) + str(Time)
        'Cuerpodel mensaje
         pasword = Trim(str(CLng((100 - 999) * Rnd + 999))) & Trim(str(CLng((100 - 999) * Rnd + 999)))
        .Mensaje = "Hola:" + Space(1) + Trim(Nombre) + Chr(13) + "Hay un pago:" + Chr(13) + "A:" + KEY_EMPRESA + Chr(13) + "Monto:" + moneda(FrmVentas.DtcMoneda.BoundText) + Space(1) + Format(Val(Format(FrmVentas.lblTotal.Caption, "###0.00")) - Val(Format(FrmVentas.lblPago.Caption, "###0.00")), "###0.00") + Chr(13) + "*****************************" + Chr(13) + "PASSWORD:" + pasword + Chr(13) + "*****************************"
               
        'Adjunto (opcional)
        .Adjunto = "" 'Trim(txtAttach.Text)
        
        'Opcional (Prioridad del mensaje)
        .Prioridad = Alta
        'Opcional (si requiere autentificación)
        .UsarLoginSMTP = True
        'Requerido si requiere autentificación
        .usuario = "percy19_is@hotmail.com" 'txtUserName
        .password = "200119828372000" 'txtPassword
        
        'txtServer.Text = .SMTPHost
       'Opcional (por defectoutiliza el Tipo MIME)
       .Codificacion = MIME_ENCODE
       
       'Envia el Mail
       .EnviarEmail
    
    End With
    Screen.MousePointer = vbDefault
    
  strCadena = "INSERT INTO empresa_random(dni,id_suc,id_empresa,fecha,codigo) VALUES ('" & Trim(FrmVentas.TxtCodCliente.Text) & "','1','" & KEY_RUC & "','" & Format(KEY_FECHA, "yyyy-mm-dd") & "','" & Trim(pasword) & "')"
  CnBd.Execute (strCadena)
   



End Sub

