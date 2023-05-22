Const cdoAnonymous = 0 'Sem autenticação
Const cdoBasic = 1 'Autenticação Básica (Base64)
Const cdoNTLM = 2 'NTLM
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Teste de E-mail" 
objMessage.From = """Me"" <seu_email@seudominio.com>" 
objMessage.To = "destinatario@email.com" 
objMessage.TextBody = "Teste de Mensagem.." & vbCRLF & "Foi enviada utilizando autenticação Base64 e SSL."
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
' Endereço do servidor SMTP
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.seudominio.com"
' Tipo de autenticação
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
' Usuário de autenticação
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "usuario@seudominio.com"
' Senha de autenticação
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "sua_senha"
' Porta do SMTP
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 
' Utiliza SSL ?
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
' Timeout
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
If err.number = 0 then Msgbox "Email enviado com sucesso"