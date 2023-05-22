Dim strComputer 
Dim strCompName 
Dim strUserName 
Dim objShell

strComputer = "."
Set objShell=CreateObject("WScript.Shell")

'Variáveis que armazenam variável do sistema cmd 
'Como nome do computador e do usuario
strCompName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")

'Realiza um certo acordo de segurança com o computador  
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Pesquisa por placa de rede ativa
Set IPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
 
'Esse for é para encontrar o endereço de ip da placa de rede
'LBOUND to UBOUND - significa que a cada passo irá ser feita uma busca do 1 ao 'numero 254 para ir preenchendo o ip = x.x.x.x que no fim das contas vai 'resultar no seguinte ip: 192.168.3.5  

For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
        For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
            WScript.Echo StruserName & " - " & Strcompname & " - " & IPConfig.IPAddress(0)
        Next
    End If
Next
