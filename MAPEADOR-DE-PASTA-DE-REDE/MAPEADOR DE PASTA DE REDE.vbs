' Nome do arquivo: Logon.vbs
' Este arquivo deve ficar na pasta C:\Windows\SYSVOL\domain\scripts
'======================================================

'Mapeando as unidades
Set objNetwork = CreateObject("WScript.Network") 
objNetwork.MapNetworkDrive "Z:", "\\nomeOuIpDoServidor\nomeDaPastaCompartilhada" 
objNetwork.MapNetworkDrive "Y:", "\\nomeOuIpDoServidor\nomeDaPastaCompartilhada" 
objNetwork.MapNetworkDrive "X:", "\\nomeOuIpDoServidor\nomeDaPastaCompartilhada" 

'======================================================

'Mapeando Impressora

Set WshNetwork = Wscript.CreateObject("Wscript.Network")
WshNetwork.AddWindowsPrinterConnection "\\nomeOuIp\nomeDaImpressora", "NomeImpressora"
WshNetwork.SetDefaultPrinter "\\nomeOuIp\nomeDaImpressora", "NomeImpressora"

'======================================================

'Mensagem de boas-vindas

If Time <= "12:00:00" Then
MsgBox ("Bom dia, você ingressou na rede.")
ElseIf Time >= "12:00:01" And Time <= "18:00:00" Then
MsgBox ("Boa tarde, você ingressou na rede.")
Else
MsgBox ("Boa noite, você ingressou na rede.")
End If

WScript.Quit