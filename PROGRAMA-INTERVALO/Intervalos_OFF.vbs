Dim WShShell

Set WShShell=CreateObject("WScript.Shell")
 
If Time > "07:45:00" And Time < "11:45:00" Then	
	MsgBox("USUÁRIO LOGADO")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Libera\HoraLiberadaManha.vbs")
	
ElseIf Time > "13:45:00" And Time < "17:45:00" Then
	MsgBox("USUÁRIO LOGADO")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Libera\HoraLiberadaTarde.vbs")
	
ElseIf Time > "18:45:00" And Time < "22:15:00" Then
	MsgBox("USUÁRIO LOGADO")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Libera\HoraLiberadaNoite.vbs")

ElseIf Time > "22:15:00" And Time < "23:59:00" Then
	MsgBox ("SEM CONEXÃO")
	WShShell.Run("%Userprofile%\Desktop\Programa_ Intervalo\Intervalos_ON.vbs")

	'WShShell.Run("%Userprofile%\Desktop\Libera\HoraLiberadaMadrugada.vbs")

ElseIf Time > "00:00:00" And Time < "07:45:00" Then
	MsgBox ("SEM CONEXÃO")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Intervalos_ON.vbs")

	'WShShell.Run("C:\Users\eduardo\Desktop\Libera\HoraLiberadaMadrugada.vbs")
	
Else
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Intervalos_ON.vbs")
	


End If
