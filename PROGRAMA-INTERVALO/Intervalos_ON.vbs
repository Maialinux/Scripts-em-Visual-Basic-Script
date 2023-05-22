Dim WShShell

Set WShShell=CreateObject("WScript.Shell")
 
If Time >= "11:45:00" And Time <= "13:45:00" Then	
	MsgBox ("Bom Almoço")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Bloqueio\HoraBloqueioManha.vbs")
	
ElseIf Time >= "17:45:00" And Time <= "18:45:00" Then
	MsgBox ("Bom Café da Tarde")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Bloqueio\HoraBloqueioTarde.vbs")
	
ElseIf Time >= "22:15:00" And Time <= "23:59:00" Then
	MsgBox ("Boa Janta")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Bloqueio\HoraBloqueioNoite.vbs")

ElseIf Time >= "00:00:00" And Time <= "07:45:00" Then
	MsgBox ("Boa Madrugada")
	WShShell.Run("%Userprofile%\Desktop\Programa_Intervalo\Bloqueio\HoraBloqueioMadrugada.vbs")
	

End If
