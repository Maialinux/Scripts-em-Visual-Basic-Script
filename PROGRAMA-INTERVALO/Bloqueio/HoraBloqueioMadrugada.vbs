Dim timeTempo
Dim timeTempoAtual
Dim diferencaTime
Dim WShShell

timeTempo="07:45:00"

timeTempoAtual=time

diferencaTime=DateDiff("n",timeTempoAtual,timeTempo)

Set WShShell=CreateObject("WScript.Shell")
 
If diferencaTime > 0 Then	
	WScript.Echo "COMPUTADOR BLOQUEADO"
	WScript.Echo "HORARIO INDISPONIVEL: DAS 00:00hrs ATÉ 07:45hrs"
	WShShell.Run("logoff")
	
Else
Set WShShell=Nothing 

end if
