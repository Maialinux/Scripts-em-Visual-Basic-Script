Dim timeTempo
Dim timeTempoAtual
Dim diferencaTime
Dim WShShell

timeTempo="23:59:00"

timeTempoAtual=time

diferencaTime=DateDiff("n",timeTempoAtual,timeTempo)

Set WShShell=CreateObject("WScript.Shell")
 
If diferencaTime > 0 Then	
	WScript.Echo "COMPUTADOR BLOQUEADO"
	WScript.Echo "HORARIO INDISPONIVEL: DAS 22:15hrs ATÉ 23:59hrs"
	WShShell.Run("logoff")
	
Else
Set WShShell=Nothing 

end if
