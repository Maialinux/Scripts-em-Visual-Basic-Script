Dim timeTempo
Dim timeTempoAtual
Dim diferencaTime
Dim WShShell

timeTempo="13:45:00"

timeTempoAtual=time

diferencaTime=DateDiff("n",timeTempoAtual,timeTempo)

Set WShShell=CreateObject("WScript.Shell")
 
If diferencaTime > 0 then	
	WScript.Echo "COMPUTADOR BLOQUEADO"
	WScript.Echo "HORARIO INDISPONIVEL: DAS 11:45hrs ATÉ 13:45hrs"
	WShShell.Run("logoff")
	
Else
Set WShShell=Nothing 

End If
