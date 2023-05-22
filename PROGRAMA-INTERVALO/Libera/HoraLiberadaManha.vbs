Dim timeTempo
Dim timeTempoAtual
Dim diferencaTime,diferencaTime2
Dim WShShell

timeTempo="11:45:00"

timeTempoAtual=time

diferencaTime=DateDiff("n",timeTempoAtual,timeTempo)
diferencaTime2=DateDiff("s",timeTempoAtual,timeTempo)

Set WShShell=CreateObject("WScript.Shell")
 
If diferencaTime > 0 then	
	WScript.Echo "COMPUTADOR LIBERADO ATÉ AS 11:45hrs FALTAM "&diferencaTime&" MINUTOS(s) PARA O ENCERRAMENTO"	
	WScript.Sleep(diferencaTime2&"000")
	WShShell.Run("logoff")
Else
WShShell.Run("logoff")

End If
Set WShShell=Nothing 