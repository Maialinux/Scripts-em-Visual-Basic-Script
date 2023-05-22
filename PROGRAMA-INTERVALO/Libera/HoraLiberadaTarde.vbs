Dim timeTempo
Dim timeTempoAtual
Dim diferencaTime,diferencaTime2
Dim WShShell

timeTempo="17:45:00"

timeTempoAtual=time

diferencaTime=DateDiff("n",timeTempoAtual,timeTempo)
diferencaTime2=DateDiff("s",timeTempoAtual,timeTempo)

Set WShShell=CreateObject("WScript.Shell")
 
If diferencaTime > 0 Then	
	WScript.Echo "COMPUTADOR LIBERADO ATÉ AS 17:45hrs FALTAM "&diferencaTime&" MINUTO(s) PARA O ENCERRAMENTO"	
	WScript.Sleep(diferencaTime2&"000")
	WShShell.Run("logoff")
Else
WShShell.Run("logoff")

End If
Set WShShell=Nothing 