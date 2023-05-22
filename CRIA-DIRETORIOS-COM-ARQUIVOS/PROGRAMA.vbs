'DECLARA AS VARIAVEIS
'*****************************************************************************

Dim objFSO
Dim strDirectory
Dim strArqNovo
Dim objTextFile

'*****************************************************************************

'PEGO O SYSTEMA DE LER,ESCREVER E GRAVAR ARQUIVOS

'*****************************************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")
strDirectory = "C:\\PassaGeral"

If objFSO.FolderExists(strDirectory) Then 
Set objDirectory = objFSO.GetFolder(strDirectory)

Else
Set objDirectory = objFSO.CreateFolder(strDirectory)

End If

strArqNovo = "NewFile.txt"
Set objTextFile = objFSO.CreateTextFile(strDirectory & "\\" & strArqNovo, True,8)

objTextFile.WriteLine("Este é") 
objTextFile.WriteLine("Um novo arquivo de texto") 
objTextFile.Close 

'objDirectory = objFSO.CopyFolder(strDirectory, "C:\\Program Files (x86)\\" ) 

Set objTextFile = Nothing

strArqNovo = "NewFile2.txt"
Set objTextFile = objFSO.CreateTextFile(strDirectory & "\\" & strArqNovo, True,8)

objTextFile.WriteLine("Este é") 
objTextFile.WriteLine("Um novo arquivo 2") 
objTextFile.Close 

objDirectory = objFSO.CopyFolder(strDirectory, "C:\\Program Files (x86)\\" ) 

Set objTextFile = Nothing
Set objDirectory  = Nothing 
Set objFSO = Nothing
'MsgBox ("LIMPEZA COMPLETADA NO WINDOWS EXPLORER")
Wscript.quit
