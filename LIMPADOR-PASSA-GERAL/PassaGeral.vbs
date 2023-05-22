Option Explicit

'VARIÁVEIS
Dim FSO 
Dim objShell
Dim strCaminhoDoUsuario 

'FORMA DE ADQUIRIR O PODER DE CRIAR,EDITAR,DELETAR PASTAS OU ARQUIVOS
Set FSO=CreateObject("Scripting.FileSystemObject")

'FORMA DE TER ACESSO AO SHELL (COMANDOS DO PROMP DE COMANDO )
Set objShell=CreateObject("WScript.Shell")

'DEFININDO CAMINHO DO USUÁRIO ATRAVÉS DE VARIÁVEIS DE AMBIENTE DO CMD
strCaminhoDoUsuario = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

'FORMA DE DELETAR ARQUIVOS E PASTAS DO USUÁRIO
'NESSE CASO O USUÁRIO É O USUÁRIO ATUAL
 
FSO.DeleteFile(strCaminhoDoUsuario&"\Downloads\*.*"),True
FSO.DeleteFolder(strCaminhoDoUsuario&"\Downloads\*.*"),True
 
'FSO.DeleteFile(strCaminhoDoUsuario&"\Documents\*.*"),True
'FSO.DeleteFolder(strCaminhoDoUsuario&"\Documents\*.*"),True

FSO.DeleteFile(strCaminhoDoUsuario&"\Pictures\*.*"),True
FSO.DeleteFolder(strCaminhoDoUsuario&"\Pictures\*.*"),True

FSO.DeleteFile(strCaminhoDoUsuario&"\Music\*.*"),True   
FSO.DeleteFolder(strCaminhoDoUsuario&"\Music\*.*"),True
 
FSO.DeleteFile(strCaminhoDoUsuario&"\Videos\*.*"),True
FSO.DeleteFolder(strCaminhoDoUsuario&"\Videos\*.*"),True 

FSO.DeleteFile(strCaminhoDoUsuario&"\Desktop\*.*"),True
FSO.DeleteFolder(strCaminhoDoUsuario&"\Desktop\*.*"),True 


MsgBox ("LIMPEZA COMPLETADA NO WINDOWS EXPLORER")
Wscript.quit
