Dim objExcel, objSpread, intRow, strCN, strSam, strSheet,grupo1,grupo2

strSheet = "C:\script\addGrupo\UserINGrupo.xls"
grupo1 = InputBox("DIGITE O GRUPO NO QUAL DESEJAS INSERIR OS USU�RIOS")
grupo2 = InputBox("DIGITE O GRUPO DE PROXY QUE DESEJAS INSERIR OS USU�RIOS")

Set objNet = CreateObject("WScript.Network" ) 

' Open the Excel spreadsheet
Set objExcel = CreateObject("Excel.Application")
Set objSpread = objExcel.Workbooks.Open(strSheet)
intRow = 3 'Row 1 often contains headings
Do Until objExcel.Cells(intRow,1).Value = ""
strSam = Trim(objExcel.Cells(intRow, 1).Value)

strNetBIOSDomain = objNet.UserDomain 
strComputer = objNet.ComputerName 

Set objGroup = GetObject("WinNT://" & strComputer & "/"& grupo1 &",group") 
Set objGroup2 = GetObject("WinNT://" & strComputer & "/"& grupo2 &",group") 
Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" & strSam & ",user" ) 

' Ignora se o usu�rio j� pertencer ao grupo
On Error Resume Next 
'wscript.Echo "J� existe esse no grupo "
objGroup.Add(objUser.ADsPath) 
objGroup2.Add(objUser.ADsPath) 
On Error Goto 0 

intRow = intRow + 1
Loop
objExcel.Quit

wscript.Echo "Usu�rio inserido no grupo com Sucesso "
