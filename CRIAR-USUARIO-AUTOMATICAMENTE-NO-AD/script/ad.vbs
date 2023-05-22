Option Explicit
Dim objRootLDAP, objContainer, objUser, objShell
Dim objExcel, objSpread, intRow
Dim strUser, strOU, strSheet, strSheet2
Dim strCN, strSam, strFirst, strLast, strPWD, stroffice, strmail, strtitle, strdepartment, strprincipalname, strTitulo, strMensagem, strcompany, strdescription, strtelephoneNumber
Dim strProfilePath, strHomeDrive, strHomeDirectory, strScriptPath

strOU = "OU=Manha, OU=Professores, OU=etcr ," 

strSheet = "C:\script\Usersbah.xls"


' Bind to Active Directory, Users container.

Set objRootLDAP = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://"& strOU & objRootLDAP.Get("defaultNamingContext"))

' Open the Excel spreadsheet
Set objExcel = CreateObject("Excel.Application")
Set objSpread = objExcel.Workbooks.Open(strSheet)
intRow = 3 'Row 1 often contains headings

' Here is the 'DO…Loop' that cycles through the cells
' Note intRow, x must correspond to the column in strSheet
Do Until objExcel.Cells(intRow,1).Value = ""
strSam = Trim(objExcel.Cells(intRow, 1).Value)
strCN = Trim(objExcel.Cells(intRow, 2).Value)
strFirst = Trim(objExcel.Cells(intRow, 3).Value)
strLast = Trim(objExcel.Cells(intRow, 4).Value)
strPWD = Trim(objExcel.Cells(intRow, 5).Value)
stroffice = Trim(objExcel.Cells(intRow, 6).Value)
strmail = Trim(objExcel.Cells(intRow, 7).Value)
strprincipalname = strSam & "@dominio.local"
strtitle = Trim(objExcel.Cells(intRow, 8).Value)
strdepartment = Trim(objExcel.Cells(intRow, 09).Value)
strdescription = Trim(objExcel.Cells(intRow, 10).Value)
strtelephoneNumber = Trim(objExcel.Cells(intRow, 11).Value)
strcompany = "Contoso Corporation"
strProfilePath = Trim(objExcel.Cells(intRow, 12).Value)
strHomeDrive = Trim(objExcel.Cells(intRow, 13).Value)
strHomeDirectory = Trim(objExcel.Cells(intRow, 14).Value)
strScriptPath = Trim(objExcel.Cells(intRow, 15).Value)

' Build the actual User from data in strSheet.
Set objUser = objContainer.Create("User", "cn=" & strCN)
objUser.sAMAccountName = strSam
objUser.givenName = strFirst
objUser.sn = strLast
objUser.SetInfo
objUser.physicalDeliveryOfficeName = stroffice
objUser.mail = strmail
objUser.userPrincipalName= strprincipalname
objUser.displayName = strCN
objUser.title = strtitle
objUser.department = strdepartment
objUser.company = strcompany
objUser.description = strdescription
objUser.telephoneNumber = strtelephoneNumber
objUser.profilePath = strProfilePath 
objUser.homeDrive = strHomeDrive
objUser.homeDirectory = strHomeDirectory 
objUser.scriptPath = strScriptPath 

' Separate section to enable account with its password
objUser.userAccountControl = 512
objUser.pwdLastSet = 0
objUser.SetPassword strPWD
objUser.SetInfo

intRow = intRow + 1
Loop
objExcel.Quit

strTitulo = "CONCLUÍDO COM SUCESSO!"
strMensagem = _
"CONTAS DE USUARIO CRIADAS COM ÊXITO!" & vbcrlf & vbcrlf & _
" Não se esqueça de movê-las para a OU específica e em seguida criar suas respectivas mailboxes." & vbcrlf & _
"" & vbcrlf & _
""

'BtnCode = WshShell.Popup(strMensagem, 5, "Informação:", 64 + 0)
msgbox strMensagem, 0 + 64, strTitulo
Dim WShShell
Set WShShell=WScript.CreateObject("WScript.Shell") 
WshShell.run "C:\script\addGrupo\addUser_IN_Group.vbs"
WScript.Quit

