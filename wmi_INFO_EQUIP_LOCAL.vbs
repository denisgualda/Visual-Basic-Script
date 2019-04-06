'####################################################################################
'####################################################################################
'INFORMACIO GENERAL DE L'EQUIP'							 ############################
'####################################################################################
'####################################################################################


'***************************************************************
'ESCRIU EN FITXER
'***************************************************************
Dim strEscriu
Const ForAppending = 8
Dim strLogFile, strDate
srLogFile = ".\info.txt"


'Set objFSO = CreateObject("Scripting.FileSystemObject")
'if objFSO.FileExists(strLogFile) Then
	'objFSO.DeleteFile(strLogFile)
'End if
'Set objLogFile = objFSO.OpenTextFile(strLogFile, ForAppending, True)


'*************************************************************
'INFORMACIO DE SISTEMA
'*************************************************************
Dim wmiItemsSysinfo
Set wmiService = GetObject("winmgmts:\\" & StrComputer)
Set wmiItemsSysinfo = wmiService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")		'Informacio de l'equip
mess = mess & "INFORMACIO DEL SISTEMA" & VBCrlf
mess = mess & "-----------------------------------------------------------" & VBCrlf
For Each item in wmiItemsSysinfo
	With item
		mess = mess & "- Model: " & .Name & VBCrlf
		mess = mess & "- Fabricant: " & .Vendor & VBCrlf
		mess = mess & "- S/N: " & .IdentifyingNumber & VBCrlf
	End With
Next
Dim wmiItemsOSinfo
Set wmiItemsOSinfo = wmiService.ExecQuery("SELECT * FROM Win32_OperatingSystem")		'Informació sistema operatiu
For each item in wmiItemsOSinfo
	With item
		mess= mess & "- S.O: " & .Caption & " " & .OSArchitecture & VBCrlf
		mess = mess & "Nom equip: " & .CSName & VBCrlf
	End With
Next
mess = mess & VBCrlf


'***************************************************************
'INFORMACIO RAM
'***************************************************************
Dim StrComputer, mess
Dim wmiService, wmiItems, item

StrComputer = "."
'Set wmiService = GetObject("winmgmts:\\" & StrComputer)
Set wmiItems = wmiService.ExecQuery("SELECT * FROM win32_PhysicalMemory")

mess = mess & "MEMORIA RAM" & VBCrlf
mess = mess & "-----------------------------------------------------------" & VBCrlf
For Each item in wmiItems
	With item
		mess = mess & "Modul instalat a: " & .BankLabel & " "  & .DeviceLocator &  VBCrlf & "Capacitat: " & left(.Capacity/1024^3,6) & " GB " & VBCrlf & "Fabricant: " & .ManuFacturer & VBCrlf & "Num.serie: " & .PartNumber & VBCrlf & "Velocitat: " & .Speed & VBCrlf
		mess = mess & VBCrlf
	End With
Next


'***************************************************************
'INFORMACIO DISC DUR
'***************************************************************

mess = mess & VBCrlf
mess = mess & "DISC DUR" & VBCrlf
mess = mess & "-----------------------------------------------------------" & VBCrlf
Set wmiItemsDiskDrive = wmiService.ExecQuery("SELECT * FROM Win32_DiskDrive")

For Each item in wmiItemsDiskDrive
	With item
		mess = mess & "Model: " & .Model & VBCrlf
	End With
Next

'************************
'Consulta sobre Win32_LogicalDisk

Set wmiItemsLogicalDisk = wmiService.ExecQuery("SELECT * FROM Win32_LogicalDisk")


For Each item in wmiItemsLogicalDisk
	With item
		'***************************************
		'left(.Size/1024^3,6) o left(.FreeSpace/1024^3,6) mostrem el resultat de la consulta que retorna en bytes a GB i agafa els 6 primers caràcters.

		mess = mess & "Tamany: "  & left(.Size/1024^3,6) & " GB " & VBCrlf & "Espai Lliure: " & left(.FreeSpace/1024^3,6) & " GB " & VBCrlf 

	End With
Next


'***************************************************************
'INFORMACIO CPU
'***************************************************************

StrComputer = "."
Set wmiService = GetObject("winmgmts:\\" & StrComputer)
Set wmiItems = wmiService.ExecQuery("SELECT * FROM win32_Processor")
 
mess = mess & "PROCESSADOR  " & VBCrlf
mess = mess & "-----------------------------------------------------------" & VBCrlf
For Each item in wmiItems
	With item
		mess = mess & "-" & VBCrlf
		mess = mess & "Processador: " & .Name & .Addresswidth & VBCrlf
		mess = mess & VBCrlf
	End With
Next

Set wmiItems = wmiService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")


'***************************************************************
'INFORMACIO XARXA
'***************************************************************

For Each item in wmiItems
	With item
		mess = mess & "Adaptador de xarxa: " & .Description & VBCrLf 											'**************************VBCrlf --> Serveix per fer intro'
		mess = mess & "-------------------------------------------------------------------" & VBCrlf
		For Each strIPSubnet in .IPSubnet
			subxarxa= subxarxa & strIPSubnet
			'*******************************************************************************************************************************
			'Retallem els 2 últims digits que agafa de la Mascara de subxarxa (64) per deixar unicament els 12 digits vàlids 255.255.255.255
			subxarxa = Left(subxarxa,15)
		Next
		For Each strIPAddress in .IPaddress 
			mess = mess & "IP: " & strIPAddress & "   " & " MAC: " & .MACAddress & " Masc: " & subxarxa & VBCrLf		
		Next
			mess = mess & "DHCP habilitat: " & .DHCPEnabled & VBCrlf
			mess = mess & "-------------------------------------------------------------------" & VBCrlf
	End With


'***************************************************************
'IMPRESORES INSTAL·LADES
'***************************************************************

StrComputer = "."
Set wmiService = GetObject("winmgmts:\\" & StrComputer)
Set wmiItems = wmiService.ExecQuery("SELECT * FROM win32_Printer")
 
mess = mess & "IMPRESSORES  " & VBCrlf
mess = mess & "-----------------------------------------------------------" & VBCrlf
For Each item in wmiItems
	With item
		mess = mess & "-" & VBCrlf
		mess = mess & "Nom impressora: " & .DeviceID & VBCrlf
		mess = mess & "Driver: " & .DriverName & VBCrlf
		mess = mess & VBCrlf
	End With
Next

WScript.Echo mess

strEscriu = mess
objLogFile.WriteLine strEscriu
objLogFile.Close





