' ----------------------------------------------------------------------------------
' CITRA IT - EXCELÊNCIA EM TI
' Script para identificar na rede as máquinas configuradas com IP Fixo
' Autor: luciano@citrait.com.br
' Data: 11/07/2020  Versão: 1.0 release inicial
' Homologado para estações Windows 7, 8, e 10.
' Importante: Altere o caminho NETWORK_PATH para a pasta de rede que será 
'  salvo as informações das máquinas.
' ----------------------------------------------------------------------------------
'Option Explicit
On Error Resume Next

' Caminho da rede onde salvar o arquivo de saida.
' Importante: Deve terminar em barra o caminho!
NETWORK_PATH = "\\SERVER\SHARE$\"


' Obtendo o usuário logado
Set objShell	= CreateObject("WSCript.Shell")
ConnectedUser	= objShell.ExpandEnvironmentStrings("%USERNAME%")
ComputerName	= objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")



' Consultando o WMI para listar as placas de rede
Set wmiServices = GetObject("winmgmts:\\.\root\cimv2")
Set objResult = wmiServices.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = TRUE AND DHCPEnabled = FALSE")
DHCPEnabled = True
Hostname = ""
IPAddr = ""

' Verificando a configuração de cada placa de rede encontrada
If objResult.Count > 0 Then
	DHCPEnabled = False
	For Each obj in objResult
		For Each ip in obj.IPAddress
			IPAddr = IPAddr & " " & ip
		Next
		If obj.DNSHostName <> "" Then
			Hostname = obj.DNSHostName
		End if
	Next
End If


' Reunindo informações para salvar no arquivo de saída
txtData = "Computername=" & Computername & vbCrLf & vbCrLf
txtData = "User=" & Username & vbCrLf & vbCrLf
If DHCPEnabled = False Then
	txtData = txtData & "DHCPEnabled=0" & Hostname & vbCrLf & vbCrLf
Else
	txtData = txtData & "DHCPEnabled=1" & Hostname & vbCrLf & vbCrLf
End If
txtData = txtData & "IPAddr=" & IPAddr & vbCrLf & vbCrLf
txtData = txtData & "DateTime=" & Now()

' Criando o arquivo de saída na rede
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSO.Open(NETWORK_PATH & Computername & ".txt", 2, True)
outputFile.Write txtData
outputFile.Close

' Limpeza de Variaveis
Set objFSO		= Nothing
Set outputFile	= Nothing
Set objResult	= Nothing
Set wmiServices	= Nothing

' Finalizando o script
Wscript.Quit




