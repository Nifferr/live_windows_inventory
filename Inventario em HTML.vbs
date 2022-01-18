'Descricao: Inventario automatizado de máquinas e servidores
'Autor: Nicolas Flores Ferreira	
'Data : Julho/2013
'Atualizado: Julho/2021
'Versão: V1.0


On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set objNetwork = CreateObject("WScript.Network")
    STRcomputer = objNetwork.ComputerName
    STRTipoServer = "MS"
    'Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
    On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objtextfile = objFSO.deleteFile(STRcomputer & ".html")
    Set objtextfile = objFSO.CreateTextFile(STRcomputer & ".html", Forwritting)



    Objtextfile.writeline "<!DOCTYPE html>"
    Objtextfile.writeline "<head>"
    Objtextfile.writeline "<link rel='stylesheet' href='css/bootstrap.min.css' type='text/css' media='screen' />"
    Objtextfile.writeline "<link rel='stylesheet' href='css/bootstrap-responsive.min.css' type='text/css' media='screen' />"
    Objtextfile.writeline "<link rel='stylesheet' href='css/flexslider.css' type='text/css' media='screen' />"
    Objtextfile.writeline "<link rel='stylesheet' href='css/style.css' type='text/css' />"
    Objtextfile.writeline "<link href='http://fonts.googleapis.com/css?family=Open+Sans:300italic,400italic,800italic,400,300,800,700,600' rel='stylesheet' type='text/css'>"
    Objtextfile.writeline "</head>"
    Objtextfile.writeline "<body>"
	
    Objtextfile.writeline "<div class='head-section'>"
    Objtextfile.writeline "<div class='container'>"
    Objtextfile.writeline "<div class='span6 pull-left'>"

    Lin_log = 4
    Set objNetwork = CreateObject("WScript.Network")

    objtextfile.WriteLine "<h2>" & STRcomputer & "</h2>"

    Objtextfile.writeline "<!--<p>Introduce your app in few seconds to your customers in a clean and beautiful way. Get it right now.</p>-->"
    Objtextfile.writeline "<ol>"
    Objtextfile.writeline "<li><a href='#SO'>Sistema Operacional</a></li>"
    Objtextfile.writeline "<li><a href='#proc'>Processadores</a></li>"
    Objtextfile.writeline "<li><a href='#bios'>Bios e Hardware</a></li>"
    Objtextfile.writeline "<li><a href='#mem'>Memória</a></li>"
    Objtextfile.writeline "<li><a href='#rede'>Configurações de Rede</a></li>"
    Objtextfile.writeline "<li><a href='#tcp'>Configurações TCP/IP</a></li>"
    Objtextfile.writeline "<li><a href='#discos'>Configurações de Discos</a></li>"
    Objtextfile.writeline "<li><a href='#controladores'>Placas  Controladores</a></li>"
    Objtextfile.writeline "<li><a href='#backup'>Unidade de Backup</a></li>"
    Objtextfile.writeline "<li><a href='#usuarios'>Usuarios Locais</a></li>"
    Objtextfile.writeline "<li><a href='#software'>Softwares Instalados</a></li>"
    Objtextfile.writeline "<li><a href='#servicos'>Status dos Serviços</a></li>"
    Objtextfile.writeline "<li><a href='#compartilhamentos'>Compartilhamentos  Locais</a></li>"
    Objtextfile.writeline "<li><a href='#impressora'>Impressoras Locais</a></li>"
    Objtextfile.writeline "<li><a href='#portas'>Portas de Impressora</a></li>"
    Objtextfile.writeline "<li><a href='#event'>Event Viewer</a></li>"
    Objtextfile.writeline "</ol>"

    Objtextfile.writeline "<!--<a href='#' class='dl-btn'>nome dentro de caixa</a>-->"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "<div class='span4 pull-right'>"
    Objtextfile.writeline "<div class='flexslider'>"
    Objtextfile.writeline "<ul class='slides'>"
    Objtextfile.writeline "<li><img src='img/slider/si1.png' /></li>"
    Objtextfile.writeline "<li><img src='img/slider/si3.png' /></li>"
    Objtextfile.writeline "<li><img src='img/slider/si4.png' /></li>"
    Objtextfile.writeline "</h3>"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "<div class='newsletter-section'>"
    Objtextfile.writeline "<div class='container'>"
    Objtextfile.writeline "<a href='#' class='dl-btn'>Relatório abaixo</a>"
    Objtextfile.writeline "<!--<h4 class=' pull-left'>Signup now and get the latest news from us. (No spam)</h4>-->"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "</div>"
    Objtextfile.writeline "</div>"

    Objtextfile.writeline "<!-- /Header -->"
    Objtextfile.writeline "<!-- Details-Section -->"
    Objtextfile.writeline "<div class='details-section'>"
    Objtextfile.writeline "<div class='desc'>"
    Objtextfile.writeline "<div class='container'>"
    Objtextfile.writeline "<h2>Sistema Operacional</h2>"
					
    Objtextfile.writeline "<p>"
    Objtextfile.writeline "<!--SO--><br>"


    set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" _
                  & STRcomputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objItem In colItems
     dtmConvertedDate.Value = objItem.InstallDate
       dtmInstallDate = dtmConvertedDate.GetVarDate


     objtextfile.WriteLine "<!---" & objItem.Caption & "--->"
       objtextfile.WriteLine objItem.Caption & "<br>"
           objtextfile.WriteLine "Versão: " & objItem.Version & "<br>"
       objtextfile.WriteLine "Service Pack: " & objItem.ServicePackMajorVersion & "<br>"
       objtextfile.WriteLine "Outras descrições: " & objItem.OtherTypeDescription & "<br>"
     objtextfile.WriteLine "Boot Device: " & objItem.BootDevice & "<br>"
     objtextfile.WriteLine "Diretorio Instalação: " & objItem.WindowsDirectory & "<br>"
     objtextfile.WriteLine "Data de Instalação: " & dtmInstallDate & "<br>"
     objtextfile.WriteLine "Organização: " & objItem.Organization & "<br>"
     objtextfile.WriteLine "Usuario Registrado: " & objItem.RegisteredUser & "<br>"
     objtextfile.WriteLine "Serial Number: " & objItem.SerialNumber & "<br>"
   Next
   Set ZoneSet = GetObject("winmgmts:").InstancesOf ("Win32_TimeZone")
   for each System in ZoneSet
    objtextfile.WriteLine "Time Zone:" & System.StandardName
   next

   objtextfile.WriteLine "</p>"
   objtextfile.WriteLine "</p>"
   objtextfile.WriteLine "</div>"
   objtextfile.WriteLine "</div>"
   objtextfile.WriteLine "<div class='container feature-2'>"

   '
   ' Coletando Processadores
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
     objtextfile.WriteLine "<center><h3><a name='#proc'>Processadores</a></h3>"

   For Each objItem In colItems
     objtextfile.WriteLine objItem.Description & "<br>"
     objtextfile.WriteLine objItem.Name & "<br>"
     objtextfile.WriteLine objItem.MaxClockSpeed & " MHZ</center>"
   Next

   '
   ' Coletando Informações da BIOS
   Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
     objtextfile.WriteLine "<center><h3><a name='#bios'>BIOS / Hardware</a></h3>"

   For Each objbios In colBIOS
     objtextfile.WriteLine "Fabricante: " & objbios.Manufacturer & "<br>"
     objtextfile.WriteLine "Serie: " & objbios.SerialNumber & "<br>"
     objtextfile.WriteLine objbios.Name & "<br>"
     objtextfile.WriteLine "Release Date: " & (Mid(objbios.ReleaseDate, 7, 2)) & "/" & (Mid(objbios.ReleaseDate, 5, 2)) & "/" & (Left(objbios.ReleaseDate, 4)) & "<br>"
     objtextfile.WriteLine "SMBIOS Version: " & objbios.SMBIOSBIOSVersion & "<br>"
     objtextfile.WriteLine "BIOS Version: " & CStr(objbios.Version) & "<br>"
   Next

   ' Coletando Modelo do Equipamento
   Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
   For Each objComputer In colSettings
     objtextfile.WriteLine "Modelo Equipamento: " & objComputer.Model & "</center>"
   Next


   w_tipos = Array("Unknown", "Other", "DRAM,Synchronous DRAM", "Cache DRAM", "EDO,EDRAM", "VRAM", "SRAM", _
          "RAM", "ROM", "Flash", "EEPROM", "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", _
          "RDRAM", "DDR")
   totalslots = 0

   'coletando informações da memória
   objtextfile.WriteLine "<center><h3><a name='#mem'>Memória</a></h3>"

   Set SlotMem = objWMIService.ExecQuery _
     ("Select * from Win32_PhysicalMemoryArray")
     For Each objItem In SlotMem
       totalslots = totalslots + objItem.MemoryDevices
     Next
   totalpentes = 0
   Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
   cont_mem=0
   For Each objItem In colItems
     objtextfile.WriteLine objItem.Tag
     objtextfile.WriteLine " " & Int(objItem.Capacity / 1024 / 1024)
     objtextfile.WriteLine " " & w_tipos(objItem.MemoryType)
     objtextfile.WriteLine " " & objItem.BankLabel
     objtextfile.WriteLine " " & "Ativa<br>"
     totalpentes = totalpentes + 1
   cont_mem = cont_mem + objItem.Capacity
   Next

     objtextfile.WriteLine "<!--" & Int(cont_mem / 1024 / 1024) & "-->"

   For i = totalpentes + 1 To totalslots
     objtextfile.WriteLine "<br>"
     objtextfile.WriteLine "Physical Memory "  & i - 1
     objtextfile.WriteLine "Vazio"
     objtextfile.WriteLine "<br>"
   Next
   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery( _
     "SELECT * FROM Win32_PageFileUsage",,48)
   objtextfile.WriteLine "<br><h3>Arquivo de Paginação</h3><br>"
   objtextfile.WriteLine "<p><table border=2 width=500>"
   objtextfile.WriteLine "<tr><td>localização</td><td>Tamanho</td><td>Utilização Atual</td><td>" & _
   "Pico de Utilização</td></tr>"
   For Each objItem In colItems
     wlinha = " "
     wlinha = "<tr><td>"
     wlinha=wlinha & objItem.Description & "</td><td>"
     wlinha=wlinha & objItem.AllocatedBaseSize & "MB" & "</td><td>"
     wlinha=wlinha & objItem.CurrentUsage & "MB"& "</td><td>"
     wlinha=wlinha & objItem.PeakUsage & "MB" & "</td></tr>"
     objtextfile.WriteLine wlinha
   Next
   objtextfile.WriteLine "</table>"
   objtextfile.WriteLine "</p></center>"

   '
   ' Coletando Placas de Rede
     objtextfile.WriteLine "<center><h3><a name='#rede'>Componentes de Rede</a></h3>"

   w_Status = Array("Device is working properly.", _
   "Device is not configured correctly.", _
   "Windows cannot load the driver for this device.", _
   "Driver for this device might be corrupted, or the system may be low on memory or other resources.", _
   "Device is not working properly. One of its drivers or the registry might be corrupted.", _
   "Driver for the device requires a resource that Windows cannot manage.", _
   "Boot configuration for the device conflicts with other devices.", _
   "Cannot filter.", "Driver loader for the device is missing.", _
   "Device is not working properly; the controlling firmware is incorrectly reporting the resources for the device.", _
   "Device cannot start.", "Device failed.", "Device cannot find enough free resources to use.", _
   "Windows cannot verify the device's resources.", "Device cannot work properly until the computer is restarted.", _
   "Device is not working properly due to a possible re-enumeration problem.", _
   "Windows cannot identify all of the resources that the device uses.", _
   "Device is requesting an unknown resource type.", "Device drivers need to be reinstalled.", _
   "Failure using the VxD loader.", "Registry might be corrupted.", _
   "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device.", _
   "Device is disabled.", "System failure. If changing the device driver is ineffective, see the hardware documentation.", _
   "Device is not present, not working properly", _
   "Windows is still setting up the device.", "Windows is still setting up the device.", _
   "Device does not have valid log configuration.", "Device drivers are not installed.", _
   "Device is disabled; the device firmware did not provide the required resources.", _
   "Device is using an IRQ resource that another device is using.", _
   "Device is not working properly; Windows cannot load the required device drivers.")

   w_statusinfo = Array("Disconnected", "Connecting", "Connected", "Disconnecting", _
              "Hardware Not present", "Hardware disabled", "Hardware malfunction", _
              "Media disconnected", "Authenticating", "Authentication succeeded", _
              "Authentication failed", "Invalid Address", "Credentials required")


   Set colItems2 = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") '
   objtextfile.WriteLine "<p>"
   objtextfile.WriteLine "<table border=2 width=1100>"
   objtextfile.WriteLine "<tr><td>Tipo de Adaptador</td><td>Descriçao</td><td>Mac Address</td><td>" & _
   "Fabricante</td><td>Status</td><td>Nome da Conexao</td><td>Velocidade</td></tr>"
   For Each objItem In colItems2
     If IsNull(objItem.AdapterType) then
      wTipo = "-"
     else
      wTipo = objItem.AdapterType
     end if
     If IsNull(objItem.MACAddress) then
      wMac = "- "
     Else
      wMac = objItem.MACAddress
     end if
   '  If IsNull(objItem.StatusInfo) then
   '   wMac = "-"
   '  Else
   '   wMac = objItem.StatusInfo
   ' end if
     If IsNull(objItem.NetConnectionID) Then
      wconnection = "-"
     Else
      wconnection = objItem.NetConnectionID
     end if
     If IsNull(objItem.speed) Then
      wvelocidade = "-"
     Else
      wvelocidade = objItem.speed
     end if

     objtextfile.WriteLine "<tr><td>" & wTipo & "</td><td>" & objItem.Description & "</td><td>" & _
     wMac & "</td><td>" & objItem.Manufacturer & "</td><td >" & _
     w_Status(objItem.ConfigManagerErrorCode) & "</td><td>" & wconnection & "</td><td>" & wvelocidade & "</td></tr>"

   Next
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "</p>"

   objtextfile.writeline "<center><h3><a name='#tcp'>Endereços TCP/IP</a></h3>"
   Set colftp = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

   'objtextfile.writeline "<table border=2 width=100%>"
   'objtextfile.writeline "<tr><td>Placa de rede</td><td>Endereco TCP/IP</td><td>Mascara</td><td>" & _
   '"Tipo IP</td><td>Default Gateway</td><td>DNS Server</td><td>Wins Server</td></tr>"


'Nome placa
For Each objftp In colftp
  wlinha = " "
  wlinha = "<table border=2 width=100% bgcolor=#666666>"
  wlinha = wlinha + "<tr><td width=15%>Placa</td><td><!---" & objftp.Description & "-->" + objftp.Description + "</td><td>"

  'DHCP ou Estatico
  If objftp.dhcpenabled Then
    wlinha = wlinha + "IP Dinamico" + "</td><td>"
  Else
    wlinha = wlinha + "IP Estatico" + "</td><td>"
  End If

wlinha = wlinha + "</table>"
wlinha = wlinha + "<table border=1 width=100%>"

'End_IP
  wlinha = wlinha + "<tr><td width=15%>IP</td><td>"
  strIP = 1
  For Each StrIPaddress In objftp.IPAddress
    wlinha = wlinha + "<!--" & strIP & StrIPaddress &  strIP & "-->" + StrIPaddress + "</td><td>"
    strIP = strIP + 1
    'objtextfile.writeline "<!End_IP" & StrIPaddress & "Fim_IP>"
  Next

'Mascara
  wlinha = wlinha + "<tr><td width=15%>Mascara</td><td>"
  For Each strIPSubnet In objftp.IPSubnet
    wlinha = wlinha + strIPSubnet + "</td><td>"
  Next

'Gateway
  wlinha = wlinha + "<tr><td width=15%>Gateway</td><td>"
  For Each strDefaultIPGatewaY In objftp.DefaultIPGateway
    If IsEmpty(strDefaultIPGatewaY) Then
      wlinha = wlinha + "0.0.0.0" + "<br>"
    Else
      wlinha = wlinha + strDefaultIPGatewaY + "<br>"
    End If
  Next
  wlinha = wlinha + "</td><td>"

'DNS
  wlinha = wlinha + "<tr><td width=15%>DNS</td><td>"
  For Each strDNSServer In objftp.DNSServerSearchOrder
    If IsEmpty(strDNSServer) Then
      wlinha = wlinha + "0.0.0.0" + "<br>"
    Else
      wlinha = wlinha + strDNSServer + "<br>"
    End If
  Next
  wlinha = wlinha + "</td><td>"

'Wins
  wlinha = wlinha + "<tr><td width=15%>Wins</td><td>"
  If IsNull(objftp.WINSPrimaryServer) Then
    wlinha = wlinha + "0.0.0.0" + "<br>"
  Else
    wlinha = wlinha + objftp.WINSPrimaryServer + "<br>"
  End If
  If IsNull(objftp.WINSSecondaryServer) Then
    wlinha = wlinha + "0.0.0.0" + "</td><td>"
  Else
    wlinha = wlinha + objftp.WINSSecondaryServer + "</td><td>"
  End If

wlinha = wlinha + "</table>"
objtextfile.writeline wlinha
Next

   objtextfile.writeline "</table>"

   '
   objtextfile.WriteLine "<center>"
     objtextfile.WriteLine "<h3><a name='#discos'> Discos Locais</a></h3>"
     objtextfile.WriteLine "<table border=2 width=400>"
     objtextfile.WriteLine "<tr><td>Unidade</td><td>Tipo</td></tr>"
     Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colDisks = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
      For Each objDisk in colDisks
         W_ID= objDisk.DeviceID
     Select Case objDisk.DriveType
       Case 1
         W_ID1 = "Tipo de Disco não Detectado"
       Case 2
         W_ID1 = "Disco removível ou Disquete"
       Case 3
         W_ID1 = "Disco Rígido Local"
       Case 4
         W_ID1 ="Drive de Rede"
       Case 5
         W_ID1 = "Unidade de CD"
       Case 6
         W_ID1 ="RAM disk."
       Case Else
         W_ID1 = "Tipo de Disco não Detectado"
     End Select
     objtextfile.WriteLine "<tr><td>" & W_ID & "</td><td>" & W_ID1 & "</td></tr>"
   Next
   objtextfile.WriteLine "</table>"
   objtextfile.WriteLine "<br>"


   '
     objtextfile.WriteLine "<h3><a name='#discos'> Detalhe dos Discos Locais</a></h3>"
     Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colDiskDrives = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
     strDisk = 1
   For each objDiskDrive in colDiskDrives

     objtextfile.WriteLine "Caption:" & ltrim(objDiskDrive.Caption ) & "<br>"
     objtextfile.WriteLine "Device ID:" & ltrim(objDiskDrive.DeviceID ) & "<br>"
     objtextfile.WriteLine "Interface Type:" & ltrim(objDiskDrive.InterfaceType) & "<br>"
     objtextfile.WriteLine "Manufacturer:" & ltrim(objDiskDrive.Manufacturer) & "<br>"
     objtextfile.WriteLine "Model:" & ltrim(objDiskDrive.Model) & "<br>"
     objtextfile.WriteLine "Name:" & ltrim(objDiskDrive.Name) & "<br>"
     objtextfile.WriteLine "Partitions:" & ltrim(objDiskDrive.Partitions) & "<br>"
     objtextfile.WriteLine "SCSI Bus:" & ltrim(objDiskDrive.SCSIBus) & "<br>"
     objtextfile.WriteLine "SCSI Logical Unit:" & ltrim(objDiskDrive.SCSILogicalUnit) & "<br>"
     objtextfile.WriteLine "SCSI Port:" & ltrim(objDiskDrive.SCSIPort) & "<br>"
     objtextfile.WriteLine "SCSI TargetId:" & ltrim(objDiskDrive.SCSITargetId) & "<br>"
     objtextfile.WriteLine "<!--Disco" & strDisk & int(objDiskDrive.Size/1024/1024) & "Fim_Disco" & strDisk & "--><br>"
     objtextfile.WriteLine "Size:" & int(objDiskDrive.Size/1024/1024) & "MB<br>"
     objtextfile.WriteLine "Status:" & ltrim(objDiskDrive.Status) & "<br>"
	 strDisk = strDisk+1
   Next

   '
   ' Discos lógicos
   Const HARD_DISK = 3

     objtextfile.WriteLine "<h3><a name='#discos'> Partições</a></h3>"
   Set colDiskDrives = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
   objtextfile.WriteLine "<table border=2 width=500>"
   objtextfile.WriteLine "<tr><td>Unidade</td><td>Tamanho (MB)</td><td>Tipo Partiçao" & _
   "</td><td>Espaço Livre(MB)</td></tr>"
   For Each objdisk In colDiskDrives
     objtextfile.WriteLine "<tr><td>" & objdisk.DeviceID & "</td><td>" & " " _
     & Int(objdisk.Size / 1024 / 1024) & "</td><td>" & " " _
     & objdisk.FileSystem & "</td><td>" _
      & int(objDisk.FreeSpace/(1024*1024)) & "</td></tr>"
   Next
   objtextfile.WriteLine "</table>"


   '

     objtextfile.WriteLine "<h3><a name='#discos'> Unidade de CDROM</a></h3>"
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive")
   For Each objItem in colItems
      objtextfile.WriteLine "<tr><td>" & objItem.Caption & "</td></td>"
   Next


   '
   'Controladoras de Disco"
   objtextfile.WriteLine "<h3><a name='#controladoras'> Controladoras de Disco</a></h3>"
   objtextfile.WriteLine "<table border=2 width=600>"
   objtextfile.WriteLine "<tr><td>Controladora</td><td>Driver</td><td>Fabricante</td><td>Status</td></tr>"
   Set wplaca = objWMIService.ExecQuery("Select * from Win32_SCSIController")
   For Each objplaca In wplaca
   objtextfile.WriteLine "<tr><td>" & objplaca.Name & "</td><td>" & " " _
   & objplaca.DriverName & "</td><td>" & " " _
   & objplaca.Manufacturer & "</td><td>" & " " _
   & objplaca.Status & "</td></tr>"
   Next
   objtextfile.WriteLine "</table>"

   '
   objtextfile.WriteLine "<h3><a name='#backup'> Unidade de Backup</a></h3>"
   objtextfile.WriteLine "<table>"
   Set wfita = objWMIService.ExecQuery("Select * from Win32_TapeDrive")
   For Each objfita In wfita
     objtextfile.WriteLine "<tr><td>" & objfita.Caption & "</td></tr>"
   Next
   objtextfile.WriteLine "</table>"

   '
   'Coleta usuarios locais
   Set colAccounts = GetObject("WinNT://" & STRcomputer & "")
   Set colGroups = GetObject("WinNT://" & STRcomputer & "")
   colAccounts.Filter = Array("user")
   colGroups.Filter = Array("group")
   objtextfile.WriteLine "<h3><a name='#usuarios'> Usuarios Locais</a></h3>"
   objtextfile.WriteLine "<table border=2 width=1000>"
    objtextfile.WriteLine "<tr><td>Login</td><td>Nome Completo</td><td>Descriçao" & _
   "</td><td>Status</td><td>Grupos</td></tr>"
   For Each objUser In colAccounts
     Wgrupo = " "
     If objUser.AccountDisabled then
    wstatus = "Inativa"
     Else
    wstatus = "Ativa"
     End if
     For Each objGroup In colGroups
       For Each objuserMBR In objGroup.Members
         If objuserMBR.Name = objUser.Name Then
           Wgrupo = Wgrupo & objGroup.Name & "<br>"
         End If
       Next
     Next
     If Len(objUser.FullName) <= 1 Then
       wfulname = "-"
     Else
      wfulname = objUser.FullName
     End If
     If Len(objUser.Description) <= 1 then
    wdescricao = "-"
     else
    wdescricao = objUser.Description
     end if
     objtextfile.WriteLine "<tr><td>" & objUser.Name & "</td><td>" & " " _
     & wfulname & "</td><td>" _
     & wdescricao & "</td><td>" _
     & wstatus & "</td><td>" _
     & Wgrupo & "</td></tr> "
   Next
   objtextfile.WriteLine "</table>"

   '
   'Coleta Software
   Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
   strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
   strEntry1a = "DisplayName"
   strEntry1b = "QuietDisplayName"
   strEntry2 = "InstallDate"
   strEntry3 = "VersionMajor"
   strEntry4 = "VersionMinor"
   strEntry5 = "EstimatedSize"

   objtextfile.WriteLine "<h3><a name='#software'> Softwares Instalados</a></h3>"
   Objtextfile.writeline "</h3>"
   'objtextfile.WriteLine "<table border=2>"
    objtextfile.WriteLine "<tr><td>Software / Hotfix</td></tr>"
   Set objReg = GetObject("winmgmts://" & STRcomputer & _
    "/root/default:StdRegProv")
   objReg.EnumKey HKLM, strKey, arrSubkeys
   For Each strSubkey In arrSubkeys
   '  wlinha = "<tr><td>"
     wlinha=""
    intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
     strEntry1a, strValue1)
    If intRet1 <> 0 Then
     objReg.GetStringValue HKLM, strKey & strSubkey, _
      strEntry1b, strValue1
    End If
    If strValue1 <> "" Then
     wlinha = wlinha + strValue1 '+ "</td><td>"
    End If
    objReg.GetDWORDValue HKLM, strKey & strSubkey, _
    strEntry3, intValue3
    objReg.GetDWORDValue HKLM, strKey & strSubkey, _
     strEntry4, intValue4
    If intValue3 <> "" Then
      wlinha = wlinha + intValue3 + "." + intValue4 '"</td></tr>"
    End If
    If Len(wlinha) > 8 Then
     objtextfile.WriteLine wlinha + "<br>"
    End If
   Next

   'objtextfile.WriteLine "</table>"

   '
   'Coleta Serviços
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")
    Set colListOfServices = objWMIService.ExecQuery _
       ("Select * from Win32_Service")

   objtextfile.WriteLine "<h3><a name='#servicos'> Serviços</a></h3>"

   objtextfile.WriteLine "<table border=2 width=800>"
   objtextfile.WriteLine "<tr><td><b>Serviço</b></td><td><b>Startup</b></td><td><b>" & _
   "Status</b></td><td><b>Usuario de Startup</b></b></td></tr>"
   For Each objService In colListOfServices
     if objService.State = "Stopped" then
    wfonte= "<font color=red>"
     else
    wfonte = "<font>"
     end if
     objtextfile.WriteLine "<tr><td><i>" & wfonte & objService.Caption & "</font></i></td><td><i>" & _
     wfonte & objService.StartMode & "</font></i></td><td><i>" & _
     wfonte & objService.State & "</font></i></td><td><i>" & _
     wfonte & objService.StartName & "</font></i></td></tr>"
   Next

   objtextfile.WriteLine "</table>"

   '
   'Coleta Tamanho dos Diretorios
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")

   Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")


   objtextfile.WriteLine "<h3><a name='#compartilhamentos'>Compartilhamentos</a></h3>"

   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b>Compartilhamento</b></td><td><b>" & _
   "Caminho</b></b></td></tr>"

   For Each objShare In colShares
     objtextfile.WriteLine "<tr><td><i>" & objShare.Name & "</i></td><td><i>" _
     & objShare.Path & "</i></td></tr>"
   Next

   objtextfile.WriteLine "</table>"

   '
   'Coleta Impressora
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")
   Set colInstalledPrinters = objWMIService.ExecQuery _
     ("Select * from Win32_PrinterDriver")

   objtextfile.WriteLine "<h3><a name='#impressora'> Drivers de Impressoras</a></h3>"

   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b>Impressora Local</b></b></td></tr>"

   For Each objPrinter In colInstalledPrinters
     objtextfile.WriteLine "<tr><td><i>" & objPrinter.Name & "</i></td></tr>"
   Next

   objtextfile.WriteLine "</table>"

   '
   'Coleta portas de impressora
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")

   Set colPorts = objWMIService.ExecQuery _
     ("Select * from Win32_TCPIPPrinterPort")

   objtextfile.WriteLine "<h3><a name='#portas'> Portas de Impressora</a></h3>"

   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b>Impressora Local</b></td><td><b>Endereco Host</b></td><td><b>" & _
   "Porta</b></td><td><b>Protocolo</b></b></td></tr>"
   For Each objPort In colPorts
     objtextfile.WriteLine "<tr><td><i>" & objPort.Description & "</i></td><td><i>" & _
     objPort.HostAddress & "</i></td><td><i>" & _
     objPort.Name & "</i></td><td><i>" & _
     objPort.PortNumber & "</i></td><td><i>" & _
     objPort.Protocol & "</i></td></tr>"
   Next

   objtextfile.WriteLine "</table>"

   '

   objtextfile.WriteLine "<h3><a name='#event'>Event Viewer</a></h3>"

   objtextfile.WriteLine "<table border=2 width=100>"
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate,(Security)}!\\" & _
       strComputer & "\root\cimv2")
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='Security'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>Security </i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i>"
   Next
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='application'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>Application</i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i>"
   Next
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='system'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>System</i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i></tr>"
   Next


   objtextfile.WriteLine "</table>"


   Objtextfile.writeline "</p>"

   Objtextfile.writeline "</div>"
   Objtextfile.writeline "</div>"
   Objtextfile.writeline "<!-- /Details-Section -->"
   Objtextfile.writeline "<!-- Footer-Section -->"
   Objtextfile.writeline "<!-- /Footer-Section -->"
   Objtextfile.writeline "<!-- Javascript -->"
   Objtextfile.writeline "<script type='text/javascript' src='js/jquery-1.8.2.min.js'></script>"
   Objtextfile.writeline "<script type='text/javascript' src='js/bootstrap.min.js'></script>"
   Objtextfile.writeline "<script type='text/javascript' src='js/jquery.flexslider-min.js'></script>"
   Objtextfile.writeline "<script type='text/javascript' src='js/custom.js'></script>"
   Objtextfile.writeline "</body>"
   Objtextfile.writeline "</html>"


wscript.echo "Inventário Concluido"