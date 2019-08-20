'Creado Por @AnderOkgo'
dim WMI, Mac, NadDescripcion,Ip,ServicePack,PcNombre,SONombre,ProNombre,BoardCompleta,BoardModelo
dim CDRNombre,EquipoSerial, VideoCompleta,SonidoNombre,DiscoCompleta,Usuario,Dominio,MemoriaRam,Fabricante,Modelo,Arquitectura
set WMI = GetObject("winmgmts:\\.\root\cimv2")
Set fs = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("Wscript.Shell")
pc= WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

function RedInfo()
dim Nads, Nad
set Nads = WMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True") 
for each Nad in Nads
if not isnull(Nad.MACAddress) then 
Mac = Nad.MACAddress
NadDescripcion = Nad.description
Ip=Nad.IPAddress(0)
end if
next
end function

function SOInfo()
Set Os = WMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each O in Os
SONombre = O.Caption
PcNombre = O.CSName
ServicePack = SONombre  & " " & O.ServicePackMajorVersion & "." & O.ServicePackMinorVersion
Next
end function

function ProcesadorInfo()
Set pros = WMI.ExecQuery("SELECT * FROM Win32_Processor", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each pro in pros
ProNombre = LTrim(pro.Name)
Next
end function

function BoardInfo()    
Set Bs = WMI.ExecQuery("Select * from Win32_BaseBoard")     
For Each b In Bs  
BoardFabricante = b.Manufacturer
BoardModelo = b.Product
BoardCompleta = BoardModelo& " " & BoardFabricante
Next 
end function

function CDROMInfo()
Set CDS = WMI.ExecQuery("Select * from Win32_CDROMDrive",,48)
For Each CD in CDS
CDRNombre= CD.Caption
Next
end function

function EquipoId()
Set Bioss = WMI.ExecQuery("Select * from Win32_SystemEnclosure")
For Each Bios in Bioss
EquipoSerial=Bios.SerialNumber
Next
end function

function videoInfo()
Set Monitores = WMI.ExecQuery("Select * from Win32_DisplayConfiguration",,48)
For Each monitor in monitores
VideoTrageta = monitor.DeviceName
VideoResolucion = monitor.PelsWidth & " x " & monitor.PelsHeight & " x " & monitor.BitsPerPel & " bits"
VideoCompleta= VideoTrageta & " " & VideoResolucion
Next
end function

function SonidoInfo()
Set Sonidos = WMI.ExecQuery("Select * from Win32_SoundDevice",,48)
For Each Sonido in Sonidos
SonidoNombre = Sonido.Caption
Next
end function

function DiscoDuroInfo()
Set Discos = WMI.ExecQuery("Select * from Win32_DiskDrive",,48)
For Each Disco in Discos
DiscoLabel = Disco.Caption   
Peso= round(Disco.Size/1000000000) & "Gb"
DiscoCompleta= DiscoLabel & " " & peso
Exit For
next
end function

function  SistemaInfo()
Set Sistemas = WMI.ExecQuery("Select * from Win32_ComputerSystem",,48)
For Each sistema in Sistemas
Usuario = sistema.UserName
Dominio = sistema.Domain
MemoriaRam = round((sistema.TotalPhysicalMemory/1000000000)) & "Gb"
Fabricante = sistema.Manufacturer
Modelo = sistema.Model
Next
end function

function ArquitecturaInfo()
Set WshShell =  CreateObject("WScript.Shell")
Set WshProcEnv = WshShell.Environment("Process")
Arquitectura= WshProcEnv("PROCESSOR_ARCHITECTURE") 
end function

ArquitecturaInfo()
SistemaInfo()
DiscoDuroInfo()
SonidoInfo()
videoInfo()
EquipoId()
CDROMInfo()
BoardInfo()
RedInfo()
SOInfo()
ProcesadorInfo()

Texto= "NOMBRE DEL EQUIPO	" & pc & vbcrlf   & "NOMBRE DE USUARIO	" & Usuario & vbcrlf  &   "NOMBRE DE DOMINIO	" & Dominio & vbcrlf & "SISTEMA OPERATIVO	"  & ServicePack & vbcrlf & "DIRECCION  IP	" & Ip & vbcrlf &  "DIRECCION MAC	" & Mac & vbcrlf & "TARJETA MADRE	"  & BoardCompleta & vbcrlf & "PROCESADOR	" & ProNombre  & vbcrlf  & "MEMORIA  RAM	" & MemoriaRam  & vbcrlf & "VIDEO	" & VideoCompleta & vbcrlf&  "SONIDO	" & SonidoNombre & vbcrlf &  "RED	" & NadDescripcion & vbcrlf   & "DISCO DURO	" & DiscoCompleta   & vbcrlf   & "DISPOSITIVO OPTICO	" & CDRNombre & vbcrlf  & "SERIAL EQUIPO	" & EquipoSerial  & vbcrlf  & "ARQUITECTURA	" & Arquitectura & vbcrlf  & "FABRICANTE	" &  Fabricante & vbcrlf  & "MODELO	" & Modelo

msgbox(Texto)
NombArc="Info" & pc & ".xls"
Set ts = fs.OpenTextFile(NombArc,2, True)
ts.WriteLine (Texto)
ts.close