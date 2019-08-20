limite=6
label="AnderToshiba"
Set WshShell = WScript.CreateObject("Wscript.Shell")
Set obj_FSO = CreateObject("Scripting.FileSystemObject") 
Set Archivo = obj_FSO.OpenTextFile("Iconos\Contador.txt", 1)
num = Archivo.ReadLine
Archivo.Close 
Set ArchivoEscribir = obj_FSO.OpenTextFile("Iconos\Contador.txt", 2)
if CInt(num) < limite then
ArchivoEscribir.WriteLine num+1
else
ArchivoEscribir.WriteLine "1"
end if
ArchivoEscribir.Close
Set ArchivoAuto = obj_FSO.OpenTextFile("Autorun.inf", 2)
texto="[AutoRun]" & vbCrLf & "icon=Iconos\icono" & num & ".ico" & vbCrLf & "label=" & label
ArchivoAuto.WriteLine texto
comando="cmd /c label " & label & " | attrib +h +s  autorun.inf"
WshShell.Run comando,0