'1.Hace copia diferencial de archivos  por extensión(es), de todos los discos del sistema o una ruta especifica.
'2.Hace un Escaneo de los archivos del sistema o una ruta especifica por extensión(es) o todos.

Dim ext, ext2, except, sys, dir

escan = MsgBox ("Si: Para Escanear, No: Para Copiar", vbYesNo, "Elija una opcion")
Select Case escan
	Case vbYes
	escan = true
	Case vbNo
	escan = false
End Select

result = MsgBox ("En Todo el Sitema?", vbYesNo, "All sys?")
Select Case result
	Case vbYes
	sys = true
	Case vbNo
	sys = false
End Select

ext = inputbox("Digite la extension o las extensiones (separadas por coma)")
if ext <> "" Then
	ext = Split(ext,",")
Else
	ext2 = "AllExt"
End if

'Crear archivos y carpetas necesarios
Set fso = CreateObject("Scripting.FileSystemObject")
If not fso.FolderExists("sandbox") Then fso.CreateFolder "sandbox" End If
if not fso.FileExists("sandbox\Auditoria.txt") then fso.createtextfile("sandbox\Auditoria.txt") End if
Set ts  = fso.OpenTextFile("sandbox\Auditoria.txt",8, True)

if sys = true then
	'Lee todas las unidades de disco.
	except = inputbox("Digite unidad de disco a omitir ej: c ") 'Unidades que se va a omitir en la busqueda
	Set discos = fso.drives
	ts.WriteLine ("---> Inicio	" & now)
	ts.WriteLine ("Accion	Ruta Carpeta	 Nombre Archivo 	 Extension 	 Ruta Completa")
	For each disk in discos
		On error resume next
		dir = disk.driveletter 
		if dir <> UCase(except) then
			Set IDir  = fso.getfolder(dir & ":\")
			ListDirs(IDir)
		end if
	Next
else
	'Lee una ubicación especifica.
	dir = inputbox("Digite la direccion del directorio de busqueda")
	If dir = "" Then
		msgbox "Cancelado: no digito la ruta de busqueda",48
		WScript.quit
	End if

	Set IDir = fso.getfolder(dir)
	ts.WriteLine ("----> Inicio	" & now)
	ts.WriteLine ("Accion	Ruta Carpeta	 Nombre Archivo 	 Extension 	 Ruta Completa")
	ListDirs(IDir)
end if

ts.WriteLine ("----> Final	" & now)
ts.close
msgbox "Terminado",64

Function ListDirs(IFol)
	Set WshShell  = CreateObject("WScript.Shell")
	SYSTEMROOT    = WshShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
	Set directory = fso.GetFolder (IFol.path) 'Set file = fso.GetFile(fichero) 'obtener objeto file basado en una ruta
	If UCase(directory) <> SYSTEMROOT then 'No entra en el directorio de windows
		NewPath = "sandbox\" & fso.GetBaseName(directory) & "\"
		For Each fichero IN directory.Files
			If escan = true Then
				If ext2 = "AllExt" Then
					ts.WriteLine ("Archivo	" & directory.path & "	" & fso.GetBaseName(fichero.Name) & "	" & fso.GetExtensionName(fichero.Name))
				Else
					For i = 0 to UBound(ext)
						If UCase(fso.GetExtensionName(fichero.Name)) = UCase(ext(i)) Then
							ts.WriteLine ("Archivo	" & directory.path & "	" & fso.GetBaseName(fichero.Name) & "	" & fso.GetExtensionName(fichero.Name))
						End if
					Next
				End if
			Else
				If ext2 = "AllExt" Then
					msgbox "Cancelado: Seria Interminable copiar todo el sistema",48
					WScript.quit
				Else
					For i = 0 to UBound(ext)
						if UCase(fso.GetExtensionName(fichero.Name)) = UCase(ext(i)) Then
							If not fso.FolderExists(NewPath) Then fso.CreateFolder NewPath End If

							If fso.FileExists(NewPath & fichero.Name) Then
							  	If (fso.GetFile(NewPath & fichero.Name).DateLastModified <> fichero.DateLastModified) Then
								    AlterName = fso.GetBaseName(NewPath & fichero.Name) & "_" & Replace(Replace(FormatDateTime(fichero.DateLastModified,0),"/","-"), ":","-") & fso.GetExtensionName(fichero.Name)
								    fso.CopyFile fichero.path, NewPath  & "\" & AlterName, true 
								    ts.WriteLine ("Copiado	" & directory.path & "	" & fso.GetBaseName(fichero.Name) & "	" & fso.GetExtensionName(fichero.Name))
								End If
							Else
							    fso.CopyFile fichero.path, NewPath & fichero.Name, true 
								ts.WriteLine ("Copiado	" & directory.path & "	" & fso.GetBaseName(fichero.Name) & "	" & fso.GetExtensionName(fichero.Name))
							End If
						end if
					Next
				End If
			End If
		Next
	end if

	'Lee los subdirectorios de la carpeta actual y los los envia a la función ListDirs  
	On error resume next
	Set SubsIFol = IFol.subfolders
	For each SF in SubsIFol
		ListDirs(SF)
	Next
End Function