'1.Hace copia diferencial de archivos  por extensión(es), de todos los discos del sistema o una ruta especifica.
'2.Hace un Escaneo de los archivos del sistema o una ruta especifica por extensión(es) o todos.

Dim ext, dir

ext     = inputbox("Digite la extension ")
new_ext = inputbox("Digite la nueva extension.")
if ext = "" Then
    msgbox "Cancelado: no digito la extension",48
	WScript.quit
end if

'Crear archivos y carpetas necesarios
Set fso = CreateObject("Scripting.FileSystemObject")

'Lee una ubicación especifica.
dir = inputbox("Digite la direccion del directorio de busqueda")
If dir = "" Then
	msgbox "Cancelado: no digito la ruta de busqueda",48
	WScript.quit
End if

Set IDir = fso.getfolder(dir)
ListDirs(IDir)
msgbox "Terminado",64

Function ListDirs(IFol)
	Set directory = fso.GetFolder (IFol.path) 'obtener objeto file basado en una ruta
	If UCase(directory) <> SYSTEMROOT then 'No entra en el directorio de windows
		NewPath = fso.GetBaseName(directory) & "\"
		For Each fichero IN directory.Files
			AlterName = fso.GetBaseName(NewPath & fichero.Name) & "." & new_ext
			fso.MoveFile fichero.path, IFol & "\" & AlterName
		Next
	end if
	
	'Lee los subdirectorios de la carpeta actual y los los envia a la función ListDirs  
	'On error resume next
	Set SubsIFol = IFol.subfolders
	
	For each SF in SubsIFol
		ListDirs(SF)
	Next
End Function