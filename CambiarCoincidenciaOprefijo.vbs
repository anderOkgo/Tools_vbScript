dire        = inputbox("Digite la direccion del directorio de busqueda")
result      = MsgBox ("Si: Para Aregar prefijo" + vbCrLf + "No: Para Modificar Coincidencia", vbYesNo, "Escoja una opcion")
'busq        = inputbox("Digite la Coincidencia a Buscar de lo contrario dejar vacio")
'pref        = inputbox("Digite el prefijo o reemplazo de Coincidencia")
Set fso     = CreateObject("Scripting.FileSystemObject")
Set IDir    = fso.getfolder(dire)


Set objshell = createobject("wscript.shell")
'strMyPath = objShell.SpecialFolders("MyDocuments")
'Set Archivo = fso.CreateTextFile(strMyPath & "\MiArchivo.vbs", True)
'Archivo.WriteLine "Set fso = CreateObject(" & """" & "Scripting.FileSystemObject" & """" & ")"
Select Case result
Case vbYes
pref = inputbox("Digite el prefijo que desea Agregar")
Case vbNo
busq = inputbox("Digite la Coincidencia a Buscar")
pref = inputbox("Digite el reemplazo de la Coincidencia")
End Select

ListDirs(IDir)

'Set objshell = createobject("wscript.shell")
msgbox "Finalizado",64
'fso.DeleteFile (strMyPath & "\MiArchivo.vbs")

Function ListDirs(IFol)
	sum = ""
	Set directorio = fso.GetFolder (IFol.path)
	For Each fichero IN directorio.Files
		Set file = fso.GetFile(fichero)
		Select Case result
		Case vbYes
		b = pref & fichero.Name
		Case vbNo
		b = replace(fichero.Name, busq, pref)
		End Select
		'msgbox("fso.moveFile  " & """" & directorio & "\" & fichero.Name & """" & "," & """" &  directorio & "\" & b & """" & vbCrLf)
		'fso.moveFile directorio & "\" & fichero.Name,   directorio & "\" & b
		'Archivo.WriteLine ("fso.moveFile  " & """" & directorio & "\" & fichero.Name & """" & "," & """" &  directorio & "\" & b & """" & vbCrLf)
		sum = sum + "fso.moveFile  " & """" & directorio & "\" & fichero.Name & """" & "," & """" &  directorio & "\" & b & """" & vbCrLf
	Next
	'Archivo.Close 
	'objShell.run(strMyPath & "\MiArchivo.vbs")
	Execute sum

	Set SubsIFol = IFol.subfolders
	'On error resume next
	For each SF in SubsIFol
		ListDirs(SF)
	Next
End Function