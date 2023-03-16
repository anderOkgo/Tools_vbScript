'_________________________'
'[Created By AnderInner]  '
'   2012 Version 2.00     '
'-------------------------'


Const ForReading = 1
call rename
Public Function sintx(a)
bit = ForReading
Finpal = ""
longi = Len(a)
For i = ForReading To longi
conte = Mid(a, i, ForReading)
If bit = ForReading Then
Finpal = Finpal + UCase(conte)
Else
Finpal = Finpal + conte
End If
If conte = " " Then
bit = ForReading
Else
bit = 0
End If
Next
sintx=(Replace(Finpal," ",""))
end function

sub rename()
Set exe =  CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set f1 = fso.CreateTextFile("ExeRen.bat", True)
Set ts = fso.OpenTextFile("Lista.inf", ForReading)
y=0
x=0
nom="So"
exe.exec("cmd /c ver>ver.txt | exit")
exte=inputbox("Digite la extension de los archivos","Entrada de datos")
pre=inputbox("Digite prefijo en caso de querer agregarlo")
nexte=inputbox("Digite la nueva extension ¡SOLO! en caso de querer cambiarla","Opcional",exte)
eenum = inputbox("Si desea una enumeracion incremetal digite: 1")
if nexte= "" then
nexte=exte
end if
exe.run("cmd /c if exist "  &  nom  & "." & exte  & " (rename "  &  nom  & "." & exte & " """ &  nom  & " (" & 0 & ")." & nexte & """ )")
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set ts2 = fso2.OpenTextFile("ver.txt",ForReading)
Do While Not ts2.AtEndOfStream
stext=(Replace(ts2.ReadLine," ",""))
If Trim(sText) <> "" Then
fine=mid(stext,ForReading,18)
if fine="MicrosoftWindowsXP" then
else
x=x+1
end if
end if
loop
ts2.close

Do While Not ts.AtEndOfStream
stext=sintx(ts.ReadLine)
If Trim(sText) <> "" Then
cont=cont+ForReading
end if
Loop
ts.Close
cant=len(cont)
For i = ForReading To cant
ceros = ceros + "0"
Next

Set ts = fso.OpenTextFile("Lista.inf", ForReading)
Do While Not ts.AtEndOfStream
stext=sintx(ts.ReadLine)
If Trim(sText) <> "" Then
c=c+ForReading
lar = Len(ceros)
For t = ForReading To lar
If Len(c) = t Then
ceros2 = Mid(ceros, ForReading, lar - t)
Else
End If
Next
msgboxA=(" if exist "  &   """"  &  nom  & " (" & x & ")"    &    "." & exte  & """")
msgboxB=("(rename " &  """"  &  nom  & " (" & x & ")"    &    "." & exte  & """")
if eenum = "1" then
msgboxC=("" &  """" & pre & "" & ceros2 & y+ForReading  & "." & sText & "." & nexte  & """"     & ")")
else
msgboxC=("" &  """" &   sText      &    "." & nexte  & """"     & ")")
end if
f1.WriteLine ("" & msgboxA & " " & msgboxB & " " & msgboxC)
y=y+ForReading
x=x+ForReading
End If
Loop
f1.WriteLine ("exit")
f1.Close
ts.Close

exe.run("cmd /k start notepad exeren.bat" )
Wscript.sleep 900
exe.sendkeys "^e"
Wscript.sleep 5
exe.sendkeys "^c"
Wscript.sleep 5
exe.sendkeys "%{F4}"
Wscript.sleep 2500
fso.DeleteFile "ExeRen.bat"
fso.DeleteFile "ver.txt"
end sub
