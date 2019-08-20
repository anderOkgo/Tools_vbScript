cortar=inputbox("Digite la cantidad caracters a quitar")
Const ForReading = 1
Set fso = CreateObject("Scripting.FileSystemObject")
Set f1 = fso.CreateTextFile("Lista.inf", True)
Set ts = fso.OpenTextFile("Lista2.inf", ForReading)
Set objshell = createobject("Wscript.shell")
msg = objshell.popup("¿ La numeracion es Simetrica ?",10,"Mensaje Popup",36)
Do While Not ts.AtEndOfStream
if msg=6 then
else
i=i+1
end if
cut=len(i)
sText = replace(ts.ReadLine,"	","")
tot=len(stext)
If Trim(sText) <> "" Then
finPal=mid(stext,cut+cortar,tot)
f1.WriteLine (finpal)
End If
loop
f1.Close

