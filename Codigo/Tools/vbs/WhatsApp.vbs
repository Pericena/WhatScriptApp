' InputBoxes

MsgBox "Espere a que WhatsApp Web este abierto completamente?", 1024 + vbSystemModal, "Preparando"

Contact = InputBox("Con cual contacto quiere hacer SPAM?", "WhatsApp")
Message = InputBox("Cual es el mensaje que quiere enviar?","WhatsApp")
T = InputBox("Cuantas veces quiere mandar el mensaje?","WhatsApp")
If MsgBox("El mensaje se enviado correctamente", 1024 + vbSystemModal, "WhatsApp") = vbOk Then

 
' Abre o WhatsApp
Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("https://web.whatsapp.com/", 1)
 
' Loading Time 
If MsgBox("Espere a que WhatsApp Web este abierto completamente?" & vbNewLine & vbNewLine & "Escoja SI para continuar o NO para cancelar el ataque", vbYesNo + vbQuestion + vbSystemModal, "WhatsApp") = vbYes Then
 
' Editado por Limontec.com para evitar bloqueio do WhatsApp
MsgBox "Espere a que WhatsApp Web este abierto completamente?", 1024 + vbSystemModal, "Preparando"


 
' Abre o WhatsApp Search Bar
WScript.Sleep 50
WshShell.SendKeys "{TAB}"
 
 

' Busca pelo contato
WScript.Sleep 50
WshShell.SendKeys Contact
WScript.Sleep 5000
WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000
 
' Loop das mensagens
For i = 0 to T
WScript.Sleep 5
WshShell.SendKeys Message
WScript.Sleep 5
WshShell.SendKeys "{ENTER}"
Next
 
' Fim do Script
WScript.Sleep 3000
MsgBox "Spam-WhatsApp no " + Contact + "Se completo exitosamente", 1024 + vbSystemModal, "Spam-WhatsApp"

'voz
Dim speaks,speech
speaks="El mensaje se enviado correctamente"
set speech=CreateObject("sapi.spvoice")
speech.Speak speaks 
 
' Mensagem de Script cancelado
Else
MsgBox "El processo fue cancelado Spam-WhatsApp", vbSystemModal, "Spam-WhatsApp Cancelado"
End If
Else
End If