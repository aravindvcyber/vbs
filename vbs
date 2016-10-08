Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
ObjXML.async=true
objXML.load "res://msxml3.dll/xml/defaultss.xsl"
If objXML.parseError.errorCode <> 0 Then
   Set myErr = objXML.parseError
   WScript.Echo "You have error " + myErr.reason
Else
   WScript.Echo objXML.parseError.reason
   objXML.save "defaultss.xsl"
End If
