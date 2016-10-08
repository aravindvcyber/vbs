Option Explicit
Const STSADM_PATH ="C:\Program Files\Common Files\Microsoft Shared\web server extensions\12\BIN\stsadm.exe"
Const ROOT_URL = "http://root/"
Const FILE_NAME = "E:\listofsites.xml"
 
Dim objShell, objExec, objXml, objXml2,objXml3, objXml4, objSc,objFso, objFile, objWeb
Dim strResult, strSubResult, strUrl, strCmd, strOwner, strSecOwner, strXML, returnValue, returnValue1, returnValue2, returnValue3 
 
'Retrieves all site collections in XML format.
 
WScript.Echo "Creating shell object and calling root enumsites command"
 
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec(STSADM_PATH & " -o enumsites -url " & ROOT_URL)
 
strResult = objExec.StdOut.ReadAll
 
 
'Load XML in DOM document so it can be processed.
 
WScript.Echo "Loading XML File"
Set objXml = CreateObject("MSXML2.DOMDocument")
Set objXml2 = CreateObject("MSXML2.DOMDocument")
Set objXml3 = CreateObject("MSXML2.DOMDocument")
     Set objXml4 = CreateObject("MSXML2.DOMDocument")

objXml.async=false
returnValue2=objXml.LoadXML(strResult)
WScript.Echo "Traversing the Site Collections…"
 if returnValue2 = true then
 if objXml.DocumentElement.Attributes.GetNamedItem("Count").Text <> "0" Then

 WScript.Echo "Number of SiteCollections Found : '" & objXml.DocumentElement.Attributes.GetNamedItem("Count").Text & "'"

 end if
    end if
 
WScript.Echo "Creating File System Object"
'Create the FileSystemObject and write to file.
Set objFso = CreateObject("Scripting.FileSystemObject")
 
if objFso.FileExists(FILE_NAME) then
    objFso.DeleteFile FILE_NAME, True
    WScript.Echo "Deleting the old files…"
end if
 
set objFile = objFso.CreateTextFile(FILE_NAME, True)
 
'Loop through each site collection and call enumsubwebs to get the child URL's.
 
objFile.WriteLine("<WebApplication>")
For Each objSc in objXml.DocumentElement.ChildNodes
    strUrl = objSc.Attributes.GetNamedItem("Url").Text
    strOwner = objSc.Attributes.GetNamedItem("Owner").Text
        strSecOwner=""
    strSecOwner = objSc.Attributes.GetNamedItem("SecondaryOwner").Text
    strCmd = STSADM_PATH & " -o enumsubwebs -url """ + strUrl + """"
 
    Set objExec = objShell.Exec(strCmd)
    strResult = objExec.StdOut.ReadAll
        objXml.async=false
returnValue3=objXml4.LoadXML(strResult)

 if returnValue3 = true then
 if objXml4.DocumentElement.Attributes.GetNamedItem("Count").Text <> "0" Then

          WScript.Echo "Number of Sub Sites Found : '" & objXml4.DocumentElement.Attributes.GetNamedItem("Count").Text & "'"

 end if
    end if
   
    objFile.WriteLine("<SiteCollection SiteCollectionURL='" & strUrl & "' PrimaryOwner = '" & strOwner & "' SecondaryOwner = '" & strSecOwner & "'>" )
    objFile.WriteLine(strResult)
    WScript.Echo "Traversing the sub Webs…"
    call GetSubSites(strResult)
    objFile.WriteLine("</SiteCollection>")
  
Next
objFile.WriteLine("</WebApplication>")
 
set objFile = nothing
set objFso = nothing
set objXml = nothing
set objXml2 = nothing
set objXml3 = nothing
set objXml4 = nothing
set objExec = nothing
 WScript.Echo "File created Successfully"
notepad FILE_NAME
sub GetSubSites(strResult)
objXml.async=false
returnValue1 = objXml2.LoadXML(strResult)
if returnValue1 = true then
for Each objWeb in objXml2.DocumentElement.ChildNodes
    strCmd = STSADM_PATH & " -o enumsubwebs -url """ + objWeb.text + """"
    Set objExec = objShell.Exec(strCmd)
    strResult = objExec.StdOut.ReadAll
 
    objXml.async=false
    returnValue = objXml3.LoadXML(strResult)
    if returnValue = true then
 if objXml3.DocumentElement.Attributes.GetNamedItem("Count").Text <> "0" Then
 objFile.WriteLine(strResult)
                  WScript.Echo "Number of Inner SubSites Found : '" & objXml3.DocumentElement.Attributes.GetNamedItem("Count").Text & "'"
 WScript.Echo "Traversing the Inner sub Webs…"
 call GetSubSites(strResult)
 end if
    end if

next
end if
end sub
