<%
on error resume next
function SessionNet(variabile)
SessionNet=ReadNet(variabile,"session")
end function

function ApplicationNet(variabile)
ApplicationNet=ReadNet(variabile,"application")
end function

function ReadNet(variabile,tipo)
Dim objXMLHTTP, StrURL
StrURL = GetAmministrazionePath() & "library/sharesession/readNet.aspx?type=" & tipo & "&name="&variabile

Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")

objXMLHTTP.Open "GET", StrURL, false
 
StrCookie = Request.ServerVariables("HTTP_COOKIE")
objXMLHTTP.setRequestHeader "COOKIE", StrCookie

objXMLHTTP.Send

ReadNet=CStr(objXMLHTTP.ResponseText)
Set xml = Nothing
end function

%>