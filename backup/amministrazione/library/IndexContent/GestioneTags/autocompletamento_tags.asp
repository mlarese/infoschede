<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../../Tools4admin.asp" -->
<!--#INCLUDE FILE="../../Tools.asp" -->
<% 

dim valueToSearch, inputValue, inputValueList, valueToSet

inputValue = request.Form("tagsInput")
inputValueList = split(inputValue, ",")
valueToSearch = Trim(inputValueList(ubound(inputValueList)))


valueToSet = ""
if ubound(inputValueList) > 0 then
	dim i
	for i = 0 to ubound(inputValueList)-1
		valueToset = valueToSet + trim(inputValueList(i)) + ", "
	next
end if


if valueToSearch<>"" then
	CALL Lightbox_Autocomplete_QUERY("SELECT tag_value FROM tb_contents_tags WHERE tag_value LIKE '" & ParseSql(valueToSearch, adChar) & "%'", valueToSet)
end if


 %>