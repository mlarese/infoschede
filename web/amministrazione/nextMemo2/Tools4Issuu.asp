<!--#INCLUDE FILE="../library/ClassCryptography.asp" -->
<%
 '***********************************************************************
 '	costanti ed impostazioni per Issuu
 '***********************************************************************
  
 dim ISSUU_url
 ISSUU_url = "http://api.issuu.com/1_0"
 
 dim ISSUU_access
 ISSUU_access = "public"
 
dim ISSUU_action
ISSUU_action = "issuu.document.url_upload"

dim ISSUU_apiKey 
ISSUU_apiKey = "esvymoekv2daimmvmv2jeyuqpc4hp691"

dim ISSUU_secretKey 
ISSUU_secretKey = "e1d4mgd6p1njunhxkm0cezc4s9jp39xb"

dim ISSUU_category
ISSUU_category = "014000"

dim ISSUU_type
ISSUU_type = "003000"

dim ISSUU_commentsAllowed
ISSUU_commentsAllowed = "true"

dim ISSUU_downloadable
ISSUU_downloadable = "true"

dim ISSUU_ratingsAllowed
ISSUU_ratingsAllowed = "false"

dim ISSUU_tags
ISSUU_tags = "agenzia rallo, viaggi, catalogo viaggi, viaggi online"
if cString(Application("ISSUU_tags")) <> "" then
	ISSUU_tags = Application("ISSUU_tags")
end if

dim ISSUU_baseurl
ISSUU_baseurl = "http://issuu.com/nextcatalog/docs/"

dim ISSUU_parameters
ISSUU_parameters = "?mode=window&backgroundColor=%23222222"


'compone le stringhe per il post e la stringa per la firma del post
sub AddIssuuPost(postField, postValue, encode, byref postData, byref signatureData)
	
	if cString(postData)<> "" then
		postData = postData + "&"
	end if
	
	postData = postData + postField + "=" + IIF(Encode, Server.UrlEncode(postValue), postValue)
	signatureData = signatureData + postField + postValue
	
end sub


'aggiunge la stringa di firma calcolata secondo le specifiche issuu
sub AddIssuuSignature(byref postData, signatureData)
	
	dim crypt
	set crypt = new CryptographyManager
	
	postData = postData & "&signature=" & lCase(crypt.md5_of_string(ISSUU_secretKey + signatureData))
	
	set crypt = nothing
	
end sub


%>