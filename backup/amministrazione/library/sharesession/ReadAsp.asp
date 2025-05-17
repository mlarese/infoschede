<%
dim buffer,var,vvar
dim separatore

separatore = "#"
on error resume next
'if Request.ServerVariables("REMOTE_ADDR")<>"127.0.0.1" then response.end
tipo=request("type")
name=request("name")
if tipo="session" then
    if name = "" then
		buffer=""
		for each var in session.Contents
			vvar = session(var)
			if instr(1,TypeName(vvar),"()",vbTextCompare)<=0 then
				if buffer<>"" then
					buffer = buffer & separatore
				end if
				buffer = buffer & var&":"&vvar
			end if
		next
		Response.write buffer
	else
	Response.Write (session(name))
    end if
end if
if tipo="application" then
	if name = "" then
		buffer=""
		for each var in Application.Contents
			vvar = Application(var)
			if instr(1,TypeName(vvar),"()",vbTextCompare)<=0 then
				if buffer<>"" then
					buffer = buffer & separatore
				end if
				buffer = buffer & var&":"&vvar
			end if
		next
		Response.write buffer
	else
		Response.Write (Application(name))
	end if
    end if
Response.End()
%>