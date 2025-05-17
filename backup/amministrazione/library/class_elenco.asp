<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
class elenco

'variabili e proprieta' locali 
Public fields_elenco 'array di campi visibili nell'elenco
Public fields_rif_id 'id della principale tabella di riferimento 
Public fields_align 
Public fields_tipo
Public fields_link
Public fields_content
'Public page_intervento
Public page_this
Public connessione
Public quanti_records
Public ColumnWidth
Public Caption
Private qColumns
Public ResetPager


Public sub inizializza(ByVal qfel)
	fields_elenco = Array(qfel)
	fields_align = Array(qfel)
	fields_tipo = Array(qfel)
	fields_link = Array(qfel)
	fields_content = Array(qfel)
	redim preserve fields_elenco(qfel)
	redim preserve fields_align(qfel)
	redim preserve fields_tipo(qfel)
	redim preserve fields_link(qfel)
	redim preserve fields_content(qfel)
	qColumns = qfel
	ColumnWidth = Array(qfel)
	redim preserve ColumnWidth(qfel)
end sub

Public Sub elenca(la_query,filtro)
	dim conn, rs, rs_rif, Pager
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open connessione,"",""
	set rs = server.CreateObject("ADODB.Recordset")
	set rs_rif = server.CreateObject("ADODB.Recordset")
	
		set Pager = new PageNavigator
	if ResetPager then
		Pager.Reset()
	end if
	
	CALL Pager.OpenSmartRecordset(conn, rs, la_query, quanti_records)	
	
	response.write "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"+vbCRLF
	if Pager.recordcount>0 then
		response.write "<tr><td style='border: 1px solid Gray; border-bottom:0px;'><table width='100%' style=""border:0px solid Gray"" cellspacing='1' cellpadding='0' border='0'>"+vbCRLF
		response.write "<tr><td valign=""baseline""><font class='testo11b'>"
		response.write "&nbsp;Trovati n&deg; " & Pager.recordcount & " record "
		if rs.PageSize>0 then
			response.write "in n&deg; " & Pager.PageCount & " pagine"
		end if
		response.write "</font></td></tr>"+vbCRLF
		response.write "</table></td></tr>"+vbCRLF
	end if
	response.write "<tr><td><table width='100%' cellspacing='1' cellpadding='0' border='0' class='elenco'>"+vbCRLF
	response.write "<tr>"+vbCRLF
	for a = 1 to ubound(fields_elenco)
		response.write "<td " 
		if cstr(ColumnWidth(a))<>"" then
			response.write "width='" & cstr(ColumnWidth(a)) & "' "
		end if
		response.write " align='"+fields_align(a)+"' bgcolor='#E6E6E6' style='border-bottom: 1px solid Gray;'><font class='testo10b'>"+fields_elenco(a)+"</font></td>"+vbCRLF
	next
	response.write "</tr>"+vbCRLF
	
	
	if Pager.recordcount > 0 then
		rs.AbsolutePage = Pager.PageNo
		while not rs.eof and rs.AbsolutePage = Pager.PageNo
	  		response.write "<tr bgcolor='#F5F4F4'>"+vbCRLF
		  	for a = 1 to ubound(fields_elenco)
				Select Case fields_tipo(a)
					Case "field"
						response.write "<td align='" & fields_align(a) & "' valign='top' style='padding-left:5px; padding-right:5px;'><font class='testo11n'>" & rs( fields_content(a) ) & "</font></td>"+vbCRLF
					Case "multilinks"
						riferimenti = Split(fields_link(a), ";")
						idtabrif = riferimenti(1)
						nomerif = riferimenti(2)
						pagrif = riferimenti(3)
						if instr(1,riferimenti(0), "<ID>")>0 then
							sql_rif = replace(riferimenti(0), "<ID>", rs(fields_rif_id))
						else
							sql_rif = riferimenti(0) & rs(fields_rif_id)
						end if
						response.write "<td align='" & fields_align(a) & "' valign='top'>"
						rs_rif.open sql_rif, conn, adOpenStatic, adLockOptimistic
						if not rs_rif.eof then
							rs_rif.MoveFirst
							do while not rs_rif.eof%>
								<a href="<%=pagrif%>?ID=<%=rs_rif(idtabrif)%>&proven=<%=page_this%>"  class="links_multi" style="white-space: nowrap;">
									<%=rs_rif(nomerif)%>
									<%rs_rif.MoveNext
									if not rs_rif.eof then %>
										, 
									<%end if %>
								</a>
							<%loop
						else%>
							<font class='testo11n'>&nbsp;</font>
						<%end if
						response.write "</td>"+vbCRLF
						rs_rif.close
					case "checkbox"
						response.write "<td align='" & fields_align(a) & "' bgcolor='#F5F4F4'><input class=""checkbox"" type=""checkbox"" style=""height:13px;"" "
						if rs( fields_content(a) ) then
							response.write " checked "
						end if
						 response.write "disabled></td>"
					Case "link"
						id_rec = rs( fields_content(1))
						nome_rec = replace(rs( fields_content(2)),"'","\'")
						response.write "<td align='" & fields_align(a) & "' valign='top'>"+vbCRLF
						if instr(1, fields_link(a), "href", vbTextCompare) > 0 then
							response.write "<a " & replace(replace(fields_link(a),"<ID>",id_rec),"<NOME>",nome_rec)  & " class='simul_puls_1'>" & fields_content(a) & "</a></td>"+vbCRLF
						else
							response.write "<a href='" & fields_link(a) & "?ID=" & id_rec & "' class='simul_puls_1'>" & fields_content(a) & "</a></td>"+vbCRLF
						end if
				end select
			next
			response.write "</tr>"+vbCRLF
	  		rs.MoveNext
		wend
	else
		response.write "<tr><td colspan=" & ubound(fields_content) & ">" &_
		 				"<font class='testo11b'>&nbsp;Nessun record trovato.</font>" & _
		 				"</td></tr>"
	end if
	response.write "</table>"+vbCRLF
	response.write "</td></tr>"+vbCRLF
	if rs.PageSize>1 then
		response.write "<tr><td style=""font:11px Arial #000; padding-top:3px;"">"
		CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "testo11n", "simul_puls_1", "puls_gray")
		response.write "</td></tr>"+vbCRLF
	end if
	response.write "<tr><td><font style=""font:40px arial;"">&nbsp;</font></td></tr>"
	response.write "</table>"+vbCRLF
	rs.close
	conn.close
	conn = null
end sub

end class
 %>