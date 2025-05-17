<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<% '*******************************************************************************************************************************
ParentFrameName = "IFrameContattiInterni" %>
<!--#INCLUDE FILE="../library/Intestazione_iframe.asp" -->
<% '*******************************************************************************************************************************


dim conn, rsr, rsi, rsa, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsi = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

'gestione contatti interni
sql = "SELECT * FROM tb_Indirizzario INNER JOIN tb_cnt_lingue ON tb_Indirizzario.lingua = tb_cnt_lingue.lingua_codice " & _
	  " WHERE CntRel=" & request("ID") & " ORDER BY ModoRegistra "
rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText


%>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-right:0px; border-left:0px; border-bottom:0px;" style="width:100% !important;">
<tr>
	<th colspan="4" style="border-bottom:0px;">CONTATTI INTERNI / SEDI ALTERNATIVE</th>
</tr>
<tr>
	<td colspan="4" class="content_right" style="border-top: 1px solid #AAAAAA;">
		<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ContattiInterniNew.asp?CNT=<%= request("ID") %>', 'cntInt', 540, 405, true)">
			NUOVO CONTATTO / SEDE
		</a>
	</td>
</tr>
<tr>
	<td colspan="4">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
			</tr>
			<% if not rsr.eof then %>
				
				<tr>
					<th class="L2">contatto / sede</th>
					<th class="L2">indirizzo</th>
					<th class="l2_center">sede</th>
					<th class="L2" width="3%">&nbsp;</th>
					<th class="L2" width="23%" colspan="2" style="text-align:center;">operazioni</th>
				</tr>
				
				<% while not rsr.eof %>
					<%
					sql = " SELECT tb_ValoriNumeri.*, tb_tipNumeri.nome_tipoNumero FROM tb_tipNumeri INNER JOIN tb_ValoriNumeri " &_
						  " ON tb_tipNumeri.id_tipoNumero = tb_ValoriNumeri.id_TipoNumero " &_
						  " WHERE id_indirizzario=" & cIntero(rsr("IDElencoIndirizzi")) & _
						  " ORDER BY tb_ValoriNumeri.id_TipoNumero, email_default"
					rsi.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
					dim rowspan
					rowspan = 2
					if rsi.eof then
						rowspan = 1
					end if
					%>
					<tr>
						<td class="content_b" style="" title="ruolo / qualifica: <%= rsr("QualificaElencoIndirizzi") %>" rowspan="<%=rowspan%>"><%= ContactFullName(rsr) %></td>
						<td class="content" style="">
							<%= ContactAddress(rsr) %>
							<% if cIntero(rsr("CntSede"))>0 then 
								if ContactAddress(rsr)<>"" then %><br><%end if
								sql = "SELECT * FROM tb_Indirizzario WHERE idElencoIndirizzi = " & rsr("CntSede")
								rsa.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
								if not rsa.eof then%>
									<span class="note">sede: <strong><%= ContactFullName(rsa) %></strong> - <%= ContactAddress(rsa) %></span>
								<% end if
								rsa.close
							end if %>
							&nbsp;
						</td>
						<td class="content_center" style="" rowspan="<%=rowspan%>">
							<input type="checkbox" class="checkbox" disabled <%= chk(rsr("isSocieta")) %> title="<%= IIF(rsr("isSocieta"), "sede alternativa o periferica", "contatto interno") %>">
						</td>
						<td class="content_center" style="padding-top:5px;" rowspan="<%=rowspan%>">
							<% if rsr("lingua")<>"" then %>
								<img src="../grafica/flag_mini_<%= rsr("lingua") %>.jpg" alt="Lingua: <%= rsr("lingua_nome_it") %>">
							<% end if %>
						</td>
						<td class="content_center" style="padding-top:4px; width:10%;" rowspan="<%=rowspan%>">
							<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ContattiInterniMod.asp?CNT=<%= request("ID") %>&ID=<%= rsr("IdElencoIndirizzi") %>', 'cntInt', 950, 400, true)">
								MODIFICA
							</a>
						</td>
						<td class="content_center" style="padding-top:4px; width:10%;" rowspan="<%=rowspan%>">
							<% if cInteger(rsr("LockedByApplication"))>0 then
								sql = "SELECT sito_nome FROM tb_siti WHERE id_sito IN (" & rsr("ApplicationsLocker") & "0 )"%>
								<a class="button_L2_disabled" title="contatto / sede non cancellabile perch&egrave; bloccato dalle applicazioni: <%= GetValueList(conn, rsa, sql) %>.">
									CANCELLA
								</a>
							<% else %>
								<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('CONTATTI','<%= rsr("IDElencoIndirizzi") %>');">
									CANCELLA
								</a>
							<% end if %>
						</td>
					</tr>
					<%
					if not rsi.eof then
						%>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="1" style="width:100%;">
									<%
									while not rsi.eof 
										%>
										<tr>
											<td class="label" style="width:25%;">
												<%=rsi("nome_tipoNumero")%>:&nbsp;
											</td>
											<td class="content">
												<%=rsi("ValoreNumero")%>
											</td>
										</tr>
										<%
										rsi.moveNext
									wend
									%>
								</table>
							</td>
						</tr>
					<% end if 
					rsi.close
					rsr.movenext
				wend 
			end if%>
		</table>
	</td>
</tr>
<tr>
	<td colspan="4" class="content_right">
		&nbsp;
	</td>
</tr>
</table>
</div>
</body>
</html>
<% 
rsr.close
conn.close 
set rsr = nothing
set rsi = nothing
set conn = nothing

%>
